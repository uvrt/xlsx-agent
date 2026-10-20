"""X5 (P1-6): the formula-aware cell compare, ``Workbook.compare(a, b, cells=True)`` and
``xlsx_compare``'s ``cells``: every cell by address, formulas by their normalised text, a
shared formula's member as its expanded text, bounded, and the changes outside a write
policy.

Every workbook is the made-up corner-shop template (``tests/shop.py``), its parts changed by
hand where a test needs a workbook no write of ours makes (a formula changed while its
cached value stays).
"""

from __future__ import annotations

import io
import json
import logging
import zipfile

import pytest

from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit.report import MAX_RANGES
from xlsx_agent.formula import normalise
from xlsx_agent.integrity import KINDS, STRUCTURE_KINDS
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS, policy

from shop import FIRST_INPUT, ITEMS, Shop

STOCK = "xl/worksheets/sheet1.xml"
TOTALS = "xl/worksheets/sheet3.xml"
LAST = FIRST_INPUT + len(ITEMS) - 1


def _zip(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return buffer.getvalue()


def _changed(shop: Shop | None = None, **edits: list[tuple[str, str]]) -> bytes:
    """The shop with each part's text replaced as given (``stock=[(old, new)]``)."""
    parts = (shop or Shop()).parts()
    names = {"stock": STOCK, "totals": TOTALS, "workbook": "xl/workbook.xml"}
    for key, pairs in edits.items():
        text = parts[names[key]].decode()
        for old, new in pairs:
            assert old in text, old
            text = text.replace(old, new)
        parts[names[key]] = text.encode()
    return _zip(parts)


def _cells(a: bytes, b: bytes, **options):
    return Workbook.compare(Workbook.open(a), Workbook.open(b), cells=True, **options).cells


def _kinds(diff) -> dict[str, str]:
    return {edit.address: edit.kind for edit in diff.edits}


# -- what is the same formula ------------------------------------------------------------------


@pytest.mark.parametrize("one, two", [
    ("SUM(Stock!F10:F14)", "=sum( stock!f10:f14 )"),       # case of names, refs; spaces
    ("VLOOKUP(C10,Lists!$A$2:$B$5,2,FALSE)", "vlookup(c10, lists!$a$2:$b$5, 2, false)"),
    ("A1:B5 B2:C3", "a1:b5   b2:c3"),                       # intersection kept, as one
    ("#N/A", "#n/a"),
    ("Markup*2", "MARKUP * 2"),
    ("1E3+1", "1e3 + 1"),
    ("'Q3 sales'!A1", "'q3 SALES'!a1"),
    ("SUM(A1,\nB1)", "SUM(A1,B1)"),
])
def test_case_and_space_the_grammar_gives_no_meaning_are_the_same_formula(one, two):
    assert normalise(one) == normalise(two)


@pytest.mark.parametrize("one, two", [
    ('"P-"&A1', '"p-"&A1'),                 # a string literal's case is its value
    ('"a b"&A1', '"a  b"&A1'),              # and so are its spaces
    ("A1:B5 B2:C3", "A1:B5,B2:C3"),         # intersection is not union
    ("A1:B5 B2:C3", "A1:B5B2:C3"),
    ("Stock[Price]", "Stock[price]"),       # structured references kept as written
    ("D10*E10", "D11*E11"),
])
def test_what_the_grammar_gives_meaning_is_another_formula(one, two):
    assert normalise(one) != normalise(two)


def test_text_the_lexer_cannot_read_is_compared_exactly():
    assert normalise('"open') == '"open'
    assert normalise('="open') == '"open'


# -- the kinds ---------------------------------------------------------------------------------


def test_the_same_workbook_is_the_same_cell_by_cell():
    data = Shop().to_bytes()
    diff = Workbook.compare(Workbook.open(data), Workbook.open(data), cells=True)
    assert diff.same and diff.cells.same
    assert diff.cells.sheets == ("Stock", "Lists", "Totals")
    assert diff.cells.as_dict() == {"sheets": 3, "counts": {}}
    assert diff.tables == ()                # tables only when named


def test_a_formula_changed_while_its_cached_value_stays_is_found():
    lookup = f"VLOOKUP(C{FIRST_INPUT},Lists!$A$2:$B$5,2,FALSE)"
    b = _changed(stock=[(lookup, lookup.replace("FALSE", "TRUE"))])
    diff = _cells(Shop().to_bytes(), b)
    assert _kinds(diff) == {f"Stock!I{FIRST_INPUT}": "formula_changed"}
    # the tables compare, by cached value, sees nothing there
    plain = Workbook.compare(Workbook.open(Shop().to_bytes()), Workbook.open(b))
    assert plain.same


def test_a_formula_turned_into_its_own_value_is_found():
    # F12 keeps its cached value; only its <f> goes: what a paste of values does
    b = _changed(stock=[('<f t="shared" si="0"/><v>72</v>', '<v>72</v>')])
    assert _kinds(_cells(Shop().to_bytes(), b)) == {"Stock!F12": "formula_to_literal"}


def test_each_kind_by_address():
    a = Shop().to_bytes()
    b = _changed(
        stock=[(f'<c r="D{FIRST_INPUT}" s="5"><v>120</v></c>',          # value_changed
                f'<c r="D{FIRST_INPUT}" s="5"><v>121</v></c>'),
               (f'<c r="E{FIRST_INPUT + 1}" s="3"><v>2.5</v></c>',      # literal_to_formula
                f'<c r="E{FIRST_INPUT + 1}" s="3"><f>2+0.5</f><v>2.5</v></c>'),
               (f'<c r="B4" s="5"/>', '<c r="B4" s="5"><f>A4</f></c>')],  # formula_added
        totals=[('<f>COUNTIF(Stock!G10:G14,FALSE)</f><v>2</v>', ''),      # formula_removed
                ('<f>INDEX(TaxRates,2)</f><v>0.21</v>',                    # cache_changed
                 '<f>index( TaxRates , 2 )</f><v>0.2</v>')])
    diff = _cells(a, b)
    assert _kinds(diff) == {
        "Stock!B4": "formula_added", f"Stock!D{FIRST_INPUT}": "value_changed",
        f"Stock!E{FIRST_INPUT + 1}": "literal_to_formula", "Totals!B2": "formula_removed",
        "Totals!B3": "cache_changed"}
    assert set(diff.counts()) <= set(KINDS)
    assert list(diff.counts()) == [k for k in KINDS if k in diff.counts()]


def test_a_value_is_compared_as_stored_text_numbers_as_numbers():
    shop = Shop(shared=False)                # inline strings
    a, b = Shop().to_bytes(), shop.to_bytes()
    # a shared string and an inline string of one text are one value
    assert _cells(a, b).same
    b = _changed(stock=[(f'<c r="D{FIRST_INPUT}" s="5"><v>120</v></c>',
                         f'<c r="D{FIRST_INPUT}" s="5"><v>120.0</v></c>')])
    assert _cells(a, b).same
    b = _changed(stock=[(f'<c r="D{FIRST_INPUT}" s="5"><v>120</v></c>',
                         f'<c r="D{FIRST_INPUT}" s="5" t="str"><v>120</v></c>')])
    assert _kinds(_cells(a, b)) == {f"Stock!D{FIRST_INPUT}": "value_changed"}


def test_a_shared_member_is_compared_as_its_expanded_text():
    # the group written out as one formula a cell: the same formulas, no change
    a = Shop().to_bytes()
    pairs = [(f'<f t="shared" ref="F{FIRST_INPUT}:F{LAST}" si="0">D{FIRST_INPUT}*E{FIRST_INPUT}'
              f'</f>', f'<f>D{FIRST_INPUT}*E{FIRST_INPUT}</f>')]
    stock = Shop().stock()
    for n in range(FIRST_INPUT + 1, LAST + 1):
        start = stock.index(f'<c r="F{n}"')
        old = stock[start:stock.index("</c>", start)]
        pairs.append((old, old.replace('<f t="shared" si="0"/>', f"<f>d{n} * e{n}</f>")))
    assert _cells(a, _changed(stock=pairs)).same
    # a member moved by hand to another row's line is a change
    pairs[2] = (pairs[2][0], pairs[2][0].replace('<f t="shared" si="0"/>',
                                                 f"<f>D{FIRST_INPUT}*E{FIRST_INPUT}</f>"))
    assert _kinds(_cells(a, _changed(stock=pairs[2:3]))) == {
        f"Stock!F{FIRST_INPUT + 2}": "formula_changed"}


def test_a_shared_master_written_over_changes_only_that_cell():
    data = Shop().to_bytes()
    after = Workbook.open(data)
    after.set_cells({f"Stock!F{FIRST_INPUT}": 42}, replace_formulas=True)
    diff = Workbook.compare(Workbook.open(data), after, cells=True)
    # the group was handed to F11 with its formula moved: every other member is unchanged
    assert _kinds(diff.cells) == {f"Stock!F{FIRST_INPUT}": "formula_to_literal"}


def test_an_array_formula_is_compared_with_its_block():
    a = Shop(array=True).to_bytes()
    block = (f'<c r="C5"><f t="array" ref="C5:D5">Stock!D{FIRST_INPUT}:E{FIRST_INPUT}</f>'
             f'<v>120</v></c><c r="D5"><v>0.35</v></c>')
    assert _cells(a, _changed(Shop(array=True), totals=[(block, block.replace(
        "<v>0.35</v>", "<v>0.4</v>"))])).edits[0].kind == "cache_changed"
    assert _kinds(_cells(a, _changed(Shop(array=True), totals=[(block, block.replace(
        'ref="C5:D5"', 'ref="C5:C5"'))]))) == {"Totals!C5": "formula_changed",
                                               "Totals!D5": "formula_to_literal"}
    literal = '<c r="C5"><v>120</v></c><c r="D5"><v>0.35</v></c>'
    assert _kinds(_cells(a, _changed(Shop(array=True), totals=[(block, literal)]))) == {
        "Totals!C5": "formula_to_literal", "Totals!D5": "formula_to_literal"}


def test_a_sheet_in_one_workbook_only_is_a_sheet_change_its_cells_not_compared():
    b = _changed(workbook=[('<sheet name="Totals" sheetId="3" r:id="rId3"/>',
                            '<sheet name="Totals" sheetId="9" r:id="rId3"/>')])
    diff = Workbook.compare(Workbook.open(Shop().to_bytes()), Workbook.open(b), cells=True)
    assert [s.change for s in diff.sheets] == ["removed", "added"]
    assert diff.cells.sheets == ("Stock", "Lists") and diff.cells.same


def test_what_is_withheld_from_the_model_is_not_compared():
    a, b = Workbook.open(Shop().to_bytes()), Workbook.open(Shop().to_bytes())
    b.set_cells({"Lists!A2": "Office"})
    for wb in (a, b):
        wb.withhold_sheets(["Lists"])
    diff = Workbook.compare(a, b, cells=True)
    assert diff.cells.sheets == ("Stock", "Totals") and diff.cells.same


# -- formulas only when asked ----------------------------------------------------------------


def test_formula_text_only_when_asked_and_never_a_redacted_one():
    lookup = f"VLOOKUP(C{FIRST_INPUT},Lists!$A$2:$B$5,2,FALSE)"
    b = _changed(stock=[(lookup, lookup.replace("FALSE", "TRUE"))],
                 totals=[('<f>INDEX(TaxRates,2)</f>', '<f>INDEX(TaxRates,3)</f>')])
    (one, two) = _cells(Shop().to_bytes(), b).edits
    assert (one.a, one.b, two.a, two.b) == (None, None, None, None)
    (one, two) = _cells(Shop().to_bytes(), b, formula_text=True).edits
    assert (one.a, one.b) == (lookup, lookup.replace("FALSE", "TRUE"))
    assert two.as_dict() == {"at": "Totals!B3", "kind": "formula_changed",
                             "a": "INDEX(TaxRates,2)", "b": "INDEX(TaxRates,3)"}
    old, new = Workbook.open(Shop().to_bytes()), Workbook.open(b)
    for wb in (old, new):
        wb.redact([{"redact": ["TaxRates"]}])
    (one, two) = Workbook.compare(old, new, cells=True, formula_text=True).cells.edits
    assert one.a is not None and (two.a, two.b, two.withheld) == (None, None, True)
    with pytest.raises(ValueError, match="cells=True"):
        Workbook.compare(old, new, formula_text=True)


def test_a_redaction_never_makes_two_texts_one():
    a = Workbook.open(Shop().to_bytes())
    b = Workbook.open(Shop().to_bytes())
    b.set_cells({"Stock!B4": "Ana Corner"})
    a.set_cells({"Stock!B4": "Ana corner"})
    for wb in (a, b):
        wb.redact([{"redact": ["corner"]}])
    assert _kinds(Workbook.compare(a, b, cells=True).cells) == {"Stock!B4": "value_changed"}


# -- bounded ------------------------------------------------------------------------------------


def test_the_result_is_counts_and_capped_ranges_never_content():
    data = Shop().to_bytes()
    after = Workbook.open(data)
    # 120 single cells, none beside another: more ranges than a report lists
    after.set_cells({f"Stock!{column}{row}": "Spare part"
                     for row in range(20, 140, 2) for column in ("K", "M")})
    diff = Workbook.compare(Workbook.open(data), after, cells=True)
    shaped = diff.cells.as_dict()
    assert shaped["counts"] == {"value_changed": 120}
    listed = sum(len(r) for r in shaped["cells"]["value_changed"].values())
    assert listed == MAX_RANGES and shaped["cut"] == {"value_changed": 120 - MAX_RANGES}
    assert "Spare" not in json.dumps(shaped)


def test_nothing_of_a_cell_or_formula_reaches_the_log(caplog):
    caplog.set_level(logging.DEBUG)
    b = _changed(totals=[('<f>INDEX(TaxRates,2)</f>', '<f>INDEX(TaxRates,3)</f>')])
    _cells(Shop().to_bytes(), b, formula_text=True)
    text = " ".join(record.getMessage() for record in caplog.records)
    assert "TaxRates" not in text and "Pencil" not in text


# -- against a write policy --------------------------------------------------------------------


def _edited(**options) -> tuple[Workbook, Workbook]:
    data = Shop().to_bytes()
    after = Workbook.open(data)
    after.set_cells({"Stock!B4": "Ana", f"Stock!D{FIRST_INPUT}": 5,
                     f"Stock!F{FIRST_INPUT + 1}": 1.75,
                     "Stock!K2": {"formula": "D10*2"}}, replace_formulas=True)
    return Workbook.open(data, **options), after


def test_the_changes_outside_the_policy_are_listed_by_address():
    before, after = _edited()
    cells = Workbook.compare(before, after, cells=True, policy={
        "cells": ["Stock!B4", f"Stock!D{FIRST_INPUT}:E{LAST}"]}).cells
    # in row order, then column order
    assert [e.address for e in cells.outside] == ["Stock!K2", f"Stock!F{FIRST_INPUT + 1}"]
    assert cells.as_dict()["outside"] == {
        "count": 2, "counts": {"formula_to_literal": 1, "formula_added": 1},
        "cells": {"Stock": ["K2", f"F{FIRST_INPUT + 1}"]}}
    # the formulas that read the cells written are recalculated, never written: not outside
    assert {e.kind for e in cells.edits} >= {"value_changed"}


def test_a_policy_s_formula_rules_make_formula_changes_outside_it():
    before, after = _edited()
    everywhere = {"cells": ["Stock!", "Totals!"]}
    assert Workbook.compare(before, after, cells=True, policy=everywhere).cells.outside == ()
    no_formulas = Workbook.compare(before, after, cells=True,
                                   policy={**everywhere, "formulas": False}).cells
    assert [e.address for e in no_formulas.outside] == ["Stock!K2"]
    kept = Workbook.compare(before, after, cells=True,
                            policy={**everywhere, "replace_formulas": False}).cells
    assert [e.address for e in kept.outside] == [f"Stock!F{FIRST_INPUT + 1}"]


def test_the_older_workbook_s_own_policy_is_asked_by_default():
    before, after = _edited(policy={"cells": ["Stock!B4"]})
    cells = Workbook.compare(before, after, cells=True).cells
    assert [e.address for e in cells.outside] == ["Stock!K2", f"Stock!D{FIRST_INPUT}",
                                                  f"Stock!F{FIRST_INPUT + 1}"]
    plain, other = _edited()
    assert Workbook.compare(plain, other, cells=True).cells.outside is None


def test_a_recalculated_value_is_never_outside_and_a_sheet_change_always_is():
    a = Shop().to_bytes()
    b = _changed(totals=[('<f>INDEX(TaxRates,2)</f><v>0.21</v>',
                          '<f>INDEX(TaxRates,2)</f><v>0.2</v>')],
                 workbook=[('name="Totals"', 'name="Sums"')])
    cells = _cells(a, b, policy={"cells": ["Stock!B4"]})
    assert [e.kind for e in cells.edits] == ["cache_changed"] and cells.outside == ()
    assert cells.sheets_outside == 1 and cells.as_dict()["outside"] == {"count": 0,
                                                                         "sheets": 1}


# -- structure (X6) ---------------------------------------------------------------------------

_QUANTITY_RULE = (f'<dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" '
                  f'showErrorMessage="1" errorTitle="Quantity" error="Whole numbers from 0" '
                  f'sqref="D{FIRST_INPUT}:D200"><formula1>0</formula1></dataValidation>')
_DATE_RULE = (f'<dataValidation type="date" allowBlank="1" showInputMessage="1" '
              f'sqref="H{FIRST_INPUT}:H200"><formula1>36526</formula1></dataValidation>')
_LOW = '<cfRule type="cellIs" dxfId="0" priority="2" operator="lessThan"><formula>5</formula>'


def _structure(a: bytes, b: bytes, **options) -> dict[str, str]:
    return {change.address: change.kind for change in _cells(a, b, **options).structure}


def test_names_validations_merges_and_conditional_formats_are_compared_apart():
    b = _changed(
        workbook=[('<definedName name="Markup">1.4</definedName>',
                   '<definedName name="Markup">1.5</definedName>'
                   '<definedName name="Shelf" localSheetId="0">Stock!$B$4</definedName>'),
                  ('<definedName name="TaxRates">Lists!$B$2:$B$5</definedName>', '')],
        stock=[('<mergeCell ref="A2:I2"/>', '<mergeCell ref="A4:B4"/>'),
               (_QUANTITY_RULE, _QUANTITY_RULE.replace("greaterThanOrEqual", "greaterThan")),
               (_DATE_RULE, ''),
               ('</dataValidations>', '<dataValidation type="list" sqref="J10:J20 G10:G20">'
                '<formula1>"Yes,No"</formula1></dataValidation></dataValidations>'),
               (_LOW, _LOW.replace("<formula>5</formula>", "<formula>6</formula>")),
               ('</conditionalFormatting><dataValidations',
                '</conditionalFormatting><conditionalFormatting sqref="E10:E14"><cfRule '
                'type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/>'
                '<color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>'
                '</conditionalFormatting><dataValidations')])
    diff = Workbook.compare(Workbook.open(Shop().to_bytes()), Workbook.open(b), cells=True)
    assert diff.cells.edits == ()                      # no cell changed
    assert _structure(Shop().to_bytes(), b) == {
        "Markup": "name_changed", "Stock!Shelf": "name_added", "TaxRates": "name_removed",
        "Stock!A4:B4": "merge_added", "Stock!A2:I2": "merge_removed",
        f"Stock!D{FIRST_INPUT}:D200": "validation_changed",
        "Stock!G10:G20 J10:J20": "validation_added",
        f"Stock!H{FIRST_INPUT}:H200": "validation_removed",
        f"Stock!D{FIRST_INPUT}:D{LAST}": "conditional_format_changed",
        "Stock!E10:E14": "conditional_format_added"}
    assert diff.totals() == {"name_added": 1, "name_removed": 1, "name_changed": 1,
                             "validation_added": 1, "validation_removed": 1,
                             "validation_changed": 1, "merge_added": 1, "merge_removed": 1,
                             "conditional_format_added": 1, "conditional_format_changed": 1}
    assert set(diff.cells.structure_counts()) <= set(STRUCTURE_KINDS) and not diff.same


def test_what_carries_no_meaning_is_no_structural_change():
    b = _changed(
        workbook=[('<definedName name="Categories">Lists!$A$2:$A$5</definedName>',
                   '<definedName name="categories">lists!$a$2:$a$5</definedName>')],
        stock=[(_DATE_RULE, ''),          # the same rule, stored as x14 instead
               ('</x14:dataValidations>', '<x14:dataValidation type="date" allowBlank="1">'
                '<x14:formula1><xm:f>36526</xm:f></x14:formula1>'
                f'<xm:sqref>H{FIRST_INPUT}:H200</xm:sqref></x14:dataValidation>'
                '</x14:dataValidations>'),
               ('error="Whole numbers from 0"', 'error="From 0 up"'),     # a message
               ('<mergeCell ref="A1:I1"/><mergeCell ref="A2:I2"/>',
                '<mergeCell ref="A2:I2"/><mergeCell ref="a1:i1"/>'),       # order, case
               (_LOW, _LOW.replace('dxfId="0" priority="2"', 'dxfId="3" priority="7"'))])
    assert _structure(Shop().to_bytes(), b) == {}


def test_structural_formulas_only_when_asked_and_never_a_redacted_one():
    b = _changed(workbook=[('<definedName name="Markup">1.4</definedName>',
                            '<definedName name="Markup">TaxRates*2</definedName>')],
                 stock=[(_LOW, _LOW.replace("<formula>5</formula>", "<formula>6</formula>"))])
    plain = _cells(Shop().to_bytes(), b).structure
    assert [(c.a, c.b) for c in plain] == [((), ()), ((), ())]
    assert plain[0].as_dict() == {"kind": "name_changed", "at": "Markup"}
    named, rule = _cells(Shop().to_bytes(), b, formula_text=True).structure
    assert (named.a, named.b, rule.a, rule.b) == (("1.4",), ("TaxRates*2",), ("5",), ("6",))
    old, new = Workbook.open(Shop().to_bytes()), Workbook.open(b)
    for wb in (old, new):
        wb.redact([{"redact": ["TaxRates"]}])
    named, rule = Workbook.compare(old, new, cells=True, formula_text=True).cells.structure
    assert (named.a, named.b, named.withheld) == (("1.4",), (), True)   # its own side
    assert rule.b == ("6",)


def test_every_structural_change_is_outside_a_policy_and_the_list_is_capped():
    merges = "".join(f'<mergeCell ref="K{n}:L{n}"/>' for n in range(20, 20 + MAX_RANGES + 10))
    b = _changed(stock=[('<mergeCell ref="A2:I2"/>', '<mergeCell ref="A2:I2"/>' + merges)])
    cells = _cells(Shop().to_bytes(), b, policy={"cells": ["Stock!"]})
    shown = cells.as_dict()
    assert shown["structure"]["counts"] == {"merge_added": MAX_RANGES + 10}
    assert len(shown["structure"]["changes"]) == MAX_RANGES
    assert shown["structure"]["cut"] == 10
    assert shown["structure"]["changes"][0] == {"kind": "merge_added", "at": "Stock!K20:L20"}
    assert shown["outside"] == {"count": 0, "structure": MAX_RANGES + 10}


# -- xlsx_compare -----------------------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, *books, given=None):
    session = box.session()
    if given is not None:
        policy(session, given)
    for index, data in enumerate(books):
        blob = session.add_blob(data, name=f"book{index}.xlsx")
        assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_compare_cells_gives_counts_ranges_and_what_the_session_policy_refuses(box):
    lookup = f"VLOOKUP(C{FIRST_INPUT},Lists!$A$2:$B$5,2,FALSE)"
    b = _changed(stock=[(lookup, lookup.replace("FALSE", "TRUE")),
                        ('<f t="shared" si="0"/><v>72</v>', '<v>72</v>')])
    session = _session(box, Shop().to_bytes(), b, given={"cells": ["Stock!B4"]})
    result = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2", "cells": True})
    assert result.ok, result.error
    assert result.data["tables"] == [] and result.data["differences"] == []
    assert result.data["cells"] == {
        "sheets": 3, "counts": {"formula_changed": 1, "formula_to_literal": 1},
        "cells": {"formula_changed": {"Stock": [f"I{FIRST_INPUT}"]},
                  "formula_to_literal": {"Stock": ["F12"]}},
        "outside": {"count": 2, "counts": {"formula_changed": 1, "formula_to_literal": 1},
                    "cells": {"Stock": [f"I{FIRST_INPUT}", "F12"]}}}
    assert result.summary == ("d1 and d2: 0 tables compared, cells of 3 sheets (2 outside "
                              "the write policy); 1 formula changed, 1 formula to literal")
    assert "VLOOKUP" not in result.summary


def test_xlsx_compare_formula_text_pages_the_formulas(box):
    lookup = f"VLOOKUP(C{FIRST_INPUT},Lists!$A$2:$B$5,2,FALSE)"
    b = _changed(stock=[(lookup, lookup.replace("FALSE", "TRUE"))])
    session = _session(box, Shop().to_bytes(), b)
    result = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2", "cells": True,
                                                    "formula_text": True})
    assert result.ok, result.error
    assert result.data["differences"] == [{"formula": "formula_changed",
                                           "at": f"Stock!I{FIRST_INPUT}", "a": lookup,
                                           "b": lookup.replace("FALSE", "TRUE")}]
    assert "outside" not in result.data["cells"]          # no policy in the session
    refused = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2",
                                                     "formula_text": True})
    assert refused.error.code == "invalid_arguments"
    assert refused.error.field == "formula_text"


def test_xlsx_compare_cells_lists_the_structure_apart_and_pages_its_formulas(box):
    b = _changed(workbook=[('<definedName name="Markup">1.4</definedName>',
                            '<definedName name="Markup">1.5</definedName>')],
                 stock=[('<mergeCell ref="A2:I2"/>', '')])
    session = _session(box, Shop().to_bytes(), b, given={"cells": ["Stock!B4"]})
    result = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2", "cells": True})
    assert result.ok, result.error
    assert result.data["cells"]["structure"] == {
        "counts": {"name_changed": 1, "merge_removed": 1},
        "changes": [{"kind": "name_changed", "at": "Markup"},
                    {"kind": "merge_removed", "at": "Stock!A2:I2"}]}
    assert result.data["cells"]["outside"] == {"count": 0, "structure": 2}
    assert result.summary.startswith("d1 and d2: 0 tables compared, cells of 3 sheets (2 "
                                     "outside the write policy); 1 name changed")
    texts = box.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2", "cells": True,
                                                   "formula_text": True})
    assert texts.data["differences"] == [{"structure": "name_changed", "at": "Markup",
                                          "a": ["1.4"], "b": ["1.5"]}]
