"""X5 (P1-7): data validations inspected -- ``Workbook.validations`` and ``xlsx_read``'s
``validations`` -- standard and ``x14`` rules, list sources resolved through ranges, defined
names and table columns, dynamic sources reported unresolved by address; rules kept byte
for byte through edits; a written value checked against a resolvable rule (a warning).

Every workbook is the made-up corner-shop template (``tests/shop.py``; ``orders=True`` adds
the ``Orders`` sheet of rules of every kind).
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest
from lxml import etree

from ooxml_edit.tools import Toolbox

from xlsx_agent import ValidationRule, Workbook
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from shop import FIRST_INPUT, ORDER_ROWS, Shop

STOCK, ORDERS = "xl/worksheets/sheet1.xml", "xl/worksheets/sheet4.xml"


def _part(data: bytes, name: str) -> bytes:
    return zipfile.ZipFile(io.BytesIO(data)).read(name)


def _zip(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return buffer.getvalue()


def _shown(rules: list[ValidationRule], **options) -> list[dict]:
    return [rule.as_dict(**options) for rule in rules]


def test_the_stock_sheet_s_rules_standard_then_x14():
    rules = Workbook.open(Shop().to_bytes()).validations("Stock")
    assert _shown(rules) == [
        {"sheet": "Stock", "origin": "standard", "type": "whole",
         "operator": "greaterThanOrEqual", "allow_blank": True, "cells": "D10:D200"},
        # no operator written: ECMA-376's default, between
        {"sheet": "Stock", "origin": "standard", "type": "date", "operator": "between",
         "allow_blank": True, "cells": "H10:H200"},
        {"sheet": "Stock", "origin": "x14", "type": "list", "allow_blank": True,
         "cells": "C10:C200", "source": "range", "resolved": True, "range": "Lists!A2:A5",
         "count": 4}]
    assert [r.index for r in rules] == [0, 1, 2]
    assert rules[2].address == "Stock!C10:C200"
    assert rules[2].values == ("Stationery", "Kitchen", "Hardware", "Garden")


def test_every_kind_of_list_source_and_rule():
    wb = Workbook.open(Shop(orders=True, table=True).to_bytes())
    rules = {rule.cells[0]: rule for rule in wb.validations("Orders")}
    last = ORDER_ROWS

    def picked(rule, *names):
        shown = rule.as_dict()
        return {name: shown.get(name) for name in names}

    assert picked(rules[f"B2:B{last}"], "source", "resolved", "name", "range", "count") == {
        "source": "name", "resolved": True, "name": "Categories", "range": "Lists!A2:A5",
        "count": 4}
    assert picked(rules[f"C2:C{last}"], "source", "resolved", "count") == {
        "source": "inline", "resolved": True, "count": 3}
    assert rules[f"C2:C{last}"].values == ("S", "M", "L")
    # a dependent list: picked by another cell's value, never evaluated
    assert picked(rules[f"A2:A{last}"], "source", "resolved", "reason", "uses") == {
        "source": "indirect", "resolved": False, "reason": "indirect", "uses": ["$B2"]}
    # a name whose formula is a dynamic range
    assert picked(rules[f"G2:G{last}"], "source", "resolved", "reason", "name") == {
        "source": "name", "resolved": False, "reason": "offset", "name": "Shelves"}
    assert picked(rules[f"H2:H{last}"], "source", "resolved", "range", "count") == {
        "source": "table", "resolved": True, "range": "Lists!B2:B5", "count": 4}
    assert picked(rules[f"J2:J{last}"], "source", "resolved", "reason") == {
        "source": "name", "resolved": False, "reason": "unknown_name"}
    assert picked(rules[f"E2:E{last}"], "type", "operator", "source") == {
        "type": "date", "operator": "greaterThan", "source": None}
    assert rules[f"F2:F{last}"].type == "textLength"
    assert rules[f"I2:I{last}"].type == "custom" and rules[f"I2:I{last}"].operator is None
    # x14 rules: two bounds, and several ranges
    whole = rules[f"D2:D{last}"]
    assert (whole.origin, whole.operator, whole.formula1, whole.formula2) == \
        ("x14", "between", "1", "500")
    assert rules[f"K2:K{last}"].cells == (f"K2:K{last}", "M2")
    assert rules[f"K2:K{last}"].as_dict()["cells"] == f"K2:K{last} M2"
    # the rules no static reading resolves, by address
    assert [r.address for r in wb.validations() if r.resolved is False] == [
        f"Orders!A2:A{last}", f"Orders!G2:G{last}", f"Orders!J2:J{last}"]


def test_formulas_and_values_only_when_asked():
    wb = Workbook.open(Shop(orders=True).to_bytes())
    for rule in wb.validations():
        shown = rule.as_dict()
        assert "formula1" not in shown and "formula2" not in shown and "values" not in shown
    inline = next(r for r in wb.validations("Orders") if r.source == "inline")
    assert inline.as_dict(formulas=True, values=True)["formula1"] == '"S,M,L"'
    assert inline.as_dict(values=True)["values"] == ["S", "M", "L"]


def test_a_name_that_names_another_name_is_followed():
    parts = Shop(orders=True).parts()
    book = parts["xl/workbook.xml"].decode()
    book = book.replace('<definedName name="Markup">',
                        '<definedName name="Kinds">Categories</definedName>'
                        '<definedName name="Markup">')
    parts["xl/workbook.xml"] = book.encode()
    sheet = parts[ORDERS].decode().replace("<formula1>Bins</formula1>",
                                           "<formula1>Kinds</formula1>")
    parts[ORDERS] = sheet.encode()
    rule = next(r for r in Workbook.open(_zip(parts)).validations("Orders")
                if r.cells[0].startswith("J"))
    assert (rule.source, rule.resolved, rule.name, rule.range) == \
        ("name", True, "Kinds", "Lists!A2:A5")


def test_a_list_drawn_from_a_withheld_sheet_is_not_read():
    wb = Workbook.open(Shop(orders=True).to_bytes())
    wb.withhold_sheets(["Lists"])
    stock = wb.validations("Stock")[2]
    assert (stock.resolved, stock.reason, stock.values) == (False, "withheld", None)
    # the names on a withheld sheet are not the workbook's (Workbook.withhold_sheets)
    orders = wb.validations("Orders")[0]
    assert (orders.source, orders.resolved, orders.reason) == ("name", False, "unknown_name")


# -- kept through edits ------------------------------------------------------------------------

_STANDARD = re.compile(rb"<dataValidations\b.*?</dataValidations>", re.S)
_X14_RULE = re.compile(rb"<x14:dataValidation\b.*?</x14:dataValidation>", re.S)
_X14_BLOCK = re.compile(rb"<x14:dataValidations\b.*?</x14:dataValidations>", re.S)


@pytest.mark.parametrize("cells", [
    {"Stock!B4": "Ana", f"Stock!F{FIRST_INPUT}": 1.5, "Orders!F3": "a note",
     "Totals!B3": {"formula": "INDEX(TaxRates,3)"}},           # unrelated cells
    {f"Stock!C{FIRST_INPUT}": "Garden", f"Stock!D{FIRST_INPUT}": 7, "Orders!C3": "S",
     "Orders!D3": 12, "Orders!B3": "Kitchen"},                    # cells under the rules
], ids=["unrelated", "validated"])
def test_the_rules_survive_an_edit_byte_for_byte(cells):
    data = Shop(orders=True).to_bytes()
    wb = Workbook.open(data)
    before = _shown(wb.validations(), formulas=True)
    wb.set_cells(cells, replace_formulas=True)
    after = wb.to_bytes()
    for part in (STOCK, ORDERS):
        old, new = _part(data, part), _part(after, part)
        assert old != new                                   # the sheet was written
        assert _STANDARD.findall(old) == _STANDARD.findall(new)
        assert _X14_RULE.findall(old) == _X14_RULE.findall(new)
        # the x14 block itself: the same element, its namespace declaration written first
        # by the serialiser (the same information: compared canonically)
        for one, two in zip(_X14_BLOCK.findall(old), _X14_BLOCK.findall(new)):
            assert _canonical(one) == _canonical(two)
    assert _shown(Workbook.open(after).validations(), formulas=True) == before


def _canonical(fragment: bytes) -> bytes:
    wrapped = (b'<w xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
               + fragment + b"</w>")
    return etree.tostring(etree.fromstring(wrapped), method="c14n")


# -- a written value checked -------------------------------------------------------------------


def test_a_value_a_resolvable_rule_does_not_allow_is_named_never_refused():
    wb = Workbook.open(Shop(orders=True, table=True).to_bytes())
    written = wb.set_cells({
        "Orders!B3": "kitchen",            # in the named list, case aside
        "Orders!B4": "Toys",               # not in it
        "Orders!C3": "XL",                 # not in the inline list
        "Orders!D3": 12, "Orders!D4": 2.5, "Orders!D5": 900,   # whole between 1 and 500
        "Orders!D6": "12",                 # text under a whole-number rule
        "Orders!F3": "x" * 41, "Orders!F4": "short",           # length at most 40
        "Orders!H3": 0.21,                 # in the table column's values
        "Orders!K3": 0.5,                  # decimal between 0 and Lists!B5 (0.21)
        "Orders!A3": "Mug",                # INDIRECT: cannot be checked
        "Orders!G3": "Top",                # a name over OFFSET: cannot be checked
        "Orders!E3": {"date": "2026-12-01", "format": "yyyy-mm-dd"},   # after TODAY()
        "Orders!I3": 4,                    # custom
        "Orders!L3": "free",               # no rule
        f"Stock!D{FIRST_INPUT}": -1,       # whole >= 0
    })
    assert written.invalid == ("Orders!B4", "Orders!C3", "Orders!D4", "Orders!D5",
                               "Orders!D6", "Orders!F3", "Orders!K3", f"Stock!D{FIRST_INPUT}")
    assert written.unchecked == ("Orders!A3", "Orders!G3", "Orders!E3", "Orders!I3")
    assert written.changed == len(written.cells)            # every value written
    shown = written.as_dict()
    assert shown["invalid"][0] == "Orders!B4" and "Toys" not in str(shown)


def test_a_blank_and_a_formula_are_not_checked_as_values():
    wb = Workbook.open(Shop(orders=True).to_bytes())
    written = wb.set_cells({"Orders!C2": None, "Orders!C3": {"formula": '"X"'}})
    assert (written.invalid, written.unchecked) == ((), ("Orders!C3",))


def test_a_date_rule_with_bounds_compares_serials():
    parts = Shop(orders=True).parts()
    sheet = parts[ORDERS].decode().replace(
        '<dataValidation type="date" operator="greaterThan" sqref="E2:E50">'
        '<formula1>TODAY()</formula1>',
        '<dataValidation type="date" operator="greaterThan" sqref="E2:E50">'
        '<formula1>46023</formula1>')
    parts[ORDERS] = sheet.encode()
    wb = Workbook.open(_zip(parts))
    written = wb.set_cells({"Orders!E2": {"date": "2025-12-31"},
                            "Orders!E3": {"date": "2026-02-01", "format": "yyyy-mm-dd"},
                            "Orders!E4": 46000, "Orders!E5": 46100, "Orders!E6": "soon"})
    # 2025-12-31 is 46022, not after 46023; a number is compared as the serial it is
    assert written.invalid == ("Orders!E2", "Orders!E4", "Orders!E6")


# -- xlsx_read ---------------------------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, data: bytes):
    session = box.session()
    blob = session.add_blob(data, name="shop.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_read_lists_the_rules_over_the_target_on_its_first_page(box):
    session = _session(box, Shop(orders=True).to_bytes())
    # the input cells are empty: the rules are found over the range as asked
    result = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Stock!C10:D300",
                                                 "validations": True, "limit": 2})
    assert result.ok, result.error
    assert result.data["validations"] == [
        {"sheet": "Stock", "origin": "standard", "type": "whole",
         "operator": "greaterThanOrEqual", "allow_blank": True, "cells": "D10:D200"},
        {"sheet": "Stock", "origin": "x14", "type": "list", "allow_blank": True,
         "cells": "C10:C200", "source": "range", "resolved": True, "range": "Lists!A2:A5",
         "count": 4}]
    following = box.dispatch(session, "xlsx_read", {
        "doc": "d1", "target": "Stock!C10:D300", "validations": True, "limit": 2,
        "cursor": result.next_cursor})
    assert "validations" not in following.data
    plain = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Stock!C10:D300"})
    assert "validations" not in plain.data
    with_formulas = box.dispatch(session, "xlsx_read", {
        "doc": "d1", "target": "Orders!A2", "validations": True, "formulas": True})
    assert with_formulas.data["validations"] == [
        {"sheet": "Orders", "origin": "standard", "type": "list", "allow_blank": True,
         "cells": f"A2:A{ORDER_ROWS}", "source": "indirect", "resolved": False,
         "reason": "indirect", "uses": ["$B2"], "formula1": "INDIRECT($B2)"}]


def test_xlsx_set_cells_warns_of_a_value_its_rule_does_not_allow(box):
    session = _session(box, Shop(orders=True).to_bytes())
    result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Orders!C3", "text": "XL"}, {"address": "Orders!C4", "text": "S"}]})
    assert result.ok, result.error
    assert result.data["invalid"] == ["Orders!C3"]
    assert any("(invalid)" in w for w in result.warnings)
    assert all("XL" not in w for w in result.warnings)
