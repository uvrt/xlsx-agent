"""Typed cell writes (0.2, E3): ``wb.set_cells`` and ``xlsx_set_cells``.

On the corner shop's stock list (``tests/shop.py``, made in code): every kind of value
stored as its kind, text always literal, a blank never 0, nothing cut, dates in either date
system and only into a cell that shows one (or with a format), date strings in the locale's
order, no time zone dropped, a formula written over only when asked, a shared formula group
whose master is written handed on with every member's formula kept, array and merged cells
refused, every refusal before any change, and the edit within its footprint.
"""

from __future__ import annotations

import datetime as dt
import io
import json
import zipfile

import pytest
from lxml import etree

from ooxml_common.locales import Locale
from ooxml_edit.sheets import Formula, outside_footprint
from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.formula import ShiftError, is_external, shift, tokens
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from shop import FIRST_INPUT, Shop

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NL_BE = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")


def _open(**options) -> tuple[bytes, Workbook]:
    data = Shop(**options).to_bytes()
    return data, Workbook.open(data)


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _cell(data: bytes, ref: str, part: str = "xl/worksheets/sheet1.xml"):
    root = etree.fromstring(_parts(data)[part])
    for row_number, row in enumerate(root.iter(X + "row"), 1):
        for node in row.iter(X + "c"):
            if node.get("r") == ref:
                return node
    return None


def _refused(wb: Workbook, cells, **options) -> EditError:
    before = wb.to_bytes()
    version = wb.version
    with pytest.raises(EditError) as raised:
        wb.set_cells(cells, **options)
    assert wb.to_bytes() == before and wb.version == version
    return raised.value


# -- every kind of value ------------------------------------------------------------------------


def test_every_kind_of_value_is_stored_as_its_kind():
    data, wb = _open()
    written = wb.set_cells({
        "Stock!B15": "Spare fan",
        "Stock!C15": "=not a formula",           # text: literal
        "Stock!D15": 12,
        "Stock!E15": 2.5,
        "Stock!G15": True,
        "Stock!H10": {"date": "2026-10-20"},     # a cell that shows dates
        "Stock!B4": None,                        # a styled blank stays a styled blank
        "Stock!J15": {"formula": "=D15*E15"},
        "Stock!K15": {"value": "12"},            # text that looks like a number: text
    })
    assert written.cells[0] == "Stock!B15" and written.changed == 8
    assert written.old["Stock!H10"] == "number" and written.old["Stock!B15"] == "blank"
    assert written.calculation == "invalidated"
    assert [(wb.cell(a).v, wb.cell(a).t) for a in (
        "Stock!B15", "Stock!C15", "Stock!D15", "Stock!E15", "Stock!G15", "Stock!H10",
        "Stock!K15")] == [
        ("Spare fan", "text"), ("=not a formula", "text"), (12, "number"), (2.5, "number"),
        (True, "bool"), ("2026-10-20", "date"), ("12", "text")]
    assert wb.cell("Stock!H10").serial == 46315
    formula = wb.cell("Stock!J15")
    assert formula.formula and formula.formula_text == "D15*E15" and formula.cache == "absent"
    assert _cell(wb.to_bytes(), "C15").find(X + "f") is None
    assert outside_footprint(data, wb.to_bytes(), cells={"Stock": [
        "B15", "C15", "D15", "E15", "G15", "H10", "B4", "J15", "K15"]}) == []
    assert wb.save(io.BytesIO()).verified == 8         # B4 was blank already


def test_a_blank_is_never_a_zero_and_keeps_the_format():
    _, wb = _open()
    wb.set_cells({"Stock!D10": None, "Stock!E10": None})
    node = _cell(wb.to_bytes(), "D10")
    assert node.get("s") == "5" and node.find(X + "v") is None
    assert wb.cell("Stock!D10").t == "blank" and wb.cell("Stock!D10").v is None
    error = _refused(wb, {"Stock!D11": ""})
    assert error.code == "bad_value" and "write None for a blank" in str(error)


def test_text_is_never_cut_and_numbers_are_exact():
    _, wb = _open()
    assert _refused(wb, {"Stock!B12": "x" * 32_768}).code == "text_too_long"
    wb.set_cells({"Stock!B12": "x" * 32_767})
    assert wb.cell("Stock!B12").v == "x" * 32_767
    assert _refused(wb, {"Stock!D12": float("nan")}).code == "bad_number"
    assert _refused(wb, {"Stock!D12": float("inf")}).code == "bad_number"
    assert _refused(wb, {"Stock!D12": 2 ** 53 + 1}).code == "bad_number"
    assert _refused(wb, {"Stock!D12": b"bytes"}).code == "bad_value"
    assert _refused(wb, {"Stock!D12": {"value": 1, "date": "2026-01-01"}}).code == "bad_value"


def test_a_no_op_changes_nothing():
    data, wb = _open()
    written = wb.set_cells({"Stock!D10": 120, "Stock!B10": "Pencil", "Stock!G10": True})
    assert written.changed == 0 and wb.version == 0 and wb.to_bytes() == data


# -- dates --------------------------------------------------------------------------------------


@pytest.mark.parametrize("date1904, serial", [(False, 46315), (True, 44853)])
def test_a_date_is_the_serial_of_the_workbook_s_date_system(date1904, serial):
    _, wb = _open(date1904=date1904)
    wb.set_cells({"Stock!H10": dt.date(2026, 10, 20), "Stock!H11": dt.datetime(2026, 10, 21)})
    assert wb.cell("Stock!H10").serial == serial and wb.cell("Stock!H10").v == "2026-10-20"
    assert wb.cell("Stock!H11").v == "2026-10-21"     # midnight: the date
    assert _cell(wb.to_bytes(), "H10").findtext(X + "v") == str(serial)


def test_a_date_before_the_date_system_is_refused():
    _, wb = _open(date1904=True)
    assert _refused(wb, {"Stock!H10": dt.date(1903, 12, 31)}).code == "bad_date"


def test_a_date_into_a_cell_that_shows_no_date_is_refused_unless_a_format_is_given():
    data, wb = _open()
    error = _refused(wb, {"Stock!D10": {"date": "2026-10-20"}, "Stock!H16": dt.date(2026, 1, 1)})
    assert error.code == "date_format" and error.addresses == ("Stock!D10", "Stock!H16")
    assert '"format": "yyyy-mm-dd"' in str(error)
    wb.set_cells({"Stock!H16": {"date": "2026-01-02", "format": "yyyy\\-mm\\-dd"}})
    assert wb.cell("Stock!H16").v == "2026-01-02" and wb.cell("Stock!H16").t == "date"
    # the workbook's own format with that code (id 164) is reused: one new cell format
    # (General's with it), no new number format
    styles = _parts(wb.to_bytes())["xl/styles.xml"]
    assert styles.count(b"<numFmt ") == 2 and b'cellXfs count="9"' in styles
    wb.set_cells({"Stock!H17": {"date": "2026-01-03", "format": "d mmm yyyy"}})
    styles = _parts(wb.to_bytes())["xl/styles.xml"]
    assert styles.count(b"<numFmt ") == 3 and wb.cell("Stock!H17").v == "2026-01-03"
    assert _refused(wb, {"Stock!H17": {"date": "2026-01-02", "format": "0.00"}}).code \
        == "date_format"
    assert _refused(wb, {"Stock!H17": {"date": "2026-01-02", "format": "h:mm"}}).code \
        == "date_format"


def test_a_new_cell_takes_its_column_s_format_as_one_typed_there():
    original = Shop().to_bytes()
    parts = _parts(original)
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(
        b'<col min="2" max="3" width="16" customWidth="1"/>',
        b'<col min="2" max="3" width="16" customWidth="1"/>'
        b'<col min="8" max="8" width="12" style="2" customWidth="1"/>')
    data = _zip(parts)
    wb = Workbook.open(data)
    wb.set_cells({"Stock!H40": {"date": "2026-10-20"}})       # column H shows dates
    assert wb.cell("Stock!H40").t == "date" and _cell(wb.to_bytes(), "H40").get("s") == "2"


def _zip(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, part in parts.items():
            archive.writestr(name, part)
    return buffer.getvalue()


def test_date_strings_are_iso_or_in_the_locale_s_order():
    _, wb = _open()
    wb.locale = NL_BE
    wb.set_cells({"Stock!H10": {"date": "3/4/2026"}, "Stock!H11": {"date": "2026-04-05"}})
    assert wb.cell("Stock!H10").v == "2026-04-03"          # day first, as nl-BE writes it
    assert wb.cell("Stock!H11").v == "2026-04-05"
    _, us = _open()
    us.set_cells({"Stock!H10": {"date": "3/4/2026"}})
    assert us.cell("Stock!H10").v == "2026-03-04"          # en-US: month first
    for text in ("3/4/26", "13/13/2026", "2026-02-30", "20 October 2026",
                 "2026-10-20T10:00", "tomorrow"):
        assert _refused(wb, {"Stock!H12": {"date": text}}).code == "bad_date", text


def test_a_time_zone_or_a_time_of_day_is_never_dropped():
    _, wb = _open()
    aware = dt.datetime(2026, 10, 20, tzinfo=dt.timezone.utc)
    assert _refused(wb, {"Stock!H10": aware}).code == "date_time"
    assert _refused(wb, {"Stock!H10": dt.datetime(2026, 10, 20, 9, 30)}).code == "date_time"
    assert _refused(wb, {"Stock!H10": dt.time(9, 30)}).code == "date_time"
    assert _refused(wb, {"Stock!H10": dt.timedelta(hours=2)}).code == "date_time"
    assert _refused(wb, {"Stock!H10": {"date": aware}}).code == "date_time"


# -- formulas -----------------------------------------------------------------------------------


def test_a_formula_cell_is_written_over_only_when_asked():
    _, wb = _open()
    error = _refused(wb, {"Stock!A10": "P-999", "Stock!I11": 0.5})
    assert error.code == "formula_cell" and error.addresses == ("Stock!A10", "Stock!I11")
    assert "replace_formulas" in str(error)
    wb.set_cells({"Stock!A10": "P-999", "Stock!I11": {"formula": "=0.5"}},
                 replace_formulas=True)
    assert wb.cell("Stock!A10").v == "P-999" and not wb.cell("Stock!A10").formula
    assert wb.cell("Stock!I11").formula_text == "0.5"
    # the neighbours keep theirs
    assert wb.cell("Stock!A11").formula and wb.cell("Stock!I10").formula


def test_a_formula_given_is_checked():
    _, wb = _open()
    for bad in ("=", "", '="unclosed', "=SUM(A1", "=A1)", 12):
        assert _refused(wb, {"Stock!J10": {"formula": bad}}).code == "bad_formula", bad


@pytest.mark.parametrize("options", [{}, {"rless": True}], ids=["r", "rless"])
def test_the_master_of_a_shared_group_written_hands_the_group_on(options):
    data, wb = _open(**options)
    written = wb.set_cells({f"Stock!F{FIRST_INPUT}": 50}, replace_formulas=True)
    assert written.released == (f"Stock!F{FIRST_INPUT}",)
    column = [row[0] for row in wb.read(f"Stock!F{FIRST_INPUT}:F{FIRST_INPUT + 4}")]
    assert column[1].formula_text == f"D{FIRST_INPUT + 1}*E{FIRST_INPUT + 1}"
    assert [v.shared_from for v in column[2:]] == [f"Stock!F{FIRST_INPUT + 1}"] * 3
    assert column[0].v == 50 and not column[0].formula
    assert outside_footprint(data, wb.to_bytes(), cells={"Stock": [
        f"F{FIRST_INPUT}", f"F{FIRST_INPUT + 1}"]}) == []
    assert wb.validate() == [] and wb.save(io.BytesIO()).verified == 2
    assert f"F{FIRST_INPUT}:F{FIRST_INPUT + 1}" in json.dumps(wb.last_edit.as_dict())


def test_a_member_of_a_shared_group_written_leaves_the_others_theirs():
    _, wb = _open()
    wb.set_cells({f"Stock!F{FIRST_INPUT + 2}": 1}, replace_formulas=True)
    column = [row[0] for row in wb.read(f"Stock!F{FIRST_INPUT}:F{FIRST_INPUT + 4}")]
    assert column[0].formula_text == f"D{FIRST_INPUT}*E{FIRST_INPUT}"
    assert column[3].shared_from == f"Stock!F{FIRST_INPUT}" and not column[2].formula


def test_every_cell_of_a_shared_group_written_in_one_call():
    _, wb = _open()
    wb.set_cells({f"Stock!F{row}": row for row in range(FIRST_INPUT, FIRST_INPUT + 5)},
                 replace_formulas=True)
    assert [wb.cell(f"Stock!F{row}").formula for row in range(FIRST_INPUT,
                                                              FIRST_INPUT + 5)] == [False] * 5
    assert wb.validate() == [] and b'si="0"' not in _parts(wb.to_bytes())[
        "xl/worksheets/sheet1.xml"]
    _, wb = _open()                          # the master and one member: the rest kept
    written = wb.set_cells({f"Stock!F{FIRST_INPUT}": 1, f"Stock!F{FIRST_INPUT + 1}": 2},
                           replace_formulas=True)
    assert written.released == (f"Stock!F{FIRST_INPUT}",)
    column = [row[0] for row in wb.read(f"Stock!F{FIRST_INPUT}:F{FIRST_INPUT + 4}")]
    assert column[2].formula_text == f"D{FIRST_INPUT + 2}*E{FIRST_INPUT + 2}"
    assert [v.shared_from for v in column[3:]] == [f"Stock!F{FIRST_INPUT + 2}"] * 2


def test_a_shared_master_whose_formula_cannot_be_moved_is_refused():
    original = Shop().to_bytes()
    parts = _parts(original)
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(
        f'<f t="shared" ref="F{FIRST_INPUT}:F{FIRST_INPUT + 4}" si="0">D{FIRST_INPUT}*E'
        f'{FIRST_INPUT}</f>'.encode(),
        f'<f t="shared" ref="F{FIRST_INPUT}:F{FIRST_INPUT + 4}" si="0">D{FIRST_INPUT}*E'
        f'1048576</f>'.encode())
    wb = Workbook.open(_zip(parts))
    error = _refused(wb, {f"Stock!F{FIRST_INPUT}": 1}, replace_formulas=True)
    assert error.code == "shared_formula" and "off the sheet" in str(error)


def test_array_formula_and_merged_cells_are_refused():
    _, wb = _open(array=True)
    assert _refused(wb, {"Totals!C5": 1}, replace_formulas=True).code == "array_formula"
    assert _refused(wb, {"Totals!D5": 1}).code == "array_formula"       # in the block
    error = _refused(wb, {"Stock!B1": "x"})
    assert error.code == "merged_cell" and "write A1" in str(error)
    wb.set_cells({"Stock!A1": "Shop stock"})                          # the merge's first cell
    assert wb.cell("Stock!A1").v == "Shop stock"


def test_a_table_s_header_cell_is_refused():
    _, wb = _open(table=True)
    assert _refused(wb, {"Lists!A1": "Kind"}).code == "table_header"
    wb.set_cells({"Lists!B2": 0.07})
    assert wb.cell("Lists!B2").v == 0.07


def test_addresses_are_checked():
    _, wb = _open()
    for bad in ("D10", "Stock!D10:D12", "t:Rates", "Nowhere!A1", "Stock!"):
        assert _refused(wb, {bad: 1}).code == "bad_address", bad
    assert _refused(wb, [("Stock!D10", 1), ("#1!D10", 2)]).code == "duplicate_cell"
    with pytest.raises(EditError):
        wb.set_cells({})


# -- atomic -------------------------------------------------------------------------------------


def test_a_mixed_call_refuses_every_bad_cell_and_writes_none():
    original, wb = _open()
    error = _refused(wb, {"Stock!B10": "Spare fan", "Stock!D11": float("nan"),
                          "Stock!A12": "x", "Stock!D13": {"date": "2026-01-01"},
                          "Stock!B1": "merged"})
    assert error.code == "bad_number"
    assert error.problems == (("bad_number", "Stock!D11"), ("formula_cell", "Stock!A12"),
                              ("date_format", "Stock!D13"), ("merged_cell", "Stock!B1"))
    assert "3 other refusal(s)" in str(error)
    assert wb.to_bytes() == original and not wb.history.can_undo()
    assert wb.calculation_status == "cached_unverified"


def test_a_refusal_inside_a_batch_rolls_back_the_batch():
    original, wb = _open()
    with pytest.raises(EditError) as raised:
        with wb.batch():
            wb.set_cells({"Stock!B10": "Spare fan"})
            try:
                wb.set_cells({"Stock!D11": float("nan")})
            except EditError:
                pass
    assert raised.value.code == "rolled_back" and wb.to_bytes() == original


def test_undo_and_redo_restore_the_cache_status():
    original, wb = _open()
    assert wb.cell("Totals!B1").cache == "cached_unverified"
    wb.set_cells({"Stock!D10": 7})
    edited = wb.to_bytes()
    assert wb.cell("Totals!B1").cache == "invalidated"
    wb.undo()
    assert wb.to_bytes() == original and wb.cell("Totals!B1").cache == "cached_unverified"
    wb.redo()
    assert wb.to_bytes() == edited and wb.cell("Totals!B1").cache == "invalidated"


# -- freshness ----------------------------------------------------------------------------------


def test_every_formula_value_carries_its_cache_status():
    _, wb = _open()
    values = [wb.cell(a) for a in ("Stock!A10", "Stock!F11", "Totals!B1", "Stock!D10")]
    assert [v.cache for v in values] == ["cached_unverified"] * 3 + [None]
    assert values[1].as_dict()["cache"] == "cached_unverified"
    wb.set_cells({"Lists!B2": 0.07})          # one input anywhere: every cache invalidated
    assert {wb.cell(a).cache for a in ("Stock!A10", "Stock!F11", "Totals!B1")} == {
        "invalidated"}
    assert wb.catalogue()["cache"] == "invalidated"


@pytest.mark.parametrize("calc", ["manual", "full"])
def test_a_workbook_arriving_with_manual_or_full_calculation_reads_invalidated(calc):
    _, wb = _open(calc=calc)
    assert wb.calculation_status == "invalidated"
    assert wb.cell("Totals!B1").cache == "invalidated"
    assert wb.catalogue()["cache"] == "invalidated"


def test_a_formula_on_another_workbook_is_flagged_external():
    _, wb = _open(external=True)
    value = wb.cell("Totals!B4")
    assert value.external and value.as_dict()["external"] is True
    assert not wb.cell("Totals!B1").external and wb.external_links == 1
    catalogue = wb.catalogue()
    assert catalogue["flags"]["external_links"] == 1
    assert {"code": "external_formula", "address": "Totals!B4"} in catalogue["problems"]


def test_queries_derives_compares_validations_and_profiles_say_the_cache_status():
    _, wb = _open(table=True)
    wb.declare_region("Items", f"Stock!A9:I{FIRST_INPUT + 4}")
    before = Workbook.open(Shop(table=True).to_bytes())
    before.declare_region("Items", f"Stock!A9:I{FIRST_INPUT + 4}")
    wb.set_cells({"Stock!D11": 100, "Stock!F10": {"formula": "=D10*E10*2"}},
                 replace_formulas=True)
    rows = wb.table("t:Items").query(["Id", "Value"]).rows
    assert [v.as_dict()["cache"] for v in rows[0]] == ["invalidated", "absent"]
    derived = wb.derive("t:Items", measures=[{"fn": "sum", "column": "Value", "as": "v"}])
    # the Value cells and the Id cells (formulas too), whose keys name the sources
    assert derived.cache == {"invalidated": 9, "absent": 1}
    assert derived.dataset().cache == derived.cache and derived.dataset().invalidated
    # Conservative: the Id keys are formulas too, so a sum of plain inputs is from
    # invalidated caches as well (its sources are named by those keys).
    keyed = wb.derive("t:Items", measures=[{"fn": "sum", "column": "Quantity", "as": "q"}])
    assert keyed.cache == {"invalidated": 5}
    plain = wb.derive("t:Rates", measures=[{"fn": "max", "column": "Tax", "as": "top"}])
    assert plain.cache == {} and plain.dataset().cache is None
    report = wb.validate_data({"tables": [{"table": "t:Items", "columns": [
        {"name": "Value", "type": "number"}]}]})
    assert report.cache == {"invalidated": 4, "absent": 1}
    value = next(c for c in wb.profile("t:Items") if c.column == "Value")
    assert value.as_dict()["cache"] == {"invalidated": 4, "absent": 1}
    assert "cache" not in next(c for c in wb.profile("t:Items")
                               if c.column == "Quantity").as_dict()
    diff = Workbook.compare(before, wb, tables=["t:Items"], key="Id")
    changes = {change.column: change for change in diff.tables[0].changed}
    assert changes["Value"].a.cache == "cached_unverified"
    assert changes["Value"].b.cache == "absent" and changes["Value"].b.as_dict()["cache"] \
        == "absent"
    assert changes["Quantity"].b.cache is None


# -- the lexer ------------------------------------------------------------------------------------


@pytest.mark.parametrize("text, rows, columns, moved", [
    ("D10*E10", 1, 0, "D11*E11"),
    ("$D$10+D10", 2, 1, "$D$10+E12"),
    ("D$10+$D10", 1, 1, "E$10+$D11"),
    ("SUM(Stock!D10:D14)", 1, 0, "SUM(Stock!D11:D15)"),
    ("'Q3 sales'!B2*2", 0, 2, "'Q3 sales'!D2*2"),
    ("SUM(A:A)+SUM(3:5)", 1, 1, "SUM(B:B)+SUM(4:6)"),
    ("Jan:Mar!A1", 1, 0, "Jan:Mar!A2"),
    ('IF(A1="B2",B2,"C3")', 1, 0, 'IF(A2="B2",B3,"C3")'),
    ("Rates[[#This Row],[Tax]]*A1", 1, 0, "Rates[[#This Row],[Tax]]*A2"),
    ("VLOOKUP(C10,Lists!$A$2:$B$5,2,FALSE)", 1, 0, "VLOOKUP(C11,Lists!$A$2:$B$5,2,FALSE)"),
    ("LOG10(A1)+TaxRates", 1, 0, "LOG10(A2)+TaxRates"),
    ("[1]Prices!$B$2*B3", 1, 0, "[1]Prices!$B$2*B4"),
    ("A1#+@B2", 1, 0, "A2#+@B3"),
    ("Stock!#REF!+1E3", 1, 0, "Stock!#REF!+1E3"),
])
def test_references_move_by_the_offset_and_nothing_else_does(text, rows, columns, moved):
    assert shift(text, rows, columns) == moved
    assert "".join(t.text for t in tokens(text)) == text


@pytest.mark.parametrize("text", ['"open', "'Stock!A1", "Rates[Tax", "A1048576", "XFD1"])
def test_what_cannot_be_moved_safely_is_refused(text):
    with pytest.raises(ShiftError):
        shift(text, 1, 1)


def test_another_workbook_s_references_are_external():
    assert is_external("[1]Prices!$B$2*2") and is_external("'[2]Q3 sales'!B2")
    assert is_external("[1]!Rate*2") and is_external("Rate*2", {"rate"})
    assert not is_external("Rates[Tax]*2") and not is_external('"[1]Prices!A1"')
    assert not is_external("SUM(Stock!A1:A3)")


# -- the tool -----------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, **options):
    session = box.session(**options)
    blob = session.add_blob(Shop().to_bytes(), name="stock.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_set_cells_writes_typed_cells_in_one_step(box):
    session = _session(box)
    result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B15", "text": "=literal"},
        {"address": "Stock!D15", "number": 3},
        {"address": "Stock!G15", "bool": False},
        {"address": "Stock!H10", "date": "2026-10-20"},
        {"address": "Stock!H15", "date": "2026-10-21", "format": "yyyy-mm-dd"},
        {"address": "Stock!B4", "blank": True},
        {"address": "Stock!J15", "formula": "=D15*2"}]})
    assert result.ok, result.to_json()
    data = result.data
    assert data["written"] == 7 and data["changed"] == 6 and data["calculation"] == "invalidated"
    assert data["cells"][3] == {"a": "Stock!H10", "old": "number"}
    assert "literal" not in result.summary and session.entry("d1").version == 1
    read = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Stock!B15:J15"})
    first = read.data["rows"][0]
    assert first[0] == {"a": "Stock!B15", "v": "=literal", "t": "text"}
    assert first[-1]["cache"] == "absent"
    total = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Totals!B1"})
    assert total.data["rows"][0][0]["cache"] == "invalidated"
    assert box.dispatch(session, "undo", {"doc": "d1"}).ok
    total = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Totals!B1"})
    assert total.data["rows"][0][0]["cache"] == "cached_unverified"


def test_xlsx_set_cells_refuses_with_every_problem_and_changes_nothing(box):
    session = _session(box)
    result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B10", "text": "ok"},
        {"address": "Stock!A10", "text": "over a formula"},
        {"address": "Stock!D12", "date": "2026-10-20"}]})
    assert result.error.code == "refused"
    assert result.error.details["code"] == "formula_cell"
    assert result.error.details["problems"] == [{"code": "formula_cell", "a": "Stock!A10"},
                                                {"code": "date_format", "a": "Stock!D12"}]
    assert "over a formula" not in json.dumps(result.to_json())
    assert session.entry("d1").version == 0
    for cell, field in [({"address": "Stock!B10"}, "cells[0]"),
                        ({"address": "Stock!B10", "text": "a", "number": 1}, "cells[0]"),
                        ({"address": "Stock!B10", "number": 1, "format": "yyyy"},
                         "cells[0].format")]:
        bad = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [cell]})
        assert bad.error.code == "invalid_arguments" and bad.error.field == field


def test_xlsx_set_cells_is_deferred_and_the_edit_role_s(box):
    core = {d["name"] for d in box.definitions("anthropic", groups="core")}
    assert "xlsx_set_cells" not in core
    group = {d["name"] for d in box.definitions("anthropic", groups="xlsx_edit")}
    assert "xlsx_set_cells" in group
    tool = next(t for t in TOOLS if t.name == "xlsx_set_cells")
    assert tool.roles == {"edit"} and tool.group == "xlsx_edit"
    review = {d["name"] for d in box.definitions("anthropic", defer=False, roles=["review"])}
    assert "xlsx_set_cells" not in review
    session = _session(box, roles=["review"])
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!B10", "text": "x"}]})
    assert refused.error.code == "refused" and refused.error.details["role"] == "edit"


def test_xlsx_set_cells_takes_at_most_5000_cells(box):
    session = _session(box)
    cells = [{"address": f"Stock!K{n}", "number": n} for n in range(1, 5002)]
    result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": cells})
    assert result.error.code == "limit"
    ok = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": cells[:5000]})
    assert ok.ok and ok.data["written"] == 5000 and ok.data["cells_cut"] == 4950


def test_xlsx_set_cells_refuses_a_column_withheld_from_the_model(box):
    from xlsx_agent.tools import declare

    session = box.session()
    declare(session, [{"name": "items", "sheets": "Stock", "header_row": 9, "key": "Id",
                       "fields": [{"header": "Unit price", "hidden": True}]}])
    blob = session.add_blob(Shop().to_bytes(), name="stock.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    refused = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!E10", "number": 1}]})
    assert refused.error.code == "refused" and refused.error.details["code"] == "withheld"
    assert box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!D10", "number": 1}]}).ok


def test_a_workbook_without_cell_formats_takes_no_date():
    import re

    parts = _parts(Shop().to_bytes())
    del parts["xl/styles.xml"]
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship Id="rId4"[^>]*/>', "",
                                                 rels).encode()
    for part in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet3.xml"):
        parts[part] = re.sub(rb' s="[0-9]+"', b"", parts[part])
    wb = Workbook.open(_zip(parts))
    assert wb.editor.cells.styles is None
    assert _refused(wb, {"Stock!H10": dt.date(2026, 1, 1)}).code == "date_format"
    assert _refused(wb, {"Stock!H10": {"date": "2026-01-01", "format": "yyyy-mm-dd"}}).code \
        == "date_format"
    wb.set_cells({"Stock!D10": 3})
    assert wb.cell("Stock!D10").v == 3


def test_a_change_says_which_sheets_it_changed_for_the_application(box):
    # Round 13, R2: ooxml-edit's Result.touched, by sheet; never in what the model reads.
    session = _session(box)
    result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Stock!D15", "number": 3}, {"address": "Stock!B15", "text": "new"}]})
    assert result.ok, result.to_json()
    sheets = [entry for entry in result.touched if "sheet" in entry]
    assert sheets == [{"doc": "d1", "sheet": "Stock"}]
    assert all(set(entry) == {"doc", "part"} for entry in result.touched if entry not in sheets)
    assert "touched" not in result.to_json()
    undone = box.dispatch(session, "undo", {"doc": "d1"})
    assert undone.touched == result.touched
    read = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Stock!B15"})
    assert read.touched == []
