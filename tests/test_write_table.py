"""Round 44, lane T: a new workbook (R51), a block or table written (R54, E4), rows appended.

Made-up workbooks only: a stock list, orders and a log.  What Excel does is measured in the
modules' docstrings (``xlsx_agent.edit.blank``, ``xlsx_agent.edit.tables``); the Excel
oracle's runs are ``tests/test_oracle_tables.py``.

Tool-design rules (agent-work/tool-design-rules.md) and their tests here:
1 (accept intent): ``name`` alone asks for a table, ``""`` is a blank, a row by name,
``Stock!B2`` as the origin, ``{}`` a blank cell in the tool -- ``test_intent_*``;
2 (a refusal carries the way out): a refused table name offers one that would do, a block
over values lists every cell and says ``overwrite``, an ambiguous target names both forms --
``test_refusal_*``; 3 (a no-op is not an error): writing what a block already holds --
``test_rewriting_the_same_block_changes_nothing``; 4 (size budget): refusals list at most 50
problems, the tool's cells cap -- ``test_refusals_are_capped``, ``test_the_tool_caps_cells``;
5 (repeated work, one call): many rows in one step, one undo -- ``test_one_step_one_undo``;
6 (the template owns the design): the workbook's default table style, Excel's short date for
a typed date, the format above a new table row -- ``test_template_*``.
"""

from __future__ import annotations

import datetime as dt
import io
import zipfile

import pytest
from lxml import etree

from build import Formula as BuiltFormula
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.edit.blank import build
from xlsx_agent.edit.tables import suggest_name, table_name_problem
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.shared import NEW_DOCUMENT

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _xml(data: bytes, part: str):
    return etree.fromstring(_parts(data)[part])


def _orders(*, totals: bool = True, below: dict | None = None) -> bytes:
    """An ``Orders`` table, A1:D5 with a totals row (A1:D4 without), its ``Total`` column
    calculated; ``F1`` sums the table by its name."""
    book = WorkbookBuilder()
    sheet = book.sheet("Orders")
    sheet.rows("A1", [["Item", "Qty", "Price", "Total"]])
    lines = [("Lamp", 3, 2.5), ("Fan", 4, 10.0), ("Desk", 1, 99.0)]
    for row, (item, qty, price) in enumerate(lines, 2):
        sheet.rows(f"A{row}", [[item, qty, price]])
        sheet.cell(f"D{row}", BuiltFormula("Orders[[#This Row],[Qty]]*Orders[[#This Row],"
                                           "[Price]]", cached=qty * price))
    date_style = book.style(num_fmt="yyyy-mm-dd")
    sheet.cell("C4", 99.0, style=date_style)     # a format above the new rows
    ref = "A1:D4"
    kwargs: dict = {}
    if totals:
        sheet.cell("A5", "Total")
        sheet.cell("D5", BuiltFormula("SUBTOTAL(109,Orders[Total])", cached=146.5))
        ref = "A1:D5"
        kwargs = {"totals_labels": {"Item": "Total"}, "totals_functions": {"Total": "sum"}}
    sheet.cell("F1", BuiltFormula("SUM(Orders[Total])", cached=146.5))
    for address, value in (below or {}).items():
        sheet.cell(address, value)
    sheet.table("Orders", ref, ["Item", "Qty", "Price", "Total"], write_header=False,
                calculated={"Total": "Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]"},
                style_name="TableStyleMedium2", **kwargs)
    return book.to_bytes()


# -- R51: a new workbook --------------------------------------------------------------------------


def test_a_new_workbook_has_the_parts_excel_writes():
    data = build(created=dt.datetime(2026, 11, 8, 12, 0, tzinfo=dt.timezone.utc))
    parts = _parts(data)
    assert sorted(parts) == sorted([
        "[Content_Types].xml", "_rels/.rels", "docProps/app.xml", "docProps/core.xml",
        "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/theme/theme1.xml",
        "xl/workbook.xml", "xl/worksheets/sheet1.xml"])
    styles = etree.fromstring(parts["xl/styles.xml"])
    font = styles.find(f"{X}fonts/{X}font")
    assert font.find(f"{X}sz").get("val") == "12" and font.find(f"{X}name").get("val") == \
        "Aptos Narrow"
    assert styles.find(f"{X}tableStyles").get("defaultTableStyle") == "TableStyleMedium2"
    theme = parts["xl/theme/theme1.xml"].decode()
    assert '<a:minorFont><a:latin typeface="Aptos Narrow"' in theme
    assert '<a:font script="Thai" typeface="Tahoma"/>' in theme
    sheet = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
    assert sheet.find(f"{X}sheetFormatPr").get("defaultRowHeight") == "16"
    assert sheet.find(f"{X}sheetViews/{X}sheetView").get("tabSelected") == "1"
    assert b"absPath" not in parts["xl/workbook.xml"]          # never the machine's
    assert build(created=dt.datetime(2026, 11, 8, 12, 0, tzinfo=dt.timezone.utc)) == data


def test_workbook_new_opens_unchanged_with_its_sheets():
    wb = Workbook.new(["Stock", "Notes"], author="Stock clerk", title="Stock list")
    assert [s.name for s in wb.sheets] == ["Stock", "Notes"]
    assert wb.version == 0 and wb.validate() == [] and wb.calculation_status != "invalidated"
    core = _parts(wb.to_bytes())["docProps/core.xml"].decode()
    assert "<dc:creator>Stock clerk</dc:creator>" in core and "Stock list" in core
    assert Workbook.new().sheets[0].name == "Sheet1"
    app = _parts(wb.to_bytes())["docProps/app.xml"].decode()
    assert "<vt:lpstr>Stock</vt:lpstr><vt:lpstr>Notes</vt:lpstr>" in app


@pytest.mark.parametrize("names", [["A", "a"], [""], ["x" * 32], ["a/b"], ["'x"]])
def test_a_new_workbook_refuses_names_excel_refuses(names):
    with pytest.raises(EditError) as refused:
        Workbook.new(names)
    assert refused.value.code == "bad_value"


def test_from_tables_writes_typed_rows_on_new_sheets():
    wb = Workbook.from_tables({
        "Stock": {"columns": ["Item", "Qty", "Added"], "table": "Stock",
                  "rows": [["Lamp", 3, dt.date(2026, 1, 5)], ["=not a formula", True, None]]},
        "Notes": {"columns": ["Note"], "rows": [[{"markdown": "**first**"}]]}})
    assert [s.name for s in wb.sheets] == ["Stock", "Notes"]
    assert [t.name for t in wb.tables] == ["Stock"]
    rows = [[v.v for v in row] for row in wb.read("Stock!A2:C3")]
    assert rows == [["Lamp", 3, "2026-01-05"], ["=not a formula", True, None]]
    assert wb.version == 1 and wb.validate() == []
    reopened = Workbook.open(wb.to_bytes())
    assert [t.name for t in reopened.tables] == ["Stock"]


def test_from_tables_refuses_a_bad_entry_before_anything():
    with pytest.raises(EditError):
        Workbook.from_tables({"Stock": {"columns": ["A"], "rows": [], "colour": "red"}})
    with pytest.raises(EditError):
        Workbook.from_tables({})


# -- write_table ----------------------------------------------------------------------------------


def test_write_table_writes_typed_values_and_returns_the_range():
    wb = Workbook.new(["Stock"])
    done = wb.write_table("Stock", "B2", ["Item", "Qty", "Ok", "Added", "Value"],
                          [["=x", 3, True, dt.date(2026, 1, 5), {"formula": "=C3*2"}]])
    assert (done.range, done.header, done.table, done.rows) == ("Stock!B2:F3", "Stock!B2:F2",
                                                                 None, 1)
    row = next(iter(wb.read("Stock!B3:F3")))
    assert [(v.t, v.v) for v in row[:4]] == [("text", "=x"), ("number", 3), ("bool", True),
                                             ("date", "2026-01-05")]
    assert row[4].formula and row[4].cache == "absent"
    assert wb.tables == []


def test_as_table_makes_the_table_excel_makes():
    wb = Workbook.new(["Stock"])
    done = wb.write_table("Stock", "B2", ["Item", "Qty"], [["Lamp", 3], ["Fan", 4]],
                          as_table=True)
    assert done.table == "Table1"
    data = wb.to_bytes()
    table = _xml(data, "xl/tables/table1.xml")
    assert dict(table.attrib) == {"id": "1", "name": "Table1", "displayName": "Table1",
                                  "ref": "B2:C4", "totalsRowShown": "0"}
    assert table.find(f"{X}autoFilter").get("ref") == "B2:C4"
    assert [c.get("name") for c in table.iter(f"{X}tableColumn")] == ["Item", "Qty"]
    info = table.find(f"{X}tableStyleInfo")
    assert dict(info.attrib) == {"name": "TableStyleMedium2", "showFirstColumn": "0",
                                 "showLastColumn": "0", "showRowStripes": "1",
                                 "showColumnStripes": "0"}
    sheet = _xml(data, "xl/worksheets/sheet1.xml")
    children = [etree.QName(c).localname for c in sheet]
    assert children[-1] == "tableParts" and children.index("pageMargins") < len(children) - 1
    assert wb.validate() == []
    second = wb.write_table("Stock", "E2", ["Name"], [], as_table=True)
    assert (second.table, second.range) == ("Table2", "Stock!E2:E3")   # one empty data row
    assert _xml(wb.to_bytes(), "xl/tables/table2.xml").get("id") == "2"


@pytest.mark.parametrize("name,refused", [
    # measured: Excel for Mac 16.106, 2026-11-08 (xlsx_agent.edit.tables)
    ("A1", True), ("a1", True), ("XFD1", True), ("R1C1", True), ("R", True), ("c", True),
    ("R1", True), ("C1", True), ("R5C", True), ("1abc", True), ("a b", True), ("a-b", True),
    (".a", True), ("$A$1", True), ("A1:B2", True), ("a!b", True), ("N" + "a" * 255, True),
    ("RC", False), ("Rx", False), ("_abc", False), ("\\abc", False), ("a.b", False),
    ("a?b", False), ("äbc", False), ("XFE1", False), ("XFD1048577", False), ("ABCD1", False),
    ("TRUE", False), ("N" + "a" * 254, False), ("日本", False), ("a€b", False),
])
def test_table_names_as_excel_takes_them(name, refused):
    assert (table_name_problem(name) is not None) is refused


def test_table_names_are_unique_with_defined_names_case_aside():
    book = WorkbookBuilder()
    book.sheet("Stock").rows("A1", [["Item"], ["Lamp"]])
    book.define_name("Spare", "Stock!$A$2")
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(EditError) as refused:
        wb.write_table("Stock", "C1", ["Item"], [["Fan"]], as_table="spare")
    assert refused.value.code == "bad_table_name"
    assert "'spare1'" in str(refused.value) or "would do" in str(refused.value)


def test_refusal_a_table_name_offers_one_that_would_do():
    for name in ("S1", "a b", "1st"):
        suggestion = suggest_name(name, set())
        assert table_name_problem(suggestion) is None, (name, suggestion)
    wb = Workbook.new()
    with pytest.raises(EditError) as refused:
        wb.write_table("Sheet1", "A1", ["Item"], [["Lamp"]], as_table="S1")
    assert "'S1_' would do" in str(refused.value)


def test_refusal_a_block_over_values_lists_every_cell_and_the_way_out():
    wb = Workbook.new(["Stock"])
    wb.write_table("Stock", "A1", ["Item", "Qty"], [["Lamp", 3]])
    before = wb.to_bytes()
    with pytest.raises(EditError) as refused:
        wb.write_table("Stock", "A2", ["Item", "Qty"], [["Fan", 4]])
    assert refused.value.code == "not_empty"
    assert refused.value.addresses == ("Stock!A2", "Stock!B2")
    assert "overwrite=True" in str(refused.value)
    assert wb.to_bytes() == before
    wb.write_table("Stock", "A2", ["Item", "Qty"], [["Fan", 4]], overwrite=True)
    assert [v.v for v in next(iter(wb.read("Stock!A3:B3")))] == ["Fan", 4]


def test_merges_and_tables_are_never_written_over():
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["Item"], ["Lamp"]])
    sheet.table("Lines", "A1:A2", ["Item"], write_header=False)
    sheet.merge("D1:E1")
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(EditError) as refused:
        wb.write_table("Stock", "A2", ["Qty", "Price", "Left", "Right"], [[1, 2, 3, 4]],
                       overwrite=True)
    assert refused.value.code == "table_overlap"
    with pytest.raises(EditError) as refused:
        wb.write_table("Stock", "C1", ["Qty", "Price", "Left"], [[1, 2, 3]], overwrite=True)
    assert refused.value.code == "merged_cell" and refused.value.addresses == ("Stock!D1:E1",)


def test_headers_of_a_table_are_unique_text_on_one_line():
    wb = Workbook.new()
    for headers in (["Item", "item"], ["Item", ""], ["Item", "Two\nlines"], ["Item", 3]):
        with pytest.raises(EditError) as refused:
            wb.write_table("Sheet1", "A1", headers, [["a", "b"]], as_table=True)
        assert refused.value.code == "bad_value"
    with pytest.raises(EditError):
        wb.write_table("Sheet1", "A1", None, [["a"]], as_table=True)


def test_a_row_longer_than_the_columns_is_refused_a_shorter_one_ends_blank():
    wb = Workbook.new()
    with pytest.raises(EditError) as refused:
        wb.write_table("Sheet1", "A1", ["Item"], [["a", "b"]])
    assert refused.value.addresses == ("rows[0]",)
    wb.write_table("Sheet1", "A1", ["Item", "Qty"], [["a"]])
    assert [v.t for v in next(iter(wb.read("Sheet1!A2:B2")))] == ["text", "blank"]


def test_intent_a_name_alone_asks_for_a_table_and_the_origin_may_carry_the_sheet():
    wb = Workbook.new(["Stock"])
    done = wb.write_table(None, "Stock!C3", ["Item"], [["Lamp"]], name="Lines")
    assert (done.table, done.range) == ("Lines", "Stock!C3:C4")
    with pytest.raises(EditError):
        wb.write_table("Sheet9", "Stock!A1", ["Item"], [["x"]])


def test_intent_empty_text_is_a_blank_and_a_row_may_be_by_name():
    wb = Workbook.new()
    wb.write_table("Sheet1", "A1", ["Item", "Qty"], [["", 3], {"qty": 4, "Item": "Fan"}])
    rows = [[v.v for v in row] for row in wb.read("Sheet1!A2:B3")]
    assert rows == [[None, 3], ["Fan", 4]]
    with pytest.raises(EditError) as refused:
        wb.write_table("Sheet1", "D1", ["Item"], [{"colour": "red"}])
    assert "the columns are" in str(refused.value)


def test_template_the_workbook_s_default_table_style():
    book = WorkbookBuilder()
    book.sheet("Stock").cell("A1", "x")
    wb = Workbook.open(book.to_bytes())
    styles = wb.package.tree(wb._related("styles"))
    node = etree.SubElement(styles, f"{X}tableStyles")
    node.set("count", "0")
    node.set("defaultTableStyle", "TableStyleLight9")
    wb.package.mark_dirty(wb._related("styles"))
    wb.write_table("Stock", "C1", ["Item"], [["Lamp"]], as_table=True)
    table = _xml(wb.to_bytes(), "xl/tables/table1.xml")
    assert table.find(f"{X}tableStyleInfo").get("name") == "TableStyleLight9"
    with pytest.raises(EditError) as refused:
        wb.write_table("Stock", "E1", ["Item"], [["Lamp"]], as_table=True, style="Fancy")
    assert refused.value.code == "bad_table_style"


def test_template_a_date_gets_excel_s_short_date_unless_the_cell_shows_one():
    book = WorkbookBuilder()
    shown = book.style(num_fmt="yyyy-mm-dd")
    book.sheet("Log").cell("B2", None, style=shown)
    wb = Workbook.open(book.to_bytes())
    wb.write_table("Log", "A1", ["When", "Kept"],
                   [[dt.date(2026, 1, 5), dt.date(2026, 1, 6)],
                    [{"date": "2026-01-07", "format": "d mmm yyyy"}, None]])
    formats = {v.a: v.format for row in wb.read("Log!A2:B3", formats=True) for v in row}
    assert formats["Log!B2"] == "yyyy-mm-dd"            # the cell's own date format kept
    assert formats["Log!A3"] == "d mmm yyyy"
    xf = wb.styles  # the short date: built-in 14
    cell = _xml(wb.to_bytes(), "xl/worksheets/sheet1.xml").find(f".//{X}c[@r='A2']")
    styles = _xml(wb.to_bytes(), "xl/styles.xml")
    assert styles.find(f"{X}cellXfs")[int(cell.get("s"))].get("numFmtId") == "14"
    assert xf is not None


def test_a_workbook_without_shared_strings_gets_a_table_of_them():
    wb = Workbook.new()
    assert wb._related("sharedStrings") is None
    wb.write_table("Sheet1", "A1", ["Item"], [["Lamp"]])
    parts = _parts(wb.to_bytes())
    assert b'<c r="A2" t="s">' in parts["xl/worksheets/sheet1.xml"]
    assert b'count="2" uniqueCount="2"' in parts["xl/sharedStrings.xml"]
    wb.undo()
    assert wb._related("sharedStrings") is None


def test_one_step_one_undo():
    original = Workbook.new().to_bytes()
    wb = Workbook.open(original)
    wb.write_table("Sheet1", "A1", ["N"], [[n] for n in range(500)], as_table=True)
    assert wb.version == 1 and len(wb.tables) == 1
    wb.undo()
    assert wb.to_bytes() == original and wb.tables == []
    wb.redo()
    assert len(wb.tables) == 1


def test_rewriting_the_same_block_changes_nothing():
    wb = Workbook.new()
    wb.write_table("Sheet1", "A1", ["Item"], [["Lamp"]])
    again = wb.write_table("Sheet1", "A1", ["Item"], [["Lamp"]], overwrite=True)
    assert again.written.changed == 0 and wb.version == 1


def test_refusals_are_capped():
    wb = Workbook.new()
    wb.write_table("Sheet1", "A1", None, [[n] * 10 for n in range(20)])
    with pytest.raises(EditError) as refused:
        wb.write_table("Sheet1", "A1", None, [[n] * 10 for n in range(20)])
    assert len(refused.value.problems) == 50 and len(refused.value.addresses) == 20


def test_guards_a_protected_sheet_and_a_policy_refuse():
    book = WorkbookBuilder()
    book.sheet("Stock").cell("A1", "x")
    locked = book.sheet("Locked")
    locked.protected = True
    data = book.to_bytes()
    wb = Workbook.open(data)
    with pytest.raises(EditError) as refused:
        wb.write_table("Locked", "A1", ["Item"], [["Lamp"]])
    assert refused.value.code == "protected"
    wb = Workbook.open(data, policy={"cells": ["Stock!A1:A3"]})
    with pytest.raises(EditError) as refused:
        wb.write_table("Stock", "C1", ["Item"], [["Lamp"]])
    assert refused.value.code == "not_writable"
    assert wb.to_bytes() == data


# -- append_rows ----------------------------------------------------------------------------------


def test_append_to_a_table_with_a_totals_row():
    wb = Workbook.open(_orders())
    done = wb.append_rows("t:Orders", [["Chair", 2, 20], {"item": "Shelf", "Price": 35.5}])
    assert (done.range, done.table_range, done.totals, done.filled) == (
        "Orders!A5:D6", "Orders!A1:D7", "Orders!A7:D7", ("Total",))
    data = wb.to_bytes()
    table = _xml(data, "xl/tables/table1.xml")
    assert table.get("ref") == "A1:D7" and table.find(f"{X}autoFilter").get("ref") == "A1:D6"
    sheet = _xml(data, "xl/worksheets/sheet1.xml")
    cell = {c.get("r"): c for c in sheet.iter(f"{X}c")}
    assert cell["D7"].findtext(f"{X}f") == "SUBTOTAL(109,Orders[Total])"
    assert cell["A7"] is not None and "A5" in cell
    assert cell["D5"].findtext(f"{X}f") == ("Orders[[#This Row],[Qty]]*"
                                            "Orders[[#This Row],[Price]]")
    assert cell["C5"].get("s") == cell["C4"].get("s")       # the format above, measured
    assert [t.area.ref for t in wb.tables] == ["A1:D7"]
    assert wb.validate() == []


def test_append_to_a_table_without_totals_and_undo():
    original = _orders(totals=False)
    wb = Workbook.open(original)
    done = wb.append_rows("t:Orders", [["Chair", 2, 20]])
    assert (done.range, done.totals) == ("Orders!A5:D5", None)
    table = _xml(wb.to_bytes(), "xl/tables/table1.xml")
    assert table.get("ref") == "A1:D5" and table.find(f"{X}autoFilter").get("ref") == "A1:D5"
    wb.undo()
    assert wb.to_bytes() == original


def test_append_refuses_what_lies_below_the_table():
    wb = Workbook.open(_orders(below={"B7": "note"}))
    before = wb.to_bytes()
    with pytest.raises(EditError) as refused:
        wb.append_rows("t:Orders", [["Chair", 2, 20], ["Shelf", 1, 3]])
    assert refused.value.code == "not_empty" and refused.value.addresses == ("Orders!B7",)
    assert wb.to_bytes() == before
    wb.append_rows("t:Orders", [["Chair", 2, 20]])     # one row fits


def test_append_refuses_to_move_a_totals_row_something_refers_to():
    wb = Workbook.open(_orders(below={"F3": BuiltFormula("D5*2", cached=293)}))
    with pytest.raises(EditError) as refused:
        wb.append_rows("t:Orders", [["Chair", 2, 20]])
    assert refused.value.code == "totals_referenced"
    assert refused.value.addresses == ("Orders!F3",)
    whole = Workbook.open(_orders(below={"F3": BuiltFormula("SUM(D:D)", cached=1)}))
    whole.append_rows("t:Orders", [["Chair", 2, 20]])     # a whole column follows


def test_append_refuses_a_totals_formula_by_address():
    book = WorkbookBuilder()
    sheet = book.sheet("Orders")
    sheet.rows("A1", [["Item", "Qty"], ["Lamp", 3], ["Total"]])
    sheet.cell("B3", BuiltFormula("SUM(B2:B2)", cached=3))
    sheet.table("Orders", "A1:B3", ["Item", "Qty"], write_header=False,
                totals_labels={"Item": "Total"})
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(EditError) as refused:
        wb.append_rows("t:Orders", [["Fan", 1]])
    assert refused.value.code == "totals_referenced" and refused.value.addresses == (
        "Orders!B3",)


def test_append_below_the_last_value_in_a_sheet_s_columns():
    book = WorkbookBuilder()
    plain = book.style(num_fmt="0.00")
    log = book.sheet("Log")
    log.rows("A1", [["When", "What"], ["2026-11-01", "start"]])
    log.cell("C9", "aside")                           # another column: not counted
    log.cell("A5", None, style=plain)                 # a formatted empty cell: not counted
    wb = Workbook.open(book.to_bytes())
    done = wb.append_rows("Log", [["2026-11-02", "next"], ["2026-11-03"]])
    assert done.range == "Log!A3:B4"
    done = wb.append_rows("Log!B:C", [["later"]])     # a region: any of its columns
    assert done.range == "Log!B10"
    with pytest.raises(EditError) as refused:      # a region: its rows only are read,
        wb.append_rows("Log!A1:B2", [["x", "y"]])   # and what lies below is never written over
    assert refused.value.code == "not_empty"


def test_append_to_an_empty_sheet_starts_at_row_one():
    wb = Workbook.new(["Log"])
    assert wb.append_rows("Log", [["a", "b"]]).range == "Log!A1:B1"
    assert wb.append_rows("Log!C5:D9", [["c", "d"]]).range == "Log!C5:D5"


def test_refusal_an_ambiguous_or_unknown_target_names_the_forms():
    wb = Workbook.from_tables({"Orders": {"columns": ["Item"], "rows": [["x"]],
                                          "table": "Orders"}})
    with pytest.raises(EditError) as refused:
        wb.append_rows("Orders", [["y"]])
    assert "t:Orders" in str(refused.value) and "Orders!" in str(refused.value)
    assert wb.append_rows("t:Orders", [["y"]]).table == "Orders"
    assert wb.append_rows("Orders!", [["z"]]).range == "Orders!A4"
    with pytest.raises(EditError) as refused:
        wb.append_rows("Nowhere", [["y"]])
    assert "the sheets" in str(refused.value)


# -- the tool -------------------------------------------------------------------------------------


@pytest.fixture()
def box():
    with Toolbox([*TOOLS, NEW_DOCUMENT], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def test_the_tool_writes_a_table_and_appends_to_it(box):
    session = box.session()
    made = box.dispatch(session, "new_document", {"kind": "xlsx", "name": "stock.xlsx"})
    assert made.ok and made.created == ["d1"]
    done = box.dispatch(session, "xlsx_write_table", {
        "doc": "d1", "at": "Sheet1!B2", "headers": ["Item", "Qty", "Added"], "table": True,
        "name": "Stock", "rows": [[{"text": "Lamp"}, {"number": 3.0}, {"date": "2026-01-05"}],
                                  [{"text": "Fan"}, {}, {"blank": True}]]})
    assert done.ok and done.data["range"] == "Sheet1!B2:D4" and done.data["table"] == "Stock"
    more = box.dispatch(session, "xlsx_write_table", {
        "doc": "d1", "at": "t:Stock", "headers": ["ignored"],
        "rows": [[{"text": "Desk"}, {"number": 1}]]})
    assert more.ok and more.data["table_range"] == "Sheet1!B2:D5" and more.warnings
    refused = box.dispatch(session, "xlsx_write_table", {
        "doc": "d1", "at": "Sheet1!C3", "rows": [[{"text": "x"}]]})
    assert not refused.ok and refused.error.details["code"] == "table_overlap"


def test_the_tool_writes_a_dataset(box):
    session = box.session()
    blob = session.add_blob(_orders(), name="orders.xlsx")
    box.dispatch(session, "open_document", {"blob": blob})
    derived = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Orders", "id": "few", "group_by": ["Item"],
        "measures": [{"fn": "sum", "column": "Qty", "as": "Qty"}]})
    assert derived.ok, derived.error
    done = box.dispatch(session, "xlsx_write_table", {"doc": "d1", "at": "Orders!H1",
                                                      "from": "d:few", "table": True})
    assert done.ok, done.error
    assert done.data["range"] == "Orders!H1:I4"


def test_the_tool_caps_cells(box):
    session = box.session()
    box.dispatch(session, "new_document", {"kind": "xlsx"})
    refused = box.dispatch(session, "xlsx_write_table", {
        "doc": "d1", "at": "Sheet1!A1", "rows": [[{"number": 1}] * 10] * 501})
    assert not refused.ok and refused.error.code == "limit"


def test_new_document_refuses_a_size_and_takes_a_template(box):
    session = box.session()
    refused = box.dispatch(session, "new_document", {"kind": "xlsx", "size": "A4"})
    assert not refused.ok and refused.error.field == "size"
    template = Workbook.from_tables({"Stock": {"columns": ["Item"], "rows": [["Lamp"]]}})
    blob = session.add_blob(template.to_bytes("xltx"), name="stock.xltx")
    made = box.dispatch(session, "new_document", {"kind": "xlsx", "template_blob": blob})
    assert made.ok and made.data["sheets"] == ["Stock"]
    saved = box.dispatch(session, "save_document", {"doc": made.created[0], "name": "a.xlsx",
                                                    "format": "xlsx"})
    assert saved.ok, saved.error
