"""A visible cell's formula that reaches a column hidden from the model is not given as
written: its text is withheld (``formula_withheld``) and its cached value kept; a hidden
column is hidden from its header row down, below its table's rows too; and
``open_document`` counts the declared tables (:mod:`xlsx_agent.declare`, "Hidden from the
model").  Every workbook is made up here: a log of checks in an Excel table, its note column
declared hidden, and formulas over it on the sheet and on another."""

from __future__ import annotations

import json

import pytest
from build import Formula, WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import DeclarationError, RegionError, Workbook
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

COLUMNS = ["Ref", "Item", "Owner", "Note"]
ROWS = [["R-1", "pump", "Kim", "private alpha"],
        ["R-2", "valve", "Lee", "private beta"],
        ["R-3", "hose", "Kim", "private gamma"]]
CHECKS = {"name": "checks", "sheets": "Log", "key": "Ref",
          "fields": [{"header": "Note", "hidden": True}]}

#: Formulas on another sheet: (text, whether it reaches the hidden column).
CALC = [
    ("Log!D2", True),
    ("'Log'!$D$2:$D$4", True),
    ("SUM(Log!D:D)", True),
    ("COUNTA(ChecksTable[Note])", True),
    ("COUNTA(checkstable[[Item]:[Note]])", True),
    ("COUNTA(ChecksTable[[#Headers],[Note]])", True),
    ("ROWS(ChecksTable)", True),
    ("COUNTA(ChecksTable[#All])", True),
    ("COUNTA(NoteCells)", True),
    ("NoteCount", True),
    ("INDIRECT(\"Log!B2\")", True),
    ("SUM(Log!2:2)", True),
    ("COUNTA(Log!C2:E4)", True),
    ("Log!D9", True),                 # below the table's rows
    ("Log!D1", True),                 # the hidden header
    ("COUNTA(Log!A1:B4 Log!D1:D4)", True),
    ("Log!C2", False),
    ("SUM(Log!A2:C4)", False),
    ("COUNTA(ChecksTable[Item])", False),
    ("COUNTA(ChecksTable[[#This Row],[Owner]])", False),
    ("COUNTA(Items)", False),
    ("\"Log!D2\"&\"ChecksTable[Note]\"", False),   # text, not references
    ("Other!D2", False),
    ("[1]Log!D2", False),             # another workbook's sheet
    ("LET(x,1,x+1)", False),
]


def book() -> WorkbookBuilder:
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [COLUMNS] + ROWS)
    sheet.table("ChecksTable", "A1:D4", COLUMNS, write_header=False)
    sheet.cell("D9", "private delta")                         # below the table, hidden column
    sheet.cell("E2", Formula("D2", cached="private alpha"))   # a copy of a hidden cell
    sheet.cell("E3", Formula("B3", cached="valve"))
    # A shared group whose members reach the hidden column (I2 is =D2), and one that doesn't.
    sheet.cell("G2", Formula("B2", cached="pump", kind="shared", ref="G2:I2", si=0))
    sheet.cell("H2", Formula("", cached="Kim", kind="shared", si=0))
    sheet.cell("I2", Formula("", cached="private alpha", kind="shared", si=0))
    sheet.cell("G3", Formula("B3", cached="valve", kind="shared", ref="G3:G4", si=1))
    sheet.cell("G4", Formula("", cached="hose", kind="shared", si=1))
    calc = built.sheet("Calc")
    for row, (text, _) in enumerate(CALC, start=1):
        calc.cell(f"A{row}", Formula(text, cached=row))
    built.sheet("Other").rows("A1", [["Code", "Count", "x", "Spare"], ["x", 1, 2, 3]])
    built.define_name("NoteCells", "Log!$D$2:$D$4")
    built.define_name("NoteCount", "COUNTA(Log!$D$2:$D$4)")
    built.define_name("Items", "Log!$B$2:$B$4")
    return built


def opened() -> Workbook:
    wb = Workbook.open(book().to_bytes())
    wb.declare_tables([CHECKS])
    return wb


# -- formula text ---------------------------------------------------------------------------------


@pytest.mark.parametrize("row", range(1, len(CALC) + 1))
def test_a_formula_reaching_a_hidden_cell_is_withheld_its_value_kept(row):
    text, reaches = CALC[row - 1]
    value = opened().cell(f"Calc!A{row}")
    assert value.formula and value.v == row
    if reaches:
        assert (value.formula_text, value.formula_withheld) == (None, True), text
    else:
        assert (value.formula_text, value.formula_withheld) == (text, False), text


def test_without_a_declaration_every_formula_is_given():
    wb = Workbook.open(book().to_bytes())
    assert [v.formula_text for row in wb.read("Calc!A1:A25") for v in row] == \
        [text for text, _ in CALC]


def test_a_formula_on_the_hidden_columns_sheet_copying_a_hidden_cell():
    wb = opened()
    copy, other = wb.cell("Log!E2"), wb.cell("Log!E3")
    assert (copy.v, copy.formula_text, copy.formula_withheld) == ("private alpha", None, True)
    assert (other.formula_text, other.formula_withheld) == ("B3", False)


def test_a_shared_group_reaching_a_hidden_cell_is_withheld_whole():
    wb = opened()
    row = next(iter(wb.read("Log!G2:I2")))
    assert [(v.formula_text, v.formula_withheld) for v in row] == [(None, True)] * 3
    assert [v.v for v in row] == ["pump", "Kim", "private alpha"]
    kept = [v for r in wb.read("Log!G3:G4") for v in r]
    assert [(v.formula_text, v.shared_from, v.formula_withheld) for v in kept] == [
        ("B3", None, False), (None, "Log!G3", False)]


def test_a_formula_in_the_table_naming_its_hidden_column_by_this_row():
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [["Ref", "Note", "Copy", "Again"]])
    for row in (2, 3):
        sheet.rows(f"A{row}", [[f"R-{row}", f"private {row}"]])
        sheet.cell(f"C{row}", Formula("[@Note]", cached=f"private {row}"))
        sheet.cell(f"D{row}", Formula("Copies[@Ref]", cached=f"R-{row}"))
    sheet.cell("F2", Formula("Copies[@Note]", cached="private 2"))
    sheet.cell("F3", Formula("Copies[[#This Row],[Note]]", cached="private 3"))
    sheet.table("Copies", "A1:D3", ["Ref", "Note", "Copy", "Again"], write_header=False)
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([{"name": "copies", "sheets": "Log", "key": "Ref",
                        "fields": [{"header": "Note", "hidden": True}]}])
    cells = {v.ref: v for row in wb.read("Log!A2:F3") for v in row}
    assert [cells[r].formula_withheld for r in ("C2", "C3", "F2", "F3")] == [True] * 4
    assert [cells[r].formula_text for r in ("D2", "D3")] == ["Copies[@Ref]"] * 2
    through = [v for row in wb.read("t:copies") for v in row]
    assert [(v.ref, v.formula_withheld) for v in through if v.formula] == [
        ("C2", True), ("D2", False), ("C3", True), ("D3", False)]


def test_a_name_reaching_a_name_over_hidden_cells():
    built = book()
    built.define_name("Twice", "NoteCount*2")
    built.sheet("More").cell("A1", Formula("Twice", cached=6))
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([CHECKS])
    assert wb.cell("More!A1").formula_withheld


# -- the hidden column below its table -----------------------------------------------------------


def test_the_hidden_column_is_hidden_below_its_tables_rows():
    wb = opened()
    assert wb.cell("Log!D9").t == "blank"
    assert list(wb.find_text("delta")) == []
    built = book()
    built.define_name("Below", "Log!$D$9")
    named = Workbook.open(built.to_bytes())
    named.declare_tables([CHECKS])
    with pytest.raises(RegionError) as caught:
        named.region("t:Below")
    assert caught.value.code == "declared"
    with pytest.raises(RegionError) as caught:
        wb.declare_region("below", "Log!C8:D9", header_rows=0)
    assert caught.value.code == "declared"


def test_the_hidden_column_above_its_header_stays_visible():
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [["Run", None, "Season", "spring"]])
    sheet.rows("A3", [COLUMNS] + ROWS)
    wb = Workbook.open(built.to_bytes())
    wb.declare_tables([{"name": "checks", "sheets": "Log", "header_row": 3, "key": "Ref",
                        "fields": [{"header": "Note", "hidden": True}]},
                       {"name": "run", "type": "key_value", "sheets": "Log",
                        "columns": "C:D", "first_row": 1, "last_row": 1}])
    assert wb.cell("Log!D1").v == "spring"
    assert wb.cell("Log!D4").t == "blank"


def test_another_table_below_showing_the_hidden_column_is_refused():
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    sheet.rows("A1", [COLUMNS] + ROWS)
    sheet.rows("C8", [["Code", "Count"], ["x", 1], ["y", 2]])
    with pytest.raises(DeclarationError) as caught:
        Workbook.open(built.to_bytes()).declare_tables(
            [CHECKS, {"name": "counts", "sheets": "Log", "header_row": 8, "columns": "C:D"}])
    assert caught.value.code == "hidden_elsewhere"


# -- the tools -----------------------------------------------------------------------------------


def toolbox_session():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[CHECKS])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(book().to_bytes(), name="log.xlsx")
    return toolbox, session, blob


def test_xlsx_read_with_formulas_marks_the_withheld_ones():
    toolbox, session, blob = toolbox_session()
    opened = toolbox.dispatch(session, "open_document", {"blob": blob})
    assert opened.data["summary"]["tables"] == 1     # the declared table, the Excel one replaced
    read = toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Calc!A1:A25",
                                                   "formulas": True}).to_json()
    cells = [c for row in read["data"]["rows"] for c in row]
    for cell, (text, reaches) in zip(cells, CALC):
        if reaches:
            assert "f" not in cell and cell["formula_withheld"] is True, text
        else:
            assert cell["f"] == text and "formula_withheld" not in cell, text
    plain = toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Log!E2:E3"})
    assert "formula_withheld" not in json.dumps(plain.to_json())   # only when f is asked for
    log = toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "Log!A1:I9",
                                                  "formulas": True}).to_json()
    text = json.dumps(log, ensure_ascii=False)
    assert "Note" not in text and "delta" not in text and '"D2"' not in text


def test_open_document_counts_declared_tables():
    toolbox = Toolbox(TOOLS, formats=[xlsx_format(tables=[
        CHECKS, {"name": "spare", "sheets": "Other"}])], groups=GROUPS)
    session = toolbox.session()
    blob = session.add_blob(book().to_bytes(), name="log.xlsx")
    assert toolbox.dispatch(session, "open_document",
                            {"blob": blob}).data["summary"]["tables"] == 2
    plain = Toolbox(TOOLS, formats=[xlsx_format()], groups=GROUPS)
    other = plain.session()
    blob = other.add_blob(book().to_bytes(), name="log.xlsx")
    assert plain.dispatch(other, "open_document", {"blob": blob}).data["summary"]["tables"] == 1
