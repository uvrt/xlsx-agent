"""A table's columns moved among themselves (round 46, R75): held to what Excel did.

Every case in ``golden/table_columns.json`` is one or two moves of whole table columns on
the made-up ``TableColumnsBook`` as Excel for Mac 16.106 did them (2026-11-09: Cut of the
table column's cells and "Insert Cut Cells" before another of its columns, which is what
dragging a table column does; ``tools/excel_cells.applescript`` through ``oracle.py``'s
``Session.operate``): the facts of the file Excel saved (``structure_facts.measured``, and
``table_columns_book.table_facts`` for the table's columns, filter and sort state), or that
Excel refused.  Here the same moves are made by xlsx-agent, offline, and must give the same
facts.  ``pytest -m oracle`` asks Excel again, and has Excel open and re-save ours.

Measured and left out on purpose: Excel also moves a column out of a table by the same
cut (``out-of-table``: the table loses the column, its cells keep the table's look as
cell formats); :meth:`~xlsx_agent.Workbook.rearrange_columns` never moves a column out of
or into a table, and refuses an order that would.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from ooxml_edit.sheets import column_index

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.edit.table_columns import move_table_columns

import oracle
from structure_facts import differences, facts, measured
from table_columns_book import TableColumnsBook, table_facts

GOLDEN = Path(__file__).parent / "golden" / "table_columns.json"
SCRIPT = Path(__file__).resolve().parents[1] / "tools" / "excel_cells.applescript"


def cases():
    data = json.loads(GOLDEN.read_text())
    for name, case in sorted(data["cases"].items()):
        yield pytest.param(case, id=name)


def _columns(text: str) -> tuple[int, int]:
    first, _, last = text.partition(":")
    return (column_index(first.rstrip("0123456789")),
            column_index((last or first).rstrip("0123456789")))


def operate(case: dict) -> bytes:
    """The case's moves made by xlsx-agent: each ``move_right|Shop|F3:F10|C3:C10`` the
    table's columns F..F placed before C."""
    wb = Workbook.open(TableColumnsBook(**case["variant"]).to_bytes())
    sheet = wb._cell_sheet("Shop")
    part = next(rel.target_part for rel in wb.package.relationships(sheet.part).values()
                if rel.type.endswith("/table"))
    for step in case["steps"]:
        _kind, _sheet, source, place = step.split("|")
        first, last = _columns(source)
        move_table_columns(wb, sheet, part, first, last - first + 1, _columns(place)[0])
    return wb.to_bytes()


def found(data: bytes) -> dict:
    return {"facts": measured(facts(data)), "tables": table_facts(data)}


@pytest.mark.parametrize("case", list(cases()))
def test_as_excel_did(case):
    ours = found(operate(case))
    assert differences(ours, {"facts": case["facts"], "tables": case["tables"]}) == []


def test_one_move_is_one_undo_step_back_to_the_bytes():
    original = TableColumnsBook(edges=True, comments=True).to_bytes()
    wb = Workbook.open(original)
    done = wb.rearrange_columns("Shop", ["F", "C", "D", "E"])
    assert done.changed and wb.version == 1
    edited = wb.to_bytes()
    assert wb.undo() and wb.to_bytes() == original
    assert wb.redo() and wb.to_bytes() == edited


def test_the_result_opens_and_checks(tmp_path):
    wb = Workbook.open(TableColumnsBook(edges=True, comments=True).to_bytes())
    wb.rearrange_columns("Shop", {"G": "B", "B": "C", "C": "G"})
    report = wb.save(tmp_path / "shop.xlsx")
    assert report.issues == () and wb.validate() == []
    again = Workbook.open(tmp_path / "shop.xlsx")
    assert [again.cell(f"Shop!{c}3").v for c in "BCDEFG"] == \
        ["Bin", "Code", "Price", "Total", "Note", "Qty"]


def test_moving_into_a_table_is_refused_as_excel_refuses():
    # Excel: Cut H3:H10, Insert Cut Cells at C3 -- a parameter error (refused)
    wb = Workbook.open(TableColumnsBook().to_bytes())
    with pytest.raises(EditError) as raised:
        wb.rearrange_columns("Shop", ["H", "C"])
    assert raised.value.code == "table_move"


# -- the oracle --------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


@pytest.mark.oracle
@pytest.mark.parametrize("case", list(cases()))
def test_excel_agrees(excel, case, request):
    name = request.node.callspec.id
    source = TableColumnsBook(**case["variant"]).to_bytes()
    theirs = excel.operate(source, case["steps"], name=f"table-columns-{name}", script=SCRIPT)
    assert theirs.ok, theirs
    recorded = {"facts": case["facts"], "tables": case["tables"]}
    assert differences(found(theirs.saved.read_bytes()), recorded) == []
    ours = excel.resave(operate(case), name=f"table-columns-ours-{name}")
    assert ours.ok, ours                       # opens without a repair prompt
    assert differences(found(ours.saved.read_bytes()), recorded) == []
