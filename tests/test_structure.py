"""Inserting, deleting and moving rows and columns (round 44): held to what Excel did.

Every case in ``golden/structure`` is one operation on a made-up workbook
(``structure_book.py``: ``StructureBook``, ``GridBook``, ``FormatBook``) as Excel for Mac
16.106 did it (2026-11-08, ``tools/excel_structure.applescript`` through ``oracle.py``'s
``Session.operate``): the facts of the file Excel saved (``structure_facts.measured``: every
formula as text, defined names, chart series, pivot sources, rows, columns, merges, rules
per cell, hyperlinks, comments, filters, panes, tables; no cached value), or that Excel
refused it.  Here the same operation is done by xlsx-agent, offline, and must give the same
facts -- or be refused when Excel refused.  The workbook's language is Dutch, as the Excel
measured was (a new table column is ``Kolom1``).  ``test_structure_oracle.py`` asks Excel
again (``pytest -m oracle``).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from ooxml_common.locales import Locale

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError

import structure_book
from structure_facts import differences, facts, measured

GOLDEN = Path(__file__).parent / "golden" / "structure"


def cases():
    for path in sorted(GOLDEN.glob("*.json")):
        data = json.loads(path.read_text())
        for name, case in sorted(data["cases"].items()):
            yield pytest.param(data["book"], case, id=f"{path.stem}-{name}")


def operate(book: str, case: dict) -> bytes:
    wb = Workbook.open(getattr(structure_book, book)(**case["variant"]).to_bytes())
    wb.locale = Locale(language="nl-NL")
    for step in case["steps"]:
        kind, sheet, lines, before = step.split("|")
        if kind == "move":
            move = wb.move_rows if lines[0].isdigit() else wb.move_columns
            move(sheet, lines, before.split(":")[0])
        else:
            getattr(wb, kind)(sheet, lines)
    return wb.to_bytes()


@pytest.mark.parametrize("book, case", list(cases()))
def test_as_excel_did(book, case):
    if case["excel"] == "refused":
        with pytest.raises(EditError):
            operate(book, case)
        return
    ours = measured(facts(operate(book, case)))
    assert differences(ours, case["facts"]) == []


# -- the step: undo, refusals, reports (tool design rules 2, 3, 5, 6) ----------------------------


def _book(**options) -> Workbook:
    return Workbook.open(structure_book.StructureBook(**options).to_bytes())


def test_one_step_undone_to_the_bytes_opened():
    original = structure_book.StructureBook().to_bytes()
    wb = Workbook.open(original)
    before = wb.version
    report = wb.insert_rows("Stock", 4, 2)
    assert report.changed and wb.version != before
    edited = wb.to_bytes()
    assert wb.undo() and wb.to_bytes() == original
    assert wb.redo() and wb.to_bytes() == edited


def test_save_reads_back_and_checks_the_structure(tmp_path):
    wb = _book()
    wb.set_cells({"Stock!B5": 19})
    wb.move_rows("Stock", 5, 2)
    report = wb.save(tmp_path / "moved.xlsx")
    assert report.issues == ()
    assert wb.validate() == []
    # the cell written before the move is read back where it went
    assert Workbook.open(tmp_path / "moved.xlsx").cell("Stock!B2").v == 19


def test_refusal_lists_every_problem_and_changes_nothing():
    wb = _book()
    original = wb.to_bytes()
    with pytest.raises(EditError) as raised:
        wb.move_rows("Stock", "2:3", 10)          # out of a table, an array cut
    error = raised.value
    codes = {code for code, _ in error.problems}
    assert "table_move" in codes
    assert "t:Goods" in str(error)               # rule 2: where, and why
    assert wb.to_bytes() == original and wb.version == 0


@pytest.mark.parametrize("call, code", [
    (lambda wb: wb.insert_rows("Stock", 3), "array_split"),
    (lambda wb: wb.delete_rows("Stock", 1), "table_header_row"),
    (lambda wb: wb.move_columns("Stock", "E", "B"), "table_move"),
    (lambda wb: wb.move_rows("Stock", 11, 2), "merge_split"),
    (lambda wb: wb.insert_rows("Notes", 5), "off_grid"),
])
def test_refused_as_excel_refuses(call, code):
    wb = _book(edge=True)
    with pytest.raises(EditError) as raised:
        call(wb)
    assert raised.value.code == code


def test_bad_positions_say_what_is_taken():
    wb = _book()
    for call in (lambda: wb.insert_rows("Stock", 0), lambda: wb.insert_columns("Stock", "C5"),
                 lambda: wb.delete_rows("Stock", "4:5", 3), lambda: wb.move_rows("Stock", 4, 0)):
        with pytest.raises(EditError) as raised:
            call()
        assert raised.value.code == "bad_position"
    with pytest.raises(EditError) as raised:
        wb.insert_rows("Nowhere", 2)
    assert raised.value.code == "not_a_worksheet"


def test_a_move_to_where_the_lines_are_is_no_change():
    wb = _book()
    report = wb.move_columns("Stock", "J", "K")        # rule 3: already there
    assert not report.changed and report.notes and wb.version == 0


def test_report_names_what_broke_by_its_new_address():
    report = _book().delete_rows("Stock", 4, 2)
    assert "Orders!B5" in report.broken and "name:Local" in report.broken
    assert "Stock!H2" in report.broken                 # H2 = B4, now #REF!
    assert report.removed == {"comments": 1}
    assert report.calculation == "invalidated"


def test_inserted_lines_take_the_template_format():
    # rule 6: the neighbour's height, style and rules, never its hiding
    wb = _book()
    wb.insert_rows("Stock", 5)                         # under the tall row 4
    wb.insert_rows("Stock", 11)                        # under the styled row (now 10)
    wb.insert_columns("Stock", "D")                    # right of the styled column C
    found = facts(wb.to_bytes())["sheets"]["Stock"]
    assert found["rows"][5]["ht"] == "30" and found["rows"][11] == found["rows"][10]
    columns = {c["min"]: c for c in found["cols"]}
    assert columns["3"].get("style") and columns["3"]["max"] == "3"
    assert any(c["min"] <= "4" <= c["max"] and c.get("style") for c in found["cols"])


def test_rearrange_columns_in_one_step():
    wb = _book()
    before = wb.version
    done = wb.rearrange_columns("Stock", ["K", "J"], drop=["I"])
    assert done.changed and done.dropped == ("I",)
    assert done.columns == {"K": "I", "J": "J"}
    assert wb.undo() and wb.version == before          # one undo step (rule 5)


def test_rearrange_by_mapping_and_no_op():
    wb = _book()
    done = wb.rearrange_columns("Stock", {"K": "J", "J": "K"})
    assert done.columns == {"K": "J", "J": "K"}
    cells = facts(wb.to_bytes())["sheets"]["Stock"]["cells"]
    assert cells["J1"][2] == "Count" and cells["K1"][2] == "Code"
    assert cells["H1"][0] == "SUM(B2:B7)"
    again = wb.rearrange_columns("Stock", ["J", "K"])
    assert not again.changed


def test_rearrange_refuses_a_tables_edge_and_a_bad_order():
    wb = _book()
    with pytest.raises(EditError) as raised:
        wb.rearrange_columns("Stock", ["J", "A"])        # into and out of t:Goods
    assert raised.value.code == "table_move" and wb.version == 0
    for order, drop in ((["J", "J"], ()), ({"J": "A"}, ()), (["J"], ["J"]), (None, ())):
        with pytest.raises(EditError) as raised:
            wb.rearrange_columns("Stock", order, drop=drop)
        assert raised.value.code == "bad_position"


def test_inside_a_batch_with_writes():
    wb = _book()
    with wb.batch():
        wb.set_cells({"Stock!K2": 40})
        wb.insert_rows("Stock", 2)
        wb.set_cells({"Stock!K2": 1})
    cells = facts(wb.to_bytes())["sheets"]["Stock"]["cells"]
    assert cells["K3"][2] == 40.0 and cells["K2"][2] == 1.0
    assert wb.undo() and wb.version == 0
