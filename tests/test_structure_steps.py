"""Round 44 lane P: each edit step pays for what it touched, and finds what the long way finds.

A step's structure check reads only the rows the step wrote (``check_structure``'s
``rows``); its guards read a sheet's merges, array blocks and validation rules from the
tree the cell writer parsed (``Workbook._layout``), and a block's cells from that tree
(:mod:`xlsx_agent.edit.content`), never typing the sheet's values.  Each shortcut is
checked here against the long way round -- the whole check, the streamed scan, the typed
cells -- on the test workbooks, with random edits and with damage put where a step wrote.
The timings themselves are tools/bench.py's; one generous bound runs with ``--run-slow``.
"""

from __future__ import annotations

import random
import time

import pytest

from build import Formula, WorkbookBuilder
from ooxml_edit.sheets import Area, Cell

from shop import Shop
from structure_book import StructureBook
from xlsx_agent import Workbook
from xlsx_agent.edit.content import with_content
from xlsx_agent.edit.structure import check_structure
from xlsx_agent.read import stream
from xlsx_agent.read.stream import SheetLayout, layout

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

BOOKS = {
    "shop": lambda: Shop().to_bytes(),
    "shop-rless": lambda: Shop(rless=True).to_bytes(),
    "shop-inline": lambda: Shop(shared=False).to_bytes(),
    "shop-orders": lambda: Shop(orders=True, table=True, array=True).to_bytes(),
    "structure": lambda: StructureBook().to_bytes(),
}

VALUES = ["text", "Spare fan", 7, 2.5, True, None, {"date": "2026-11-08", "format": "yyyy-mm-dd"},
          {"formula": "=1+1"}, {"markdown": "**bold** text"}]


def _writable(wb: Workbook) -> list[tuple[str, str]]:
    """``(sheet name, address)`` of cells a plain write is never refused at: outside merges,
    arrays, tables' headers and formulas, on unprotected worksheets."""
    out = []
    for sheet in wb.sheets:
        if sheet.type != "worksheet" or sheet.protected:
            continue
        taken = list(wb.merges(sheet)) + list(wb.arrays(sheet))
        taken += [t.area for t in wb._table_parts() if t.sheet == sheet.name
                  and t.area is not None]
        formulas = {(v.row, v.column) for v in wb.cells(sheet) if v.formula}
        name = sheet.name.replace("'", "''")
        for row in range(1, 30):
            for column in range(1, 12):
                if (row, column) in formulas or any(
                        a.first.row <= row <= a.last.row
                        and a.first.column <= column <= a.last.column for a in taken):
                    continue
                out.append((sheet.name, f"'{name}'!{chr(64 + column)}{row}"))
    return out


def _step_rows(wb: Workbook) -> dict[str, set[int]]:
    """The rows the last step's check read, by sheet part, as the editor kept them."""
    report = wb.last_edit
    assert report is not None
    rows: dict[str, set[int]] = {}
    for label, ranges in report.cells.items():
        sheet = wb.sheet(label.strip("'").replace("''", "'"))
        for text in ranges:
            area = Area.parse(text)
            rows.setdefault(sheet.part, set()).update(range(area.first.row, area.last.row + 1))
    return rows


# -- the step's check against the whole one ----------------------------------------------------


@pytest.mark.parametrize("book", BOOKS.values(), ids=BOOKS.keys())
def test_random_edits_report_what_the_whole_check_finds(book):
    """On a workbook with no issue, the step's check (its rows) and the whole check of the
    sheets it wrote agree after every step: both empty, or the same issues."""
    rng = random.Random(44)
    wb = Workbook.open(book())
    cells = _writable(wb)
    for _ in range(12):
        chosen = rng.sample(cells, rng.randint(1, 4))
        wb.set_cells({address: rng.choice(VALUES) for _, address in chosen},
                     replace_formulas=True)
        report = wb.last_edit
        sheets = sorted({name for name, _ in chosen})
        assert list(report.issues) == wb.validate(sheets=sheets)
        assert list(report.issues) == check_structure(
            wb, sheets=[wb.sheet(n).part for n in sheets])


#: Damage a broken writer could leave at a cell it wrote: ``(name, change of the c element)``.
DAMAGE = {
    "style": lambda c: c.set("s", "999"),
    "shared_string": lambda c: (c.set("t", "s"), _value(c, "99999")),
    "cell_row": lambda c: c.set("r", c.get("r").rstrip("0123456789") + "1048000"),
    "bad_reference": lambda c: c.set("r", "A0"),
    "cell_order": lambda c: c.addnext(_copy(c)),
}


def _value(cell, text):
    for child in list(cell):
        cell.remove(child)
    v = cell.makeelement(X + "v")
    v.text = text
    cell.append(v)


def _copy(cell):
    from copy import deepcopy

    return deepcopy(cell)


@pytest.mark.parametrize("damage", DAMAGE.values(), ids=DAMAGE.keys())
@pytest.mark.parametrize("book", BOOKS.values(), ids=BOOKS.keys())
def test_damage_where_a_step_wrote_is_found_as_the_whole_check_finds_it(book, damage):
    rng = random.Random(hash(damage) & 0xFFFF)
    wb = Workbook.open(book())
    cells = _writable(wb)
    for _ in range(4):
        name, address = rng.choice(cells)
        with wb.batch():
            wb.set_cells({address: "written"}, replace_formulas=True)
            sheet = wb.sheet(name)
            root = wb.package.tree(sheet.part)
            ref = address.rpartition("!")[2]
            node = next(c for c in root.iter(X + "c") if c.get("r") == ref)
            damage(node)
            wb.package.mark_dirty(sheet.part)
        report = wb.last_edit
        whole = wb.validate(sheets=[name])
        assert whole, "the damage is an issue"
        assert list(report.issues) == whole
        rows = check_structure(wb, sheets=[sheet.part], rows=_step_rows(wb))
        assert rows == whole
        wb.undo()


def test_the_dimension_is_held_to_the_rows_written():
    wb = Workbook.open(Shop().to_bytes())
    with wb.batch():
        wb.set_cells({"Stock!B30": "far below"})
        root = wb.package.tree(wb.sheet("Stock").part)
        root.find(X + "dimension").set("ref", "A1:I12")
        wb.package.mark_dirty(wb.sheet("Stock").part)
    assert [p.code for p in wb.last_edit.issues] == ["dimension"]
    assert wb.last_edit.issues == tuple(wb.validate(sheets=["Stock"]))


def test_damage_away_from_the_rows_written_is_the_whole_checks_alone():
    """The step's check reads the rows it wrote (and every row as a row); an issue in a
    cell of another row is validate()'s to find, and is."""
    wb = Workbook.open(Shop().to_bytes())
    with wb.batch():
        wb.set_cells({"Stock!B12": "written"})
        root = wb.package.tree(wb.sheet("Stock").part)
        node = next(c for c in root.iter(X + "c") if c.get("r") == "D10")
        node.set("s", "999")
        wb.package.mark_dirty(wb.sheet("Stock").part)
    assert wb.last_edit.issues == ()
    assert [(p.code, p.address) for p in wb.validate()] == [("style", "Stock!D10")]


def test_a_row_out_of_order_anywhere_is_found_by_the_step():
    """Rows are checked as rows across the whole sheet, whatever the step wrote."""
    wb = Workbook.open(Shop().to_bytes())
    with wb.batch():
        wb.set_cells({"Stock!B12": "written"})
        root = wb.package.tree(wb.sheet("Stock").part)
        rows = root.find(X + "sheetData")
        rows[2].set("r", "1")
        wb.package.mark_dirty(wb.sheet("Stock").part)
    assert [p.code for p in wb.last_edit.issues] == ["row_order"]
    # the whole check also places the row's cells against it (cell_row): cells of a row
    # the step did not write
    assert wb.last_edit.issues[0] in wb.validate(sheets=["Stock"])


def test_a_whole_sheet_change_is_checked_whole():
    """Formatting, layout and rows inserted change more than cells: the sheet is checked
    whole, and an issue anywhere on it is in the report."""
    wb = Workbook.open(StructureBook().to_bytes())
    with wb.batch():
        root = wb.package.tree(wb.sheet("Stock").part)
        node = next(c for c in root.iter(X + "c") if c.get("r") == "B5")
        node.set("s", "999")
        wb.insert_rows("Stock", 20)
    assert ("style", "Stock!B5") in [(p.code, p.address) for p in wb.last_edit.issues]


# -- a validate from the tree and from the bytes ----------------------------------------------


@pytest.mark.parametrize("book", BOOKS.values(), ids=BOOKS.keys())
def test_validate_reads_a_parsed_sheet_as_it_streams_its_bytes(book):
    rng = random.Random(7)
    wb = Workbook.open(book())
    cells = _writable(wb)
    wb.set_cells({address: rng.choice(VALUES) for _, address in rng.sample(cells, 5)})
    for sheet in wb.sheets:
        if sheet.part is not None:
            root = wb.package.tree(sheet.part)
            for node in list(root.iter(X + "c"))[::7]:
                node.set("s", "999")
            wb.package.mark_dirty(sheet.part)
    from_trees = wb.validate()
    assert from_trees
    assert Workbook.open(wb.to_bytes()).validate() == from_trees


# -- the guards' layout and a block's content --------------------------------------------------


@pytest.mark.parametrize("book", BOOKS.values(), ids=BOOKS.keys())
def test_the_layout_from_the_tree_is_the_scans(book):
    wb = Workbook.open(book())
    for sheet in wb.sheets:
        if sheet.part is None or sheet.type not in ("worksheet", "macrosheet"):
            continue
        scanned = SheetLayout.of_scan(wb.scan(sheet))
        assert layout(wb.package.tree(sheet.part)) == scanned
        fresh = Workbook.open(book())
        assert fresh._layout(sheet.name, edit=True) == scanned


def test_an_array_formula_without_a_reference_is_left_to_the_scan():
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    sheet.cell("A1", 1)
    parts = book.parts()
    part = "xl/worksheets/sheet1.xml"
    parts[part] = parts[part].replace(
        b"</sheetData>", b'<row r="3"><c><f t="array" ref="A3:A4">1</f><v>1</v></c></row>'
                         b"</sheetData>")
    from build import zip_parts

    wb = Workbook.open(zip_parts(parts))
    assert layout(wb.package.tree(wb.sheet("Data").part)) is None
    assert wb._layout("Data", edit=True).arrays == (Area.parse("A3:A4"),)


def _typed(wb: Workbook, sheet, area: Area) -> list[tuple[int, int]]:
    """The long way: every cell of the sheet typed, those inside ``area`` holding
    something."""
    first, last = area.first, area.last
    return sorted((v.row, v.column) for v in wb.cells(sheet)
                  if first.row <= v.row <= last.row and first.column <= v.column <= last.column
                  and (v.formula or v.t != "blank"))


@pytest.mark.parametrize("book", BOOKS.values(), ids=BOOKS.keys())
def test_a_blocks_content_is_what_typing_every_cell_finds(book):
    rng = random.Random(3)
    wb = Workbook.open(book())
    for sheet in wb.sheets:
        if sheet.part is None or sheet.type != "worksheet":
            continue
        for _ in range(25):
            top, left = rng.randint(1, 20), rng.randint(1, 10)
            area = Area(Cell(top, left), Cell(top + rng.randint(0, 30), left + rng.randint(0, 8)))
            expected = _typed(wb, sheet, area)
            assert with_content(wb, sheet, area) == expected
            assert with_content(wb, sheet, area, last=True) == expected[-1:]


def test_odd_cells_hold_what_typing_says():
    """Empty values, booleans that are not, a shared string past the table, an error, a
    formula without a value: each holds something or not as the streamed pass types it."""
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    sheet.cell("A1", "s0")
    parts = book.parts()
    part = "xl/worksheets/sheet1.xml"
    odd = (b'<row r="2"><c r="A2"><v></v></c><c r="B2" t="b"><v>maybe</v></c>'
           b'<c r="C2" t="s"><v>99</v></c><c r="D2" t="e"><v>#N/A</v></c>'
           b'<c r="E2"><f>1+1</f></c><c r="F2" t="str"><v></v></c>'
           b'<c r="G2" t="inlineStr"/><c r="H2" s="0"/><c r="I2" t="n"><v>abc</v></c></row>')
    parts[part] = parts[part].replace(b"</sheetData>", odd + b"</sheetData>")
    from build import zip_parts

    wb = Workbook.open(zip_parts(parts))
    data = wb.sheet("Data")
    area = Area(Cell(1, 1), Cell(2, 9))
    assert with_content(wb, data, area) == _typed(wb, data, area)


# -- nothing typed, the whole sheet never streamed ---------------------------------------------


def _big(rows: int) -> bytes:
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    sheet.rows("A1", [["Name", "Qty", "Price", "Flag"]]
               + [[f"item {n}", n, n * 0.5, "x"] for n in range(1, rows + 1)])
    for n in range(1, rows // 10 + 1):
        sheet.cell(f"F{n * 10}", Formula(f"B{n * 10}*C{n * 10}", cached=1))
    return book.to_bytes()


STEPS = {
    "set_cells": lambda wb: wb.set_cells({"Data!A2": "changed", "Data!B300": 3}),
    "append_rows": lambda wb: wb.append_rows("Data", [["new", 1, 2, "y"]] * 20),
    "write_table": lambda wb: wb.write_table("Data", "H1", ["A", "B"], [[1, "t"]] * 20),
    "format": lambda wb: wb.format("Data!C:C", number_format="0.00"),
    "insert_rows": lambda wb: wb.insert_rows("Data", 100, 2),
    "merge": lambda wb: wb.merge("Data!J1:K2"),
}


@pytest.mark.parametrize("step", STEPS.values(), ids=STEPS.keys())
def test_a_step_types_no_value_of_the_sheet(step, monkeypatch):
    """No step streams the sheet to type its values: what a 50,000-row sheet paid for at
    every step.  (The save's read-back and validate() still do, on purpose.)"""
    wb = Workbook.open(_big(2000))

    def refused(*args, **kwargs):
        raise AssertionError("a step typed the sheet's values")

    monkeypatch.setattr(stream, "scan", refused)
    monkeypatch.setattr(stream.SheetReader, "values", refused)
    before = wb.to_bytes()
    report = step(wb)
    assert wb.last_edit.changed and report is not None
    monkeypatch.undo()
    assert wb.validate() == []
    assert wb.undo() and wb.to_bytes() == before


@pytest.mark.slow
def test_steps_on_a_sheet_of_fifty_thousand_rows_are_quick():
    """Generous bounds (CPU seconds), many times what the steps take on a laptop
    (ROADMAP.md, round 44 lane P) and well under what they took before."""
    data = _big(50_000)
    wb = Workbook.open(data)

    def timed(run) -> float:
        start = time.process_time()
        run()
        return time.process_time() - start

    assert timed(lambda: wb.set_cells({"Data!A2": "changed"})) < 10
    assert timed(lambda: wb.set_cells({"Data!A3": "changed"})) < 4
    assert timed(lambda: wb.undo()) < 3
    assert timed(lambda: wb.validate()) < 10
    assert timed(lambda: wb.to_bytes()) < 3
