"""Sheet names with leading or trailing spaces (production R69): every call takes the name
exactly as stored.

Excel keeps a sheet name's outer spaces, and ``Stock`` and ``Stock `` (or `` Stock``) are two
sheets of one workbook (measured, Excel for Mac 16.106, 2026-11-09: such workbooks open
without a repair, the object model renames a sheet to ``Stock `` beside ``Stock``, and a
formula refers to ``'Stock '!A1``).  Each public call that takes a sheet or an address is
run against a sheet so named while its twin without the spaces is there too: the call must
act on the sheet named, never the twin -- bare (``"Stock "``) and as an address writes it
(``"'Stock '!A1"``).  A name with its spaces dropped still finds the one sheet it can mean
(tool design rule 1) and is refused, naming both, when it could mean two (rule 2).
"""

from __future__ import annotations

import io
import zipfile

import pytest

from build import WorkbookBuilder
from xlsx_agent import Workbook
from xlsx_agent.address import NameNotFound, quote_sheet
from xlsx_agent.edit import EditError

#: The spaced names swept, each with the twin it must not be taken for.
NAMES = ["Stock ", " Stock", "  Stock  "]
TWIN = "Stock"


def _book(name: str) -> bytes:
    book = WorkbookBuilder()
    for sheet, word in ((TWIN, "twin"), (name, "here")):
        s = book.sheet(sheet)
        s.rows("A1", [["SKU", "Qty", "Bin"], [f"{word}-1", 1, "x"], [f"{word}-2", 2, "y"],
                      [f"{word}-3", 3, "z"]])
    book.sheet("Other").rows("A1", [["n"]])
    return book.to_bytes()


def _parts(wb: Workbook) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(wb.to_bytes())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _part(wb: Workbook, name: str) -> str:
    return next(s.part for s in wb.sheets if s.name == name)


# Each edit, given the bare name ``n`` and the quoted one ``q`` (as an address writes it).
EDITS = {
    "set_cells": lambda wb, n, q: wb.set_cells({f"{q}!E1": 5}),
    "format": lambda wb, n, q: wb.format(f"{q}!A1", font={"bold": True}),
    "set_column_widths": lambda wb, n, q: wb.set_column_widths(f"{q}!A:B", 12),
    "set_row_heights": lambda wb, n, q: wb.set_row_heights(f"{q}!1:2", 20),
    "freeze_panes": lambda wb, n, q: wb.freeze_panes(n, rows=1),
    "set_gridlines": lambda wb, n, q: wb.set_gridlines(n, False),
    "protect_sheet": lambda wb, n, q: wb.protect_sheet(n),
    "unprotect_sheet": lambda wb, n, q: wb.unprotect_sheet(n),
    "insert_rows": lambda wb, n, q: wb.insert_rows(n, 2),
    "insert_columns": lambda wb, n, q: wb.insert_columns(n, "B"),
    "delete_rows": lambda wb, n, q: wb.delete_rows(n, 2),
    "delete_columns": lambda wb, n, q: wb.delete_columns(n, "C"),
    "move_rows": lambda wb, n, q: wb.move_rows(n, 2, 4),
    "move_columns": lambda wb, n, q: wb.move_columns(n, "A", "C"),
    "rearrange_columns": lambda wb, n, q: wb.rearrange_columns(n, ["C", "A", "B"]),
    "merge": lambda wb, n, q: wb.merge(f"{q}!E1:F1"),
    "unmerge": lambda wb, n, q: wb.unmerge(f"{q}!E1:F1"),
    "set_comment": lambda wb, n, q: wb.set_comment(f"{q}!A1", "count again"),
    "set_comments": lambda wb, n, q: wb.set_comments({f"{q}!A1": "a", f"{q}!B1": "b"}),
    "remove_comment": lambda wb, n, q: wb.remove_comment(f"{q}!A1"),
    "set_autofilter range": lambda wb, n, q: wb.set_autofilter(f"{q}!A1:C4"),
    "set_autofilter sheet": lambda wb, n, q: wb.set_autofilter(n, "A1:C4"),
    "set_autofilter quoted sheet": lambda wb, n, q: wb.set_autofilter(f"{q}!", "A1:C4"),
    "clear_autofilter": lambda wb, n, q: wb.clear_autofilter(n),
    "hide_rows": lambda wb, n, q: wb.hide_rows(n, [2]),
    "show_rows": lambda wb, n, q: wb.show_rows(n, [2]),
    "hide_columns": lambda wb, n, q: wb.hide_columns(n, ["B"]),
    "show_columns": lambda wb, n, q: wb.show_columns(n, ["B"]),
    "write_table sheet": lambda wb, n, q: wb.write_table(n, "H1", ["a", "b"], [[1, 2]]),
    "write_table address": lambda wb, n, q: wb.write_table(None, f"{q}!H1", ["a", "b"],
                                                           [[1, 2]]),
    "write_table as table": lambda wb, n, q: wb.write_table(n, "H1", ["a", "b"], [[1, 2]],
                                                            as_table="Parts"),
    "append_rows sheet": lambda wb, n, q: wb.append_rows(n, [["new", 4, "w"]]),
    "append_rows region": lambda wb, n, q: wb.append_rows(f"{q}!A1:C4", [["new", 4, "w"]]),
    "clear_sheet": lambda wb, n, q: wb.clear_sheet(n, force=True),
}

#: What a call undoing another needs done first (on the twin as well, so both start alike).
SETUP = {
    "unprotect_sheet": lambda wb, n, q: wb.protect_sheet(n),
    "unmerge": lambda wb, n, q: wb.merge(f"{q}!E1:F1"),
    "remove_comment": lambda wb, n, q: wb.set_comment(f"{q}!A1", "a"),
    "clear_autofilter": lambda wb, n, q: wb.set_autofilter(n, "A1:C4"),
    "show_rows": lambda wb, n, q: wb.hide_rows(n, [2]),
    "show_columns": lambda wb, n, q: wb.hide_columns(n, ["B"]),
}


@pytest.mark.parametrize("name", NAMES)
@pytest.mark.parametrize("call", list(EDITS))
def test_an_edit_acts_on_the_sheet_named_with_its_spaces(name, call):
    wb = Workbook.open(_book(name))
    if call in SETUP:
        for sheet in (TWIN, name):
            SETUP[call](wb, sheet, quote_sheet(sheet))
    before = _parts(wb)
    EDITS[call](wb, name, quote_sheet(name))
    after = _parts(wb)
    assert after[_part(wb, TWIN)] == before[_part(wb, TWIN)], "the twin was changed"
    assert after[_part(wb, name)] != before[_part(wb, name)], "the sheet named is unchanged"
    Workbook.open(wb.to_bytes())          # and the result opens


@pytest.mark.parametrize("name", NAMES)
def test_sheet_level_calls_take_the_exact_name(name):
    wb = Workbook.open(_book(name))
    wb.set_sheet_state(name, "hidden")
    assert [(s.name, s.visible) for s in wb.sheets][:2] == [(TWIN, True), (name, False)]
    wb.set_sheet_state(quote_sheet(name), "visible")
    assert wb.sheet_referents(name) == []
    wb.rename_sheet(name, "Spare")
    assert [s.name for s in wb.sheets] == [TWIN, "Spare", "Other"]
    wb.rename_sheet("Spare", name)        # a name with outer spaces is a name Excel takes
    assert [s.name for s in wb.sheets] == [TWIN, name, "Other"]
    wb.delete_sheet(quote_sheet(name) + "!", force=True)
    assert [s.name for s in wb.sheets] == [TWIN, "Other"]


@pytest.mark.parametrize("name", NAMES)
def test_reads_take_the_exact_name(name):
    wb = Workbook.open(_book(name))
    q = quote_sheet(name)
    assert wb.sheet(name).name == name and wb.sheet(q).name == name
    assert wb.sheet(TWIN).name == TWIN
    assert wb.sheet_of(f"{q}!A1").name == name
    assert [row[0].v for row in wb.read(f"{q}!A2:A3")] == ["here-1", "here-2"]
    assert wb.cell(f"{q}!A2").v == "here-1"
    assert [v.v for v in wb.cells(name) if v.column == 1] == ["SKU", "here-1", "here-2",
                                                              "here-3"]
    assert wb.used_range(name).ref == "A1:C4"
    assert wb.merges(name) == [] and wb.comments(name) == [] and wb.hyperlinks(name) == []
    assert wb.validations(name) == []
    assert [r.sheet for r in wb.find_regions(name)] == [name]
    assert wb.detect_region(f"{q}!A1").sheet == name


@pytest.mark.parametrize("name", NAMES)
def test_the_name_with_its_spaces_dropped_finds_the_one_sheet_it_can_mean(name):
    """No twin: ``Stock`` reaches ``Stock `` (rule 1), in edits and reads alike."""
    book = WorkbookBuilder()
    book.sheet(name).rows("A1", [["SKU"], ["here-1"]])
    book.sheet("Other").rows("A1", [["n"]])
    wb = Workbook.open(book.to_bytes())
    assert wb.sheet("Stock").name == name
    wb.freeze_panes("Stock", rows=1)
    wb.set_cells({"Stock!B1": 3})
    assert wb.cell(f"{quote_sheet(name)}!B1").v == 3
    wb.rename_sheet("Stock", "Spare")
    assert [s.name for s in wb.sheets] == ["Spare", "Other"]


def test_a_name_two_sheets_differ_from_only_in_spaces_is_refused_naming_both():
    book = WorkbookBuilder()
    for name in ("Stock ", " Stock"):
        book.sheet(name).rows("A1", [["SKU"]])
    wb = Workbook.open(book.to_bytes())
    with pytest.raises(NameNotFound) as missing:
        wb.sheet("Stock")
    assert set(missing.value.valid_options) == {"'Stock '", "' Stock'"}
    with pytest.raises(EditError) as refused:
        wb.rename_sheet("Stock", "Spare")
    assert "'Stock '" in str(refused.value) and "' Stock'" in str(refused.value)
    assert [s.name for s in wb.sheets] == ["Stock ", " Stock"]
