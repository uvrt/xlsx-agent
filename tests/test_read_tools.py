"""R7: ``xlsx_read`` and ``xlsx_profile`` through a real ooxml-edit Toolbox, and what the
library gives them (comments, hyperlinks, array blocks).

Bounded on the largest fixture: every page within ``max_result_chars``, every cut reported,
and the pages together give every cell once.  Cell content -- values, formulas, comments,
links -- is in results only, never in a summary, warning, error or log record.

Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
import logging
from datetime import date

import pytest

from build import Formula, WorkbookBuilder
from ooxml_edit.sheets import Area, Cell
from ooxml_edit.tools import Limits, Toolbox
from test_catalogue import MARKER, ROWS, WIDE, shop, warehouse

from xlsx_agent import Workbook
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.read import ALLOWED_SHOWN, PROFILE_TEXT, READ_ROWS, TEXT_SHOWN

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
LONG = 32767


def annotated() -> WorkbookBuilder:
    """A table with comments, hyperlinks (one cell, a range, a place in the workbook), a
    shared formula group, an array formula's block, a merge, and text at the cell limit;
    the marker in cells, comments, links and formulas."""
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["SKU", "Name", "Qty", "Double", "Spill"],
                      ["SKU-1", f"Bolt {MARKER}", 3],
                      ["SKU-2", "Nut", 4],
                      ["SKU-3", "Washer", 5]])
    sheet.cell("D2", Formula(f'C2*2+LEN("{MARKER}")', cached=15, kind="shared", ref="D2:D4", si=0))
    sheet.cell("D3", Formula("", cached=17, kind="shared", si=0))
    sheet.cell("D4", Formula("", cached=19, kind="shared", si=0))
    sheet.cell("E2", Formula("C2:C4*10", cached=30, kind="array", ref="E2:E4"))
    sheet.cell("E3", 40)
    sheet.cell("E4", 50)
    sheet.table("Parts", "A1:E4", ["SKU", "Name", "Qty", "Double", "Spill"], write_header=False)
    sheet.comment("B2", f"Check {MARKER}", author="Reviewer")
    sheet.comment("C3", "Counted twice")
    sheet.hyperlink("A2", url=f"https://example.com/{MARKER}")
    sheet.hyperlink("A3:A4", location="Notes!A1")
    sheet.merge("G1:H1")
    sheet.cell("G1", "Merged title")
    notes = book.sheet("Notes")
    notes.cell("A1", "x" * LONG)
    notes.cell("B1", "short")
    return book


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


@pytest.fixture(scope="module")
def large() -> bytes:
    """The largest fixture, built once."""
    return warehouse().to_bytes()


def _session(toolbox, book: "WorkbookBuilder | bytes", name: str = "book.xlsx"):
    session = toolbox.session(clock=CLOCK)
    data = book if isinstance(book, bytes) else book.to_bytes()
    blob = session.add_blob(data, name=name)
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def _read(toolbox, session, **arguments):
    return toolbox.dispatch(session, "xlsx_read", {"doc": "d1", **arguments})


def _cells(result) -> list[dict]:
    return [cell for row in result.data["rows"] for cell in row]


# -- the library -------------------------------------------------------------------------------


def test_comments_hyperlinks_and_array_blocks_are_read():
    wb = Workbook.open(annotated().to_bytes())
    comments = wb.comments("Stock")
    assert [(c.cell, c.author, c.text) for c in comments] == [
        (Cell(2, 2), "Reviewer", f"Check {MARKER}"), (Cell(3, 3), "Reviewer", "Counted twice")]
    assert MARKER not in repr(comments)
    links = wb.hyperlinks("Stock")
    assert [(l.area.ref, l.url, l.location) for l in links] == [
        ("A2", f"https://example.com/{MARKER}", None), ("A3:A4", None, "Notes!A1")]
    assert MARKER not in repr(links)
    assert [a.ref for a in wb.arrays("Stock")] == ["E2:E4"]
    assert wb.comments("Notes") == [] and wb.hyperlinks("Notes") == [] and wb.arrays("Notes") == []


# -- xlsx_read ---------------------------------------------------------------------------------


def test_a_table_reads_typed_with_addresses_and_its_columns(box):
    session = _session(box, shop())
    result = _read(box, session, target="t:Products")
    assert result.ok and result.total == 4 and result.next_cursor is None
    data = result.data
    assert (data["target"], data["sheet"], data["range"]) == \
        ("t:Products", {"name": "Stock", "id": 1}, "Stock!A2:D5")
    assert data["columns"] == ["SKU", "Category", "Price", "Added"]
    assert data["rows"][0] == [{"a": "Stock!A2", "v": "SKU-0001", "t": "text"},
                               {"a": "Stock!B2", "v": "Hardware", "t": "text"},
                               {"a": "Stock!C2", "v": 12.5, "t": "number"},
                               {"a": "Stock!D2", "v": "2026-01-05", "t": "date",
                                "serial": 46027.0}]
    assert data["rows"][3][2] == {"a": "Stock!C5", "v": None, "t": "blank", "formula": True,
                                  "cache": "absent", "issue": "no_cached_value"}
    assert "cut" not in data and result.summary == "t:Products: rows 1-4 of 4"


@pytest.mark.parametrize("target,cells,columns", [
    ("t:Products[Price]", ["Stock!C2", "Stock!C3", "Stock!C4", "Stock!C5"], ["Price"]),
    ("t:Products/#2", ["Stock!A3", "Stock!B3", "Stock!C3", "Stock!D3"],
     ["SKU", "Category", "Price", "Added"]),
    ("t:Products/SKU-0003/Price", ["Stock!C4"], ["Price"]),
    ("Stock!B2", ["Stock!B2"], None),
    ("#1!A1:B2", ["#1!A1", "#1!B1", "#1!A2", "#1!B2"], None),
], ids=["column", "row", "value", "cell", "range-by-id"])
def test_every_address_form_reads(box, target, cells, columns):
    session = _session(box, shop())
    result = _read(box, session, target=target)
    assert result.ok, result.error
    assert [c["a"] for c in _cells(result)] == [c.replace("#1!", "Stock!") for c in cells]
    assert result.data.get("columns") == columns


def test_a_range_is_clipped_to_the_used_range_and_says_so(box):
    session = _session(box, shop())
    result = _read(box, session, target="Stock!A1:Z99")
    assert result.data["range"] == "Stock!A1:D5" and result.data["requested"] == "Stock!A1:Z99"
    assert result.total == 5
    whole = _read(box, session, target="Stock!C:C")
    assert whole.data["range"] == "Stock!C1:C5" and whole.data["requested"] == "Stock!C:C"
    sheet = _read(box, session, target="#1!")                     # a whole sheet (R9)
    assert sheet.data["range"] == "Stock!A1:D5" and sheet.data["requested"] == "#1!"
    assert sheet.total == 5
    empty = _read(box, session, target="Stock!M1:N4")
    assert empty.ok and empty.data["rows"] == [] and empty.total == 0
    assert empty.data["range"] is None and empty.summary == "Stock!M1:N4: no cells to read"
    lone = _read(box, session, target="Stock!M4")
    assert _cells(lone) == [{"a": "Stock!M4", "v": None, "t": "blank"}]


def test_a_row_by_key_takes_the_key_column(box):
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["Ref", "Code", "Qty"], ["A-1", "X1", 1], ["A-2", "X2", 2]])
    sheet.table("Lines", "A1:C3", ["Ref", "Code", "Qty"], write_header=False)
    session = _session(box, book)
    refused = _read(box, session, target="t:Lines/A-2")
    assert refused.error.code == "refused" and refused.error.valid_options == ["Ref"]  # holds it
    found = _read(box, session, target="t:Lines/A-2/Qty", key="Ref")
    assert _cells(found) == [{"a": "Stock!C3", "v": 2, "t": "number"}]
    assert found.data["target"] == "t:Lines/A-2/Qty" and found.warnings == []
    stray = _read(box, session, target="Stock!A1:C3", key="Ref")
    assert stray.error.code == "invalid_arguments" and stray.error.field == "key"
    sources = _read(box, session, target="t:Lines/A-2/Qty", sources=True)
    assert (sources.error.code, sources.error.field) == ("invalid_arguments", "sources")
    assert sources.error.message == (                 # where it works, what to do instead
        "sources works on a dataset (d:) only, where each value lists the cells it came "
        "from; a t: or cell read gives each cell its own address (a), its source: read it "
        "without sources")



# -- a row by key, its key column and its columns (L8) -----------------------------------------


def kit() -> WorkbookBuilder:
    """Two tables: Kit, whose one key column is Part (Label repeats; Cost is no key), one
    part stored with other case and a trailing space; Pairs, where Code and Alt can both be
    the key."""
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["Part", "Label", "Qty", "Cost"], ["P-10", "Gear", 4, 2.5],
                      ["P-11", "Gear", 4, 1.25], ["p-12 ", "Pin", 9, 0.5]])
    sheet.table("Kit", "A1:D4", ["Part", "Label", "Qty", "Cost"], write_header=False)
    other = book.sheet("Two")
    other.rows("A1", [["Code", "Alt", "Note"], ["C-1", "Z-9", "x"], ["C-2", "Z-8", "x"]])
    other.table("Pairs", "A1:C3", ["Code", "Alt", "Note"], write_header=False)
    return book


def _values(result) -> list[list]:
    return [[cell["v"] for cell in row] for row in result.data["rows"]]


def test_columns_pick_a_t_targets_columns_in_their_order(box):
    session = _session(box, kit())
    row = _read(box, session, target="t:Kit/P-11", columns=["Cost", "Qty"])
    assert row.ok and row.warnings == [] and row.data["target"] == "t:Kit/P-11"
    assert _cells(row) == [{"a": "Stock!D3", "v": 1.25, "t": "number"},
                           {"a": "Stock!C3", "v": 4, "t": "number"}]
    assert row.data["columns"] == ["Cost", "Qty"]
    every = _read(box, session, target="t:Kit", columns=["qty", "Part"])   # case aside
    assert _values(every) == [[4, "P-10"], [4, "P-11"], [9, "p-12 "]]
    assert every.data["columns"] == ["Qty", "Part"]
    both = _read(box, session, target="t:Kit/{P-10,P-11}/Cost", columns=["Part"])
    assert _values(both) == [[2.5, "P-10"], [1.25, "P-11"]]         # the address's first
    page = _read(box, session, target="t:Kit", columns=["Qty"], limit=2)
    rest = _read(box, session, target="t:Kit", columns=["Qty"], cursor=page.next_cursor)
    assert _values(page) + _values(rest) == [[4], [4], [9]]
    other = _read(box, session, target="t:Kit", columns=["Part"], cursor=page.next_cursor)
    assert other.error.code == "invalid_arguments" and other.error.field == "cursor"
    unknown = _read(box, session, target="t:Kit", columns=["Price"])
    assert unknown.error.code == "not_found" and "Qty" in unknown.error.message
    cells = _read(box, session, target="Stock!A1:B2", columns=["Part"])
    assert (cells.error.code, cells.error.field) == ("invalid_arguments", "columns")
    assert "t: target" in cells.error.message


def test_key_read_as_the_column_to_read_reads_it(box):
    session = _session(box, kit())
    row = _read(box, session, target="t:Kit/P-11", key="Cost")      # the misread: Cost
    assert row.ok and _values(row) == [["P-11", "Gear", 4, 1.25]]   # the whole row has it
    assert row.warnings == ["key read as a column to read: 'Cost' does not hold target's "
                            "key, 'Part' does; key is the column holding it, columns picks "
                            "columns"]
    value = _read(box, session, target="t:Kit/P-11/Qty", key="Cost")
    assert _values(value) == [[4, 1.25]] and value.data["columns"] == ["Qty", "Cost"]
    table = _read(box, session, target="t:Kit", key="Cost")         # no row by key at all
    assert table.ok and table.total == 3 and "key read as a column to read" in table.warnings[0]
    column = _read(box, session, target="t:Kit[Qty]", key="Cost")
    assert _values(column) == [[4, 2.5], [4, 1.25], [9, 0.5]]


def test_key_read_as_the_rows_key_reads_the_row(box):
    session = _session(box, kit())
    row = _read(box, session, target="t:Kit", key="P-10")
    assert row.ok and _values(row) == [["P-10", "Gear", 4, 2.5]]
    assert row.warnings == ["key read as the key of the row to read: key is the column "
                            "holding it, and the row goes in target (t:Kit/<key>)"]
    value = _read(box, session, target="t:Kit[Cost]", key="P-10")
    assert _values(value) == [[2.5]]
    nowhere = _read(box, session, target="t:Kit", key="Q-1")
    assert (nowhere.error.code, nowhere.error.field) == ("invalid_arguments", "key")
    assert nowhere.error.valid_options == ["Part"] and "Q-1" not in nowhere.error.message
    assert nowhere.error.message == (
        "key is the column of t:Kit holding the key in target, one of valid_options; the row "
        "goes in target (t:Kit/<key>), the columns to read in columns")


def test_key_column_and_row_are_read_as_key_and_target(box):
    session = _session(box, kit())
    old = _read(box, session, target="t:Kit/P-10/Qty", key_column="Part")
    assert old.ok and _values(old) == [[4]] and old.warnings == ["key_column read as key"]
    row = box.dispatch(session, "xlsx_read", {"doc": "d1", "row": "t:Kit/P-10"})
    assert row.ok and _values(row) == [["P-10", "Gear", 4, 2.5]]
    assert row.warnings == ["row read as target"]
    twice = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Kit",
                                                "row": "t:Kit/P-10"})
    assert twice.error.code == "invalid_arguments" and twice.error.valid_options == ["target"]


def test_a_key_not_found_says_where_it_is_or_which_rows_are_like_it(box):
    session = _session(box, kit())
    derived = _read(box, session, target="t:Kit/P-10")           # Part is the one key column
    assert derived.ok and derived.warnings == []
    near = _read(box, session, target="t:Kit/P-12")
    assert near.error.code == "not_found" and near.error.valid_options == ["t:Kit/#3"]
    assert near.error.message == ("no row of t:Kit has that key in column 'Part'; 1 row has "
                                  "a key like it (case, spaces or a character apart): "
                                  "valid_options, by position")
    several = _read(box, session, target="t:Kit/P-1")
    assert several.error.valid_options == ["t:Kit/#1", "t:Kit/#2"]
    none = _read(box, session, target="t:Kit/W-77")
    assert none.error.code == "not_found" and none.error.valid_options == []
    choose = _read(box, session, target="t:Pairs/Z-9")             # Code or Alt: refused
    assert (choose.error.code, choose.error.field) == ("refused", "key")
    assert choose.error.valid_options == ["Alt"]                  # the one holding it
    assert choose.error.message == ("t:Pairs has no one key column (2 can be); 'Alt' holds "
                                    "target's key: give key")
    other = _read(box, session, target="t:Pairs/C-9")             # in neither: both offered
    assert other.error.code == "refused" and other.error.valid_options == ["Code", "Alt"]
    found = _read(box, session, target="t:Pairs/Z-9", key="Alt")
    assert found.ok and _values(found) == [["C-1", "Z-9", "x"]]
    misread = _read(box, session, target="t:Pairs/Z-9", key="Note")  # Note holds no key
    assert misread.ok and _values(misread) == [["C-1", "Z-9", "x"]]
    wrong = _read(box, session, target="t:Pairs/Z-9", key="Code")  # a key column, not this one
    assert wrong.ok and "'Alt' does" in wrong.warnings[0]


def test_formats_formulas_comments_links_and_merges_on_request(box):
    session = _session(box, annotated())
    plain = _read(box, session, target="Stock!A1:H4")
    assert MARKER in json.dumps(plain.data)            # the cell's own text is content
    assert not any(k in c for c in _cells(plain) for k in ("f", "comment", "link", "format"))
    assert "merges" not in plain.data
    assert plain.data["arrays"] == ["Stock!E2:E4"]      # always: a member says nothing of it
    full = _read(box, session, target="Stock!A1:H4", formats=True, formulas=True,
                 comments=True, hyperlinks=True, merges=True)
    cells = {c["a"]: c for c in _cells(full)}
    assert cells["Stock!D2"]["f"] == f'C2*2+LEN("{MARKER}")'
    assert cells["Stock!D3"]["shared_from"] == "Stock!D2" and "f" not in cells["Stock!D3"]
    assert cells["Stock!E2"]["array"] == "E2:E4" and cells["Stock!E2"]["f"] == "C2:C4*10"
    assert cells["Stock!B2"]["comment"] == {"author": "Reviewer", "text": f"Check {MARKER}"}
    assert cells["Stock!C3"]["comment"]["text"] == "Counted twice"
    assert cells["Stock!A2"]["link"] == {"url": f"https://example.com/{MARKER}"}
    assert cells["Stock!A3"]["link"] == cells["Stock!A4"]["link"] == {"location": "Notes!A1"}
    assert cells["Stock!C2"]["format"] == "General"
    assert full.data["merges"] == ["Stock!G1:H1"]
    lower = _read(box, session, target="Stock!A3:H4", merges=True)
    assert "merges" not in lower.data and lower.data["arrays"] == ["Stock!E2:E4"]


def test_long_text_is_cut_with_its_length_and_reported(box):
    session = _session(box, annotated())
    row = _read(box, session, target="Notes!A1:B1")
    first = row.data["rows"][0][0]
    assert len(first["v"]) == TEXT_SHOWN and first["length"] == LONG and first["at_limit"]
    assert row.data["cut"] == {"text": {"count": 1, "first": ["Notes!A1"]}}
    one = _read(box, session, target="Notes!A1")       # one cell: as much as the page holds
    cell = one.data["rows"][0][0]
    assert TEXT_SHOWN < len(cell["v"]) < LONG and cell["length"] == LONG
    assert len(one.to_text()) <= box.limits.max_result_chars
    small = _read(box, session, target="Notes!A1", max_chars=1000)
    assert len(small.to_text()) <= 1000 and small.data["cut"]["text"]["count"] == 1


def test_a_page_is_limited_and_the_cursor_gives_the_next(box):
    session = _session(box, shop())
    first = _read(box, session, target="t:Products", limit=3)
    assert len(first.data["rows"]) == 3 and first.data["cut"] == {"rows": 1}
    assert first.summary == "t:Products: rows 1-3 of 4; more with next_cursor"
    second = _read(box, session, target="t:Products", limit=3, cursor=first.next_cursor)
    assert [r[0]["a"] for r in second.data["rows"]] == ["Stock!A5"]
    assert second.next_cursor is None and "cut" not in second.data
    other = _read(box, session, target="t:Products[Price]", cursor=first.next_cursor)
    assert other.error.code == "invalid_arguments" and other.error.field == "cursor"
    for bad in ("c3", "9.0." + first.next_cursor.split(".")[2], "x.y.z"):
        refused = _read(box, session, target="t:Products", cursor=bad)
        assert refused.error.code == "invalid_arguments", bad
    too_many = _read(box, session, target="t:Products", limit=READ_ROWS + 1)
    assert too_many.error.code == "invalid_arguments"


def _walk(toolbox, session, **arguments) -> tuple[list[str], list]:
    """Every page of a read: the addresses given, in order, and the results."""
    addresses, results, cursor = [], [], None
    while True:
        result = _read(toolbox, session, **arguments, **({"cursor": cursor} if cursor else {}))
        assert result.ok, result.error
        results.append(result)
        addresses += [c["a"] for c in _cells(result)]
        cursor = result.next_cursor
        if cursor is None:
            return addresses, results


def _expected(top: int, rows: int, width: int, sheet: str = "Stock") -> list[str]:
    letters = [chr(65 + k) for k in range(width)]
    return [f"{sheet}!{letters[c]}{r}" for r in range(top, top + rows) for c in range(width)]


@pytest.mark.parametrize("limits", [Limits(), Limits(max_result_chars=6000, max_list_items=20)],
                         ids=["default", "tight"])
def test_the_largest_fixture_reads_bounded_and_every_cut_is_reported(large, limits):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=limits) as toolbox:
        session = _session(toolbox, large)
        addresses, results = _walk(toolbox, session, target="t:Products", limit=READ_ROWS)
        assert addresses == _expected(2, ROWS, WIDE)
        for result in results:
            assert len(result.to_text()) <= limits.max_result_chars
            assert result.total == ROWS
            if result.next_cursor is not None:
                cut = result.data["cut"]
                assert cut["chars"] is True and cut["rows"] > 0
            else:
                assert "rows" not in result.data.get("cut", {})
        assert len(results) > 1
        # every row given whole or split, the split row's rest first on the next page
        split = [r for r in results if "cells" in r.data.get("cut", {})]
        for result in split:
            assert len(result.data["rows"][-1]) == WIDE - result.data["cut"]["cells"]


def test_a_wide_regions_names_are_left_to_describe(box):
    book = WorkbookBuilder()
    names = [f"A rather long column name {k:04d}" for k in range(300)]
    sheet = book.sheet("Wide")
    sheet.rows("A1", [names, list(range(300))])
    sheet.table("Wide", "A1:KN2", names, write_header=False)
    session = _session(box, book)
    result = _read(box, session, target="t:Wide")
    assert "columns" not in result.data and result.data["cut"]["column_names"] == 300
    assert len(result.to_text()) <= box.limits.max_result_chars
    assert [c["a"] for c in _cells(result)][:2] == ["Wide!A2", "Wide!B2"]


def test_a_long_sheet_reads_a_page_at_a_time(box):
    book = WorkbookBuilder()
    book.sheet("Log").rows("A1", [[f"E-{r}", r, r * 0.5] for r in range(1, 20001)])
    session = _session(box, book)
    first = _read(box, session, target="Log!A:C")
    assert first.data["range"] == "Log!A1:C20000" and first.total == 20000
    assert len(first.data["rows"]) == 50 and first.data["cut"] == {"rows": 19950}
    most = _read(box, session, target="Log!A:C", limit=READ_ROWS)
    assert len(most.to_text()) <= box.limits.max_result_chars
    assert most.data["cut"]["chars"] is True and most.data["cut"]["rows"] > 20000 - READ_ROWS
    later = _read(box, session, target="Log!A:C", limit=3, cursor=first.next_cursor)
    assert [r[0]["a"] for r in later.data["rows"]] == ["Log!A51", "Log!A52", "Log!A53"]


def test_a_narrow_page_splits_a_row_and_resumes_it(box, large):
    session = _session(box, large)
    addresses, results = _walk(box, session, target="Stock!A2:X31", limit=2, max_chars=2000)
    assert addresses == _expected(2, 30, WIDE)
    assert all(len(r.to_text()) <= 2000 for r in results)
    assert any("cells" in r.data["cut"] for r in results if r.next_cursor)
    assert any(r.data["rows"][0][0]["a"][len("Stock!"):][0] != "A" for r in results[1:])


def test_the_limit_stops_a_page_before_the_characters_do(box, large):
    session = _session(box, large)
    result = _read(box, session, target="t:Products[Products column 3]", limit=10)
    assert len(result.data["rows"]) == 10 and result.data["cut"] == {"rows": ROWS - 10}
    assert result.total == ROWS


@pytest.mark.parametrize("target,code", [
    ("d:by_category", "not_found"),             # no such dataset in the session (R11)
    ("$step1", "not_found"),
    ("t:Stock!A1:F20", "invalid_arguments"),
    ("t:Nope", "not_found"),
    ("Nope!A1", "not_found"),
    ("t:Products/#9", "not_found"),
    ("t:Products[Nope]", "not_found"),
], ids=["dataset", "ref", "bad-region", "table", "sheet", "position", "column"])
def test_read_errors_use_the_tool_error_mapping(box, target, code):
    session = _session(box, shop())
    result = _read(box, session, target=target)
    assert not result.ok and result.error.code == code, result.error


# -- xlsx_profile ------------------------------------------------------------------------------


def _profile(toolbox, session, **arguments):
    return toolbox.dispatch(session, "xlsx_profile", {"doc": "d1", **arguments})


def test_the_profile_is_every_columns_whole_profile(box):
    session = _session(box, shop())
    result = _profile(box, session, target="t:Products")
    wb = Workbook.open(shop().to_bytes())
    assert result.ok and result.total == 4 and result.next_cursor is None
    assert result.data == {"target": "t:Products", "range": "Stock!A1:D5", "rows": 4,
                           "columns": [s.as_dict() for s in wb.profile("t:Products")]}
    named = _profile(box, session, target="t:Products", columns=["price", "SKU"])
    assert [c["column"] for c in named.data["columns"]] == ["Price", "SKU"]
    one = _profile(box, session, target="t:Products[Added]")
    assert [c["column"] for c in one.data["columns"]] == ["Added"]
    assert one.data["columns"][0]["min"] == "2026-01-05"


@pytest.mark.parametrize("arguments,code", [
    ({"target": "t:Products/#1"}, "invalid_arguments"),
    ({"target": "Stock!A1:D5"}, "invalid_arguments"),
    ({"target": "t:Products[Price]", "columns": ["SKU"]}, "invalid_arguments"),
    ({"target": "t:Products", "columns": ["Nope"]}, "not_found"),
    ({"target": "t:Nope"}, "not_found"),
    ({"target": "t:Products", "cursor": "7.0.00000000"}, "invalid_arguments"),
], ids=["row", "range", "both", "column", "table", "cursor"])
def test_profile_errors(box, arguments, code):
    session = _session(box, shop())
    result = _profile(box, session, **arguments)
    assert not result.ok and result.error.code == code, result.error


@pytest.mark.parametrize("limits", [Limits(), Limits(max_list_items=5),
                                    Limits(max_result_chars=4000)],
                         ids=["default", "few-items", "few-chars"])
def test_the_largest_fixtures_profile_is_paged_and_bounded(large, limits):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=limits) as toolbox:
        session = _session(toolbox, large)
        names, cursor, pages = [], None, 0
        while True:
            result = _profile(toolbox, session, target="t:Products",
                              **({"cursor": cursor} if cursor else {}))
            assert result.ok, result.error
            assert len(result.to_text()) <= limits.max_result_chars
            pages += 1
            names += [c["column"] for c in result.data["columns"]]
            cursor = result.next_cursor
            if cursor is None:
                break
            assert result.data["cut"]["columns"] == WIDE - len(names)
        assert names == [f"Products column {k + 1}" for k in range(WIDE)]
        assert pages > 1 or limits == Limits()


def test_profile_text_and_allowed_lists_are_cut_and_counted(box):
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["Note", "Colour"]] + [["n" * 500, f"C{r}"] for r in range(1, 4)])
    sheet.table("Notes", "A1:B4", ["Note", "Colour"], write_header=False)
    sheet.rows("D1", [[f"Colour {k}"] for k in range(ALLOWED_SHOWN + 20)])
    sheet.validation_list("B2:B4", source=f"$D$1:$D${ALLOWED_SHOWN + 20}")
    session = _session(box, book)
    result = _profile(box, session, target="t:Notes")
    note, colour = result.data["columns"]
    assert note["values"][0]["v"] == "n" * PROFILE_TEXT + "…"
    rule = colour["validation"][0]
    assert len(rule["values"]) == ALLOWED_SHOWN and rule["more"] == 20
    assert result.data["cut"] == {"text": 1, "allowed_values": 20}


def test_describe_names_xlsx_profile_for_what_it_cut(box, large):
    session = _session(box, large)
    described = box.dispatch(session, "describe", {"doc": "d1"})
    cut = described.data["cut"]
    assert cut["columns"] > 0 and "xlsx_profile" in cut["next"]
    region = described.data["regions"][0]
    left = region["more_columns"]
    left = len(left) if isinstance(left, list) else left
    assert left > 0
    profiled = _profile(box, session, target=region["address"])
    assert profiled.total == len(region["columns"]) + left


# -- content stays out of diagnostics ----------------------------------------------------------


def test_no_cell_content_reaches_a_diagnostic(box, caplog):
    caplog.set_level(logging.DEBUG)
    session = _session(box, annotated())
    results = [
        _read(box, session, target="Stock!A1:H4", formulas=True, comments=True, hyperlinks=True),
        _read(box, session, target="t:Parts", formulas=True, limit=1),
        _profile(box, session, target="t:Parts"),
        _read(box, session, target=f"t:Parts/{MARKER}"),            # a key: never echoed
        _read(box, session, target=f"t:Parts/{MARKER}", key_column="Name"),
        _read(box, session, target=f"t:Parts/{MARKER}", key="Qty"),       # a misread key
        _read(box, session, target="t:Parts", key=MARKER),                # the key as key
        _read(box, session, target=f"t:Parts/{MARKER}X"),                 # near no key
        _read(box, session, target=f"t:Parts/Bolt {MARKER}/Qty", key_column="Name"),
    ]
    assert results[-1].ok and _cells(results[-1])[0]["a"] == "Stock!C2"
    assert results[-1].summary == "a value by key of t:Parts: rows 1-1 of 1"
    formula = f'C2*2+LEN("{MARKER}")'
    assert MARKER in json.dumps(results[0].data)
    assert formula in [c.get("f") for c in _cells(results[0])]   # asked for: a result
    assert MARKER not in json.dumps(_read(box, session, target="Stock!D2").data)  # else not
    diagnostics = [r.getMessage() for r in caplog.records]
    diagnostics += [w for r in results for w in r.warnings] + [r.summary for r in results]
    diagnostics += [json.dumps(r.error.to_json()) for r in results if not r.ok]
    diagnostics += [json.dumps(record.__dict__, default=str) for record in session.log]
    assert len([r for r in results if not r.ok]) == 5
    assert diagnostics and not [d for d in diagnostics if MARKER in d]
