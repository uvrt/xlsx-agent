"""R9: ``find_text`` for a workbook -- the library's :meth:`Workbook.find_text` and the shared
tool's xlsx handler through a real ooxml-edit Toolbox.

Matches by address and kind (cells, table and region names, headers, comments, defined
names), in reading order, hidden ones marked; ``range`` as a sheet, range or table; numbers,
dates and formula text not searched; paged and bounded on the largest fixture; and the text
searched for or found never in a summary, warning, error or log record.

Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
import logging
import re
from datetime import date

import pytest

from build import Error, Formula, Rich, Run, WorkbookBuilder, zip_parts
from ooxml_edit.tools import Limits, Result, Toolbox, shared
from test_catalogue import MARKER, warehouse

from xlsx_agent import AddressError, Found, NameNotFound, Workbook
from xlsx_agent.find import MATCH_SHOWN
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def ledger() -> WorkbookBuilder:
    """Something of every kind, "Bolt" in each: a table (name, headers, cells, a totals
    label), a comment, a workbook-wide and a sheet-local name, a hidden name, a name that is
    a formula; a hidden row, a hidden column, a hidden sheet, a very hidden one; numbers,
    dates, booleans, errors and formulas that do not count.  The marker is in a cell, a
    comment and a formula."""
    book = WorkbookBuilder()
    stock = book.sheet("Stock")
    stock.rows("A1", [["SKU", "Bolt size", "Qty"],
                      ["B-1", "Bolt M4", 3],
                      ["B-2", Rich([Run("Nut "), Run("bolt", bold=True)]), 4],
                      ["B-3", f"Bolt {MARKER}", 5]])
    stock.table("Bolts", "A1:C4", ["SKU", "Bolt size", "Qty"])
    stock.comment("B3", f"Bolt checked {MARKER}")
    stock.cell("E1", Formula(f'"Bolt"&"{MARKER}"', cached=f"Bolt{MARKER}"))   # cached text
    stock.cell("E2", Formula('LEN("Bolt")', cached=4))                         # formula text
    stock.cell("E3", 2026)
    stock.cell("E4", date(2026, 1, 5))
    stock.cell("E5", True)
    stock.cell("E6", Error("#N/A"))
    stock.cell("F2", "Bolt in a hidden column")
    stock.cell("A9", "Bolt in a hidden row")
    stock.hide_column("F")
    stock.hide_row(9)
    stock.comment("F9", "Bolt note")                                # hidden row and column
    notes = book.sheet("Q3 sales", state="hidden")
    notes.cell("B2", "Spare bolt, Bolt pack")
    deep = book.sheet("Vault", state="veryHidden")
    deep.cell("A1", "Bolt")
    book.define_name("BoltList", "Stock!$B$2:$B$4")
    book.define_name("Bolt_local", "'Q3 sales'!$B$2", sheet="Q3 sales")
    book.define_name("BoltHidden", "Stock!$A$9", hidden=True)
    book.define_name("BoltRate", "0.25")
    return book


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


@pytest.fixture(scope="module")
def large() -> bytes:
    return warehouse().to_bytes()


def _session(toolbox, book: "WorkbookBuilder | bytes"):
    session = toolbox.session(clock=CLOCK)
    data = book if isinstance(book, bytes) else book.to_bytes()
    blob = session.add_blob(data, name="book.xlsx")
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def _find(toolbox, session, **arguments):
    return toolbox.dispatch(session, "find_text", {"doc": "d1", **arguments})


def _seen(found) -> list[tuple]:
    return [(f.kind, f.address) for f in found]


# -- the library -------------------------------------------------------------------------------


EVERY_BOLT = [
    ("table", "t:Bolts"),                   # row 1: at A1 the table, then its header, E1
    ("header", "t:Bolts[Bolt size]"),
    ("cell", "Stock!E1"),
    ("name", "t:BoltList"),                 # row 2: the name at B2, before the cell
    ("cell", "Stock!B2"),
    ("cell", "Stock!F2"),
    ("comment", "Stock!B3"),                # B3's own text is "Nut bolt": case counts
    ("cell", "Stock!B4"),
    ("name", "t:BoltHidden"),
    ("cell", "Stock!A9"),
    ("comment", "Stock!F9"),
    ("name", "t:'Q3 sales'!Bolt_local"),
    ("cell", "'Q3 sales'!B2"),
    ("cell", "Vault!A1"),
    ("name", "t:BoltRate"),
]


def test_every_kind_is_found_by_address_in_reading_order():
    wb = Workbook.open(ledger().to_bytes())
    found = list(wb.find_text("Bolt"))
    assert _seen(found) == EVERY_BOLT
    header = found[1]
    assert (header.cell, header.text, header.start, header.end) == ("Stock!B1", "Bolt size", 0, 4)
    assert found[0].range == "Stock!A1:C4" and found[3].range == "Stock!B2:B4"
    assert found[-1].range is None                       # a constant: listed, not a range
    # the table's header cells are its headers, not cells as well
    assert ("cell", "Stock!B1") not in _seen(found)


def test_numbers_dates_booleans_errors_and_formula_text_are_not_searched():
    wb = Workbook.open(ledger().to_bytes())
    assert list(wb.find_text("2026")) == []               # a number, a date (2026-01-05)
    assert list(wb.find_text("TRUE")) == [] and list(wb.find_text("#N/A")) == []
    assert list(wb.find_text("LEN")) == []                # formula text
    # a formula's cached text is the cell's text
    assert ("cell", "Stock!E1") in _seen(wb.find_text(f"Bolt{MARKER}"))
    assert _seen(wb.find_text('"Bolt"')) == []


def test_hidden_sheets_rows_columns_and_names_are_found_and_marked():
    wb = Workbook.open(ledger().to_bytes())
    hidden = {(f.kind, f.address) for f in wb.find_text("Bolt") if f.hidden}
    assert hidden == {("cell", "Stock!F2"), ("name", "t:BoltHidden"), ("cell", "Stock!A9"),
                      ("comment", "Stock!F9"), ("name", "t:'Q3 sales'!Bolt_local"),
                      ("cell", "'Q3 sales'!B2"), ("cell", "Vault!A1")}


def test_each_occurrence_is_a_match_case_counts_and_rich_text_is_flat():
    wb = Workbook.open(ledger().to_bytes())
    pack = [f for f in wb.find_text(re.compile("(?i)bolt")) if f.address == "'Q3 sales'!B2"]
    assert [(f.start, f.end) for f in pack] == [(6, 10), (12, 16)]
    assert _seen(wb.find_text("bolt")) == [("cell", "Stock!B3"), ("cell", "'Q3 sales'!B2")]
    assert [f.context() for f in wb.find_text("Nut bolt")] == ["Nut bolt"]
    assert list(wb.find_text(re.compile("x*"))) == []    # empty matches are none


def test_context_is_short_around_the_match():
    book = WorkbookBuilder()
    book.sheet("Notes").cell("A1", "a" * 100 + "Bolt" + "b" * 100)
    book.sheets[0].cell("A2", "c" * 1000)
    wb = Workbook.open(book.to_bytes())
    near = next(wb.find_text("Bolt"))
    assert near.context() == "…" + "a" * 40 + "Bolt" + "b" * 40 + "…"
    long = next(wb.find_text(re.compile("c+")))
    assert long.context() == "c" * MATCH_SHOWN + "…" and long.end - long.start == 1000
    assert "aaaa" not in repr(near) and "Bolt" not in repr(near)   # text kept out of repr


@pytest.mark.parametrize("within,expected", [
    ("Stock!", [e for e in EVERY_BOLT[:11]]),
    ("#1!", [e for e in EVERY_BOLT[:11]]),
    ("'Q3 sales'!", EVERY_BOLT[11:13]),
    ("#3!", EVERY_BOLT[13:14]),
    ("Stock!B2:B3", [("table", "t:Bolts"), ("name", "t:BoltList"), ("cell", "Stock!B2"),
                     ("comment", "Stock!B3")]),
    ("Stock!A:A", [("table", "t:Bolts"), ("name", "t:BoltHidden"), ("cell", "Stock!A9")]),
    ("Stock!8:9", [("name", "t:BoltHidden"), ("cell", "Stock!A9"), ("comment", "Stock!F9")]),
    ("t:Bolts", [("table", "t:Bolts"), ("header", "t:Bolts[Bolt size]"), ("name", "t:BoltList"),
                 ("cell", "Stock!B2"), ("comment", "Stock!B3"), ("cell", "Stock!B4")]),
    ("t:Bolts[Bolt size]", [("table", "t:Bolts"), ("name", "t:BoltList"), ("cell", "Stock!B2"),
                            ("comment", "Stock!B3"), ("cell", "Stock!B4")]),
    ("t:Bolts/#3", [("table", "t:Bolts"), ("name", "t:BoltList"), ("cell", "Stock!B4")]),
    ("t:BoltList", [("table", "t:Bolts"), ("name", "t:BoltList"), ("cell", "Stock!B2"),
                    ("comment", "Stock!B3"), ("cell", "Stock!B4")]),
    ("Stock!Z1:Z9", []),
])
def test_within_a_sheet_range_or_table(within, expected):
    wb = Workbook.open(ledger().to_bytes())
    assert _seen(wb.find_text("Bolt", within=within)) == expected


def test_within_refuses_what_is_no_sheet_range_or_table():
    wb = Workbook.open(ledger().to_bytes())
    with pytest.raises(AddressError, match="sheet, range or table"):
        list(wb.find_text("Bolt", within="d:by_category"))
    with pytest.raises(NameNotFound):
        list(wb.find_text("Bolt", within="Nowhere!"))
    with pytest.raises(TypeError):
        list(wb.find_text(""))


def test_an_out_of_order_sheet_gives_its_cells_then_the_rest():
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.cell("A1", "Bolt one")
    sheet.comment("A1", "Bolt note")
    sheet.cell("A3", "Bolt three")
    parts = book.parts()
    xml = parts["xl/worksheets/sheet1.xml"].decode()
    rows = re.findall(r"<row [^>]*>.*?</row>", xml)
    parts["xl/worksheets/sheet1.xml"] = xml.replace("".join(rows), "".join(reversed(rows))).encode()
    wb = Workbook.open(zip_parts(parts))
    assert _seen(wb.find_text("Bolt")) == [("cell", "Stock!A3"), ("cell", "Stock!A1"),
                                           ("comment", "Stock!A1")]


# -- the tool ----------------------------------------------------------------------------------


def test_find_text_through_a_toolbox_gives_matches_by_address_and_kind(box):
    session = _session(box, ledger())
    result = _find(box, session, text="Bolt")
    assert result.ok, result.error
    data = result.data
    assert [(m["kind"], m["address"]) for m in data["matches"]] == EVERY_BOLT
    assert data["kinds"] == {"table": 1, "name": 4, "header": 1, "cell": 7, "comment": 2}
    assert result.total == 15 and result.next_cursor is None
    assert data["sheets"] == [{"name": "Stock", "id": 1}, {"name": "Q3 sales", "id": 2},
                              {"name": "Vault", "id": 3}]
    first = data["matches"][:2]
    assert first[0] == {"address": "t:Bolts", "kind": "table", "range": "Stock!A1:C4",
                        "context": "Bolts"}
    assert first[1] == {"address": "t:Bolts[Bolt size]", "kind": "header", "cell": "Stock!B1",
                        "context": "Bolt size"}
    assert data["matches"][-1] == {"address": "t:BoltRate", "kind": "name", "context": "BoltRate"}
    assert {"address": "Vault!A1", "kind": "cell", "context": "Bolt", "hidden": True} \
        in data["matches"]
    assert result.summary == "15 match(es); 1-15 given"


def test_regex_and_range_through_a_toolbox(box):
    session = _session(box, ledger())
    result = _find(box, session, text=r"B-\d", regex=True, range="t:Bolts")
    assert [m["address"] for m in result.data["matches"]] == ["Stock!A2", "Stock!A3", "Stock!A4"]
    assert result.data["range"] == "t:Bolts"
    none = _find(box, session, text="Bolt", range="Stock!Z1:Z9")
    assert none.ok and none.total == 0 and none.data["matches"] == [] \
        and none.summary == "0 matches"
    assert _find(box, session, text="B-\\d", range="Stock!").total == 0   # not a regex


@pytest.mark.parametrize("arguments,code,field", [
    ({"text": "Bolt", "slides": [1]}, "invalid_arguments", "slides"),
    ({"text": "Bolt", "stories": "all"}, "invalid_arguments", "stories"),
    ({"text": "(?P<Bolt", "regex": True}, "invalid_arguments", "text"),
    ({"text": ""}, "invalid_arguments", "text"),
    ({"text": "Bolt", "range": "Nowhere!"}, "not_found", None),
    ({"text": "Bolt", "range": "t:Nope"}, "not_found", None),
    ({"text": "Bolt", "range": "d:by_category"}, "invalid_arguments", None),
    ({"text": "Bolt", "range": "$step1"}, "invalid_arguments", None),
    ({"text": "Bolt", "range": "Stock"}, "invalid_arguments", None),
    ({"text": "Bolt", "cursor": "nonsense"}, "invalid_arguments", "cursor"),
])
def test_find_text_errors(box, arguments, code, field):
    session = _session(box, ledger())
    result = _find(box, session, **arguments)
    assert not result.ok and result.error.code == code, result.error
    if field is not None:
        assert result.error.field == field


def test_pages_walk_every_match_once_and_cursors_belong_to_their_query():
    limits = Limits(max_list_items=4)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=limits) as toolbox:
        session = _session(toolbox, ledger())
        seen, cursor, pages = [], None, 0
        while True:
            result = _find(toolbox, session, text="Bolt", **({"cursor": cursor} if cursor else {}))
            assert result.ok and result.total == 15
            seen += [(m["kind"], m["address"]) for m in result.data["matches"]]
            pages += 1
            cursor = result.next_cursor
            if cursor is None:
                break
        assert seen == EVERY_BOLT and pages == 4
        first = _find(toolbox, session, text="Bolt")
        assert first.summary == "15 match(es); 1-4 given; more with next_cursor"
        other = _find(toolbox, session, text="Nut", cursor=first.next_cursor)
        assert other.error.code == "invalid_arguments" and other.error.field == "cursor"
        scoped = _find(toolbox, session, text="Bolt", range="Stock!", cursor=first.next_cursor)
        assert scoped.error.code == "invalid_arguments"
        past = _find(toolbox, session, text="Bolt", cursor="99.0." + first.next_cursor.split(".")[2])
        assert past.error.code == "invalid_arguments" and "past the end" in past.error.message


@pytest.mark.parametrize("limits", [Limits(), Limits(max_result_chars=3000, max_list_items=20)],
                         ids=["default", "tight"])
def test_the_largest_fixture_is_bounded_and_paged(large, limits):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=limits) as toolbox:
        session = _session(toolbox, large)
        first = _find(toolbox, session, text="a")
        assert first.ok and first.total > 10_000
        assert len(first.data["matches"]) <= limits.max_list_items
        assert len(first.to_text()) <= limits.max_result_chars
        assert sum(first.data["kinds"].values()) == first.total
        # a few pages on, still bounded, and none repeats the one before
        cursor, previous = first.next_cursor, first.data["matches"]
        for _ in range(3):
            page = _find(toolbox, session, text="a", cursor=cursor)
            assert len(page.to_text()) <= limits.max_result_chars and page.data["matches"]
            assert page.data["matches"] != previous
            cursor, previous = page.next_cursor, page.data["matches"]
        # long matches are cut in their context, so a page stays a page
        wide = _find(toolbox, session, text=".+", regex=True)
        assert len(wide.to_text()) <= limits.max_result_chars
        assert all(len(m["context"]) <= 2 * MATCH_SHOWN + 2 for m in wide.data["matches"])


def test_find_text_is_the_shared_definition_and_merges_with_another_kind():
    tool = next(t for t in TOOLS if t.name == "find_text")
    assert tool.same_definition(shared.definition("find_text")) and tool.kinds == ["xlsx"]
    other = shared.handler("find_text", kind="docx")(lambda call, doc, text, **_: Result(summary="docx"))
    with Toolbox([*TOOLS, other], formats=[FORMAT], groups=GROUPS, workers=0) as both:
        assert sorted(both.tools["find_text"].kinds) == ["docx", "xlsx"]


def test_found_is_exported():
    assert Found.__module__ == "xlsx_agent.find"


# -- content stays out of diagnostics ----------------------------------------------------------


def test_no_match_text_reaches_a_diagnostic(box, caplog):
    caplog.set_level(logging.DEBUG)
    session = _session(box, ledger())
    results = [
        _find(box, session, text=MARKER),
        _find(box, session, text=f"Bolt {MARKER}", range="t:Bolts"),
        _find(box, session, text=f"(?P<{MARKER}", regex=True),           # refused
        _find(box, session, text=MARKER, range="Nowhere!"),              # refused
        _find(box, session, text=MARKER, range=f"t:Bolts/{MARKER}"),     # a key: refused
        _find(box, session, text=MARKER, cursor="0.0.00000000"),         # refused
    ]
    assert results[0].ok and {m["address"] for m in results[0].data["matches"]} == {
        "Stock!B4", "Stock!B3", "Stock!E1"}
    assert MARKER in json.dumps(results[0].data)                  # a result: content
    assert [r.ok for r in results] == [True, True, False, False, False, False]
    diagnostics = [r.getMessage() for r in caplog.records]
    diagnostics += [w for r in results for w in r.warnings] + [r.summary for r in results]
    diagnostics += [json.dumps(r.error.to_json()) for r in results if not r.ok]
    diagnostics += [json.dumps(record.__dict__, default=str) for record in session.log]
    assert diagnostics and not [d for d in diagnostics if MARKER in d]
