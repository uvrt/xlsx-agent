"""Round 41: ``text_chars``, the characters a cell's text may have in a page of
``xlsx_read`` and ``xlsx_query`` -- of a table, of a sheet's cells, of a dataset -- and
``cut.text``, which says what was cut, the longest text's full length and the call that
gives the texts in full.  Pages keep their size budget and page as before.  Made-up
workbooks only (tests/build.py)."""

from __future__ import annotations

import pytest

from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from test_datasets import _session, call

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.read import TEXT_SHOWN

#: Each row's note: 2,500 to 9,500 characters, all past the default length.
LENGTHS = [2500, 4000, 5500, 7000, 8500, 9500]


def _note(i: int) -> str:
    return (f"note {i} " + "lorem ipsum " * 1000)[:LENGTHS[i]]


def notes() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Log")
    rows = [["Code", "Note", "Size"]] + [[f"L-{i}", _note(i), LENGTHS[i]]
                                        for i in range(len(LENGTHS))]
    sheet.rows("A1", rows)
    sheet.table("Log", f"A1:C{len(rows)}", rows[0], write_header=False)
    sheet.cell("E1", "short")
    sheet.cell("F1", "x" * 30000)
    return book


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return _session(box, notes())


def _pages(box, session, name, arguments):
    """Every page of a call, following next_cursor: the results."""
    out = [call(box, session, name, **arguments)]
    while out[-1].next_cursor:
        assert out[-1].ok, out[-1].error
        out.append(call(box, session, name, **{**arguments, "cursor": out[-1].next_cursor}))
    assert out[-1].ok, out[-1].error
    return out


def _notes_given(pages) -> dict[str, dict]:
    return {cell["a"]: cell for page in pages for row in page.data["rows"] for cell in row
            if cell["a"].startswith(("Log!B", "d:"))}


def test_by_default_text_is_cut_at_2000_with_its_length_and_the_call_for_the_rest(box, session):
    read = call(box, session, "xlsx_read", doc="d1", target="t:Log")
    cut = read.data["cut"]["text"]
    assert cut["count"] == len(LENGTHS) and cut["first"] == ["Log!B2", "Log!B3", "Log!B4"]
    assert cut["longest"] == max(LENGTHS)
    assert cut["call"] == {"doc": "d1", "target": "t:Log", "text_chars": max(LENGTHS)}
    first = read.data["rows"][0][1]
    assert len(first["v"]) == TEXT_SHOWN and first["length"] == LENGTHS[0]


def test_the_call_cut_text_gives_has_every_text_in_full_paged_within_the_budget(box, session):
    first = call(box, session, "xlsx_read", doc="d1", target="t:Log")
    pages = _pages(box, session, "xlsx_read", first.data["cut"]["text"]["call"])
    assert len(pages) > 1                   # longer texts: more pages, never bigger ones
    for page in pages:
        assert len(page.to_text()) <= box.limits.max_result_chars
        assert "text" not in page.data.get("cut", {})
    given = _notes_given(pages)
    assert [given[f"Log!B{i + 2}"]["v"] for i in range(len(LENGTHS))] == \
        [_note(i) for i in range(len(LENGTHS))]
    assert not any("length" in cell for cell in given.values())


def test_text_chars_cuts_at_its_length_and_lists_only_what_it_cut(box, session):
    read = call(box, session, "xlsx_read", doc="d1", target="t:Log", text_chars=5000,
                limit=3)
    row = {cell["a"]: cell for cell in read.data["rows"][0] + read.data["rows"][1]}
    assert row["Log!B2"]["v"] == _note(0) and "length" not in row["Log!B2"]
    assert len(read.to_text()) <= box.limits.max_result_chars
    cells = [c for r in read.data["rows"] for c in r if c["a"].startswith("Log!B")]
    cut = [c for c in cells if "length" in c]
    assert all(len(c["v"]) == 5000 and c["length"] > 5000 for c in cut)
    assert read.data["cut"]["text"]["first"] == [c["a"] for c in cut][:3]


def test_short_text_with_text_chars_is_as_without(box, session):
    plain = call(box, session, "xlsx_read", doc="d1", target="t:Log[Code]")
    longer = call(box, session, "xlsx_read", doc="d1", target="t:Log[Code]",
                  text_chars=50000)
    assert plain.data == longer.data and "cut" not in longer.data


def test_one_cell_is_cut_at_text_chars_when_given(box, session):
    whole = call(box, session, "xlsx_read", doc="d1", target="Log!B7", cells_only=True)
    assert whole.data["rows"][0][0]["v"] == _note(5)        # one cell: as the page holds
    short = call(box, session, "xlsx_read", doc="d1", target="Log!B7", cells_only=True, text_chars=100)
    cell = short.data["rows"][0][0]
    assert len(cell["v"]) == 100 and cell["length"] == LENGTHS[5]
    assert short.data["cut"]["text"]["call"]["text_chars"] == LENGTHS[5]


def test_a_max_chars_too_small_for_the_text_is_left_out_of_the_call(box, session):
    small = call(box, session, "xlsx_read", doc="d1", target="Log!B7", cells_only=True, max_chars=2000)
    assert len(small.to_text()) <= 2000
    again = small.data["cut"]["text"]["call"]
    assert "max_chars" not in again and again["text_chars"] == LENGTHS[5]
    assert call(box, session, "xlsx_read", **again).data["rows"][0][0]["v"] == _note(5)


def test_a_text_longer_than_a_page_is_cut_to_the_page_and_a_warning_says_so(box, session):
    read = call(box, session, "xlsx_read", doc="d1", target="Log!F1", text_chars=40000)
    cell = read.data["rows"][0][0]
    assert cell["length"] == 30000 and len(cell["v"]) < 30000
    assert len(read.to_text()) <= box.limits.max_result_chars
    assert read.data["cut"]["text"]["longest"] == 30000
    assert any("longer than a page holds" in w for w in read.warnings)


@pytest.mark.parametrize("alias", ["max_text_chars", "text_length"])
def test_another_name_for_text_chars_is_read_as_it(box, session, alias):
    read = call(box, session, "xlsx_read", doc="d1", target="t:Log", **{alias: 3000})
    assert read.ok and f"{alias} read as text_chars" in read.warnings
    assert len(read.data["rows"][0][1]["v"]) == 2500        # under 3000: whole
    query = call(box, session, "xlsx_query", doc="d1", table="t:Log", **{alias: 3000})
    assert query.ok and f"{alias} read as text_chars" in query.warnings
    both = call(box, session, "xlsx_read", doc="d1", target="t:Log", text_chars=10,
                **{alias: 10})
    assert both.error.code == "invalid_arguments"


def test_a_query_takes_text_chars_and_its_cut_gives_the_same_query(box, session):
    where = [{"column": "Size", "op": "gte", "value": "5000"}]
    query = call(box, session, "xlsx_query", doc="d1", table="t:Log", where=where,
                 columns=["Code", "Note"])
    cut = query.data["cut"]["text"]
    assert cut["count"] == 4 and cut["longest"] == 9500
    assert cut["call"] == {"doc": "d1", "table": "t:Log", "where": where,
                           "columns": ["Code", "Note"], "text_chars": 9500}
    pages = _pages(box, session, "xlsx_query", cut["call"])
    for page in pages:
        assert len(page.to_text()) <= box.limits.max_result_chars
        assert page.total == 4
    assert [c["v"] for c in _notes_given(pages).values()] == [_note(i) for i in range(2, 6)]
    short = call(box, session, "xlsx_query", doc="d1", table="t:Log", text_chars=10)
    assert all(len(row[1]["v"]) == 10 for row in short.data["rows"])


def test_a_dataset_read_and_queried_with_text_chars(box, session):
    derived = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Log"}, id="log",
                   group_by=["Code", "Note"])
    assert derived.ok, derived.error
    assert derived.data["cut"]["text"]["call"] == {"target": "d:log", "text_chars": 9500}
    read = call(box, session, "xlsx_read", target="d:log")
    cut = read.data["cut"]["text"]
    assert cut["count"] == len(LENGTHS) and cut["first"][0] == "d:log/#1/Note"
    assert cut["call"] == {"target": "d:log", "text_chars": 9500}
    pages = _pages(box, session, "xlsx_read", derived.data["cut"]["text"]["call"])
    assert [c["v"] for c in _notes_given(pages).values() if c["a"].endswith("Note")] == \
        [_note(i) for i in range(len(LENGTHS))]
    assert all(len(p.to_text()) <= box.limits.max_result_chars for p in pages)
    query = call(box, session, "xlsx_query", table="d:log", text_chars=100)
    assert all(len(row[1]["v"]) == 100 for row in query.data["rows"])
    assert query.data["cut"]["text"]["call"] == {"table": "d:log", "text_chars": 9500}
    one = call(box, session, "xlsx_read", target="d:log/#6/Note", text_chars=50)
    assert len(one.data["rows"][0][0]["v"]) == 50


def test_the_trace_takes_no_text_length(box):
    definition = next(t for t in TOOLS if t.name == "xlsx_trace")
    assert "text_chars" not in definition.schema["properties"]
