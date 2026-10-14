"""R11: datasets and ``xlsx_derive``.

A derive's result kept as ooxml-edit's ``Dataset`` (O2) under ``d:<id>``; ``d:`` read back by
``xlsx_read`` (its rows, a column, a row by position or key, a value) and taken by
``xlsx_derive``'s ``from``, with every value's sources carried through to the original
cells; a dataset that outlives closing its workbook, and is marked stale -- never
refreshed -- once its workbook changes.

**Staleness without edits.**  Workbooks have no edits before 0.2, and no history: the
session counts a change to such a document upwards (``DocumentEntry.bump``, which the
dispatcher calls after every changing call).  The tests here register a test-only changing
tool, ``xlsx_touch`` (``mutates=True``, changing nothing), beside the real ones, so the
change goes through a real ``Toolbox`` as an edit will.  Without a history there is no undo:
a dataset made stale stays stale (ooxml-edit's rule for such documents), which a test
states.  Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
import math
import pickle

import pytest

from build import WorkbookBuilder
from ooxml_edit.tools import Dataset, Limits, Result, Toolbox, string, tool
from test_catalogue import MARKER
from test_derive import checks, sales

from xlsx_agent import (AddressError, DatasetAddress, DatasetTable, DatasetValue, DeriveError,
                        NameNotFound, RegionError, Workbook, as_dataset, parse_address)
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.derive import SOURCES_SHOWN

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731

BY_REP = {"group_by": ["Region", "Rep"],
          "measures": [{"fn": "sum", "column": "Amount", "as": "total"}, {"fn": "count"}],
          "round": 1}


def interleaved(rows: int = 30, keyed: bool = False) -> WorkbookBuilder:
    """Two groups whose rows alternate, so each value's sources are many ranges; with
    ``keyed``, its notes are a key column (unique, text), so the sources are by key."""
    book = WorkbookBuilder()
    sheet = book.sheet("Log")
    sheet.rows("A1", [["Team", "Hours", "Note"]]
               + [["Red" if r % 2 else "Blue", r + 0.5,
                   f"n{r if keyed else r % 4} {MARKER}"] for r in range(1, rows + 1)])
    sheet.table("Log", f"A1:C{rows + 1}", ["Team", "Hours", "Note"], write_header=False)
    return book


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return Workbook.open(sales().to_bytes())


# -- the library -------------------------------------------------------------------------------


def test_a_derived_is_a_dataset_of_its_reported_rows_and_their_cells(wb):
    ds = wb.derive("t:Sales", id="by_rep", **BY_REP)
    data = ds.dataset()
    assert isinstance(data, Dataset) and data == as_dataset(ds)
    assert data.columns == ("Region", "Rep", "total", "count") and data.keys == 2
    assert data.types == ("text", "text", "number", "number")
    assert data.rows[0] == ("North", "Ann", 3.7, 2)          # rounded, as reported
    assert ds.unrounded[0]["total"] == 3.675                  # kept in the detail
    assert data.sources[0][2] == ("Sales!C2", "Sales!C5")
    assert data.detail is ds
    assert pickle.loads(pickle.dumps(data)) == data


def test_deriving_from_a_dataset_carries_provenance_to_the_original_cells(wb):
    by_rep = wb.derive("t:Sales", id="by_rep", **BY_REP)
    by_region = wb.derive(by_rep, group_by=["Region"],
                          measures=[{"fn": "sum", "column": "total", "as": "sum"},
                                    {"fn": "max", "column": "total"}, {"fn": "count"}],
                          id="by_region")
    assert by_region.address == "d:by_rep"
    # the unrounded values are measured: rounding happens once
    assert [r["sum"] for r in by_region.rows] == [3.675, 0.125, None, 0.0]
    table = DatasetTable(by_rep)
    for row, where in zip(by_region.unrounded, by_region.provenance):
        # the sources are the cells the dataset's values came from, merged again
        used = [p for p, r in enumerate(by_rep.unrounded, 1)
                if r["Region"] == row["Region"] and r["total"] is not None]
        cells = sorted({c for p in used for c in _cells(table.cell_sources(p, 2))})
        assert sorted(c for s in where["sum"].sources for c in _cells((str(s),))) == cells
        assert where["sum"].count == len(used)
        # ... and recompute the value from the original cells
        if cells:
            assert math.fsum(wb.cell(c).v for c in cells) == row["sum"]
        keys = [p for p, r in enumerate(by_rep.unrounded, 1) if r["Region"] == row["Region"]]
        assert sorted(c for s in where["Region"].sources for c in _cells((str(s),))) == \
            sorted({c for p in keys for c in _cells(table.cell_sources(p, 0))})
    # the values left out are named by their d: addresses, never their content
    assert by_region.excluded["sum"]["blank"].first == ("d:by_rep/#3/total",
                                                         "d:by_rep/#5/total")
    assert by_region.provenance[0]["sum"].as_dict()["sources"] == \
        ["Sales!C2:C3", "Sales!C5", "Sales!C7"]


def test_a_dataset_takes_filters_bins_pivots_and_sorts_as_a_table_does(wb):
    by_rep = wb.derive("t:Sales", id="by_rep", **BY_REP)
    found = wb.derive(by_rep, where=[{"column": "count", "op": "gte", "value": "2"},
                                     {"column": "total", "op": "not_blank"}],
                      bins={"column": "total", "edges": ["-1", "1", "5"]},
                      pivot="Region", sort=[("total", "desc")])
    assert found.matched == 4
    assert found.rows == [{"total": "[1, 5]", "North": 1, "South": 0, "West": 0},
                          {"total": "[-1, 1)", "North": 1, "South": 1, "West": 1}]
    with pytest.raises(DeriveError) as caught:
        wb.derive(by_rep, measures=[{"fn": "sum", "column": "Rep"}])
    assert caught.value.code == "wrong_type"


def test_a_dataset_of_dates_measures_its_iso_text_as_dates():
    book = WorkbookBuilder()
    sheet = book.sheet("Days")
    sheet.rows("A1", [["Kind", "Day"], ["a", dt.date(2026, 1, 1)], ["a", dt.date(2026, 1, 3)],
                      ["b", dt.date(2026, 2, 1)]])
    sheet.table("Days", "A1:B4", ["Kind", "Day"], write_header=False)
    workbook = Workbook.open(book.to_bytes())
    first = workbook.derive("t:Days", group_by=["Kind"],
                            measures=[{"fn": "min", "column": "Day", "as": "first"}])
    again = workbook.derive(first, measures=[{"fn": "max", "column": "first"}])
    assert first.types["first"] == "date" and again.rows == [{"max first": "2026-02-01"}]
    assert [str(s) for s in again.provenance[0]["max first"].sources] == ["Days!B2:B4"]


def test_a_foreign_dataset_keeps_its_opaque_sources():
    data = Dataset(("page", "words"), [(1, 10), (2, 5), (3, None)],
                   sources=[[("page:1",), ("page:1",)], [("page:2",), ("page:2",)],
                            [("page:3",), ("page:3",)]], keys=1)
    found = Workbook.open(sales().to_bytes()).derive(data, measures=[{"fn": "sum",
                                                                      "column": "words"}])
    assert found.rows == [{"sum words": 15.0}]
    assert found.provenance[0]["sum words"].sources == ("page:1", "page:2")
    assert found.excluded["sum words"]["blank"].first == ("d:derived/#3/words",)


@pytest.mark.parametrize("address,cells", [
    ("d:by_region", [f"d:by_region/#{p}/{c}" for p in (1, 2, 3, 4) for c in ("Region", "t")]),
    ("d:by_region[t]", [f"d:by_region/#{p}/t" for p in (1, 2, 3, 4)]),
    ("d:by_region/#2", ["d:by_region/#2/Region", "d:by_region/#2/t"]),
    ("d:by_region/West", ["d:by_region/#4/Region", "d:by_region/#4/t"]),
    ("d:by_region/South/t", ["d:by_region/#2/t"]),
    ("d:by_region/#3/region", ["d:by_region/#3/Region"]),
], ids=["dataset", "column", "row", "row-by-key", "value-by-key", "column-case-aside"])
def test_every_dataset_address_form_reads(wb, address, cells):
    by_region = wb.derive("t:Sales", group_by=["Region"],
                          measures=[{"fn": "sum", "column": "Amount", "as": "t"}],
                          id="by_region")
    rows = DatasetTable(by_region).read(address)
    assert [v.a for row in rows for v in row] == cells
    assert all(isinstance(v, DatasetValue) for row in rows for v in row)
    assert parse_address(cells[0]).id == "by_region"


def test_dataset_lookups_are_refused_with_their_codes(wb):
    by_rep = DatasetTable(wb.derive("t:Sales", id="by_rep", **BY_REP))
    with pytest.raises(RegionError) as caught:
        by_rep.read("d:by_rep/North")
    assert caught.value.code == "no_key_column" and "North" not in str(caught.value)
    with pytest.raises(RegionError) as caught:
        by_rep.read("d:by_rep/#7")
    assert caught.value.code == "position_out_of_range"
    with pytest.raises(NameNotFound) as caught:
        by_rep.read("d:by_rep[totals]")
    assert caught.value.what == "column" and caught.value.valid_options[0] == "total"
    twice = DatasetTable(Dataset(("k", "v"), [("a", 1), ("a", 2), (1, 3), (2.0, 4)], keys=1),
                         "twice")
    with pytest.raises(RegionError) as caught:
        twice.read("d:twice/a")
    assert caught.value.code == "duplicate_key"
    assert caught.value.valid_options == ["d:twice/#1", "d:twice/#2"]
    assert twice.read("d:twice/2/v")[0][0].v == 4          # a whole number by its digits
    with pytest.raises(NameNotFound) as caught:
        twice.read("d:twice/b")
    assert caught.value.what == "key" and "b" not in caught.value.message.split()


@pytest.mark.parametrize("text", ["d:2024", "d:by region", "d:" + "x" * 65, "d:a+b",
                                  "d:_ok?"])
def test_a_dataset_id_is_ooxml_edits(text):
    with pytest.raises(AddressError):
        parse_address(text)


@pytest.mark.parametrize("text", ["d:by_region", "d:_x", "d:a.b-c", "d:" + "x" * 64])
def test_ooxml_edits_dataset_ids_parse(text):
    assert parse_address(text).text == text
    with pytest.raises(ValueError):
        DatasetAddress("1x")


def test_a_d_address_needs_a_session_in_a_script(wb):
    with pytest.raises(DeriveError) as caught:
        wb.derive("d:by_rep")
    assert caught.value.code == "bad_source"


def _cells(sources) -> list[str]:
    """Every cell of range texts (one column each), as addresses."""
    out = []
    for text in sources:
        address = parse_address(text)
        area = address.area
        for row in range(area.first.row, area.last.row + 1):
            out.append(f"{address.sheet.text}!{chr(64 + area.first.column)}{row}")
    return out


# -- through a Toolbox -------------------------------------------------------------------------


@tool("xlsx_touch", "Test only: a change that changes nothing.", {"doc": string("Document id.")},
      mutates=True, kind="xlsx")
def xlsx_touch(call, doc):
    return Result(summary="touched")


@pytest.fixture
def box():
    with Toolbox([*TOOLS, xlsx_touch], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(toolbox, *books):
    session = toolbox.session(clock=CLOCK)
    for book in books:
        blob = session.add_blob(book.to_bytes(), name="book.xlsx")
        assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def call(toolbox, session, name, **arguments):
    return toolbox.dispatch(session, name, arguments)


def err(result) -> dict:
    return result.to_json()["error"]


def derive_by_rep(toolbox, session, **more):
    result = call(toolbox, session, "xlsx_derive", doc="d1", **{"from": "t:Sales"},
                  id="by_rep", **{**BY_REP, **more})
    assert result.ok, result.error
    return result


def test_xlsx_derive_keeps_a_dataset_and_shows_its_rows_with_their_sources(box):
    session = _session(box, sales())
    result = derive_by_rep(box, session)
    assert result.datasets == ["d:by_rep"] and result.total == 6
    data = result.data
    assert (data["dataset"], data["from"], data["key_columns"]) == ("d:by_rep", "t:Sales", 2)
    assert data["columns"] == ["Region", "Rep", "total", "count"]
    assert data["rounding"] == {"rule": "half_away_from_zero", "digits": {"total": 1}}
    assert data["excluded"]["total"]["number_as_text"] == {"count": 1, "first": ["Sales!C4"]}
    assert data["rows"][0][2] == {"a": "d:by_rep/#1/total", "v": 3.7, "t": "number",
                                  "sources": ["Sales!C2", "Sales!C5"]}
    assert "cut" not in data and "next" not in data
    assert result.summary == "Derived d:by_rep from t:Sales: 6 rows"
    entry = session.dataset("d:by_rep")
    assert (entry.doc, entry.stale, entry.closed) == ("d1", False, False)


def test_describe_of_any_workbook_lists_the_datasets_and_where_they_came_from(box):
    """Round 5, T11: a dataset's id, columns, types, rows and source, never its values."""
    session = _session(box, sales(), sales())
    derive_by_rep(box, session)
    again = call(box, session, "xlsx_derive", **{"from": "d:by_rep"}, id="by_region",
                 group_by=["Region"], measures=[{"fn": "sum", "column": "total"}])
    assert again.ok, again.error
    for doc in ("d1", "d2"):
        described = call(box, session, "describe", doc=doc)
        assert described.ok, described.error
        assert described.data["datasets"] == [
            {"id": "d:by_rep", "columns": ["Region", "Rep", "total", "count"],
             "types": {"Region": "text", "Rep": "text", "total": "number",
                       "count": "number"}, "rows": 6, "from": "d1/t:Sales"},
            {"id": "d:by_region", "columns": ["Region", "sum total"],
             "types": {"Region": "text", "sum total": "number"}, "rows": 4,
             "from": "d:by_rep"}]
    assert "North" not in json.dumps(described.data["datasets"])


def test_a_dataset_survives_closing_its_workbook_and_is_read_back_by_address(box):
    session = _session(box, sales())
    derive_by_rep(box, session)
    closed = call(box, session, "close_document", doc="d1")
    assert closed.ok and closed.data["datasets"] == ["d:by_rep"]
    read = call(box, session, "xlsx_read", target="d:by_rep[total]")
    assert read.ok, read.error
    assert read.data["dataset"] == {"id": "d:by_rep", "doc": "d1", "stale": False,
                                    "closed": True}
    assert [row[0]["v"] for row in read.data["rows"]] == [3.7, 0.0, None, 0.1, None, 0.0]
    assert read.warnings == []
    # derived from after its workbook closed: no doc needed, provenance carried through
    again = call(box, session, "xlsx_derive", **{"from": "d:by_rep"}, id="by_region",
                 group_by=["Region"], measures=[{"fn": "sum", "column": "total"}])
    assert again.ok, again.error
    assert again.data["rows"][0][1]["sources"] == ["Sales!C2:C3", "Sales!C5", "Sales!C7"]
    entry = session.dataset("d:by_region")
    assert (entry.doc, entry.closed, entry.stale) == ("d1", True, False)
    value = call(box, session, "xlsx_read", target="d:by_region/North/sum total")
    assert value.data["rows"] == [[{"a": "d:by_region/#1/sum total", "v": 3.675,
                                    "t": "number"}]]


def test_a_dataset_is_stale_once_its_workbook_changes_and_never_refreshed(box):
    session = _session(box, sales())
    derive_by_rep(box, session)
    call(box, session, "xlsx_derive", **{"from": "d:by_rep"}, id="by_region",
         group_by=["Region"], measures=[{"fn": "sum", "column": "total"}])
    before = call(box, session, "xlsx_read", target="d:by_rep")
    assert call(box, session, "xlsx_touch", doc="d1").ok      # the workbook's version moves on
    read = call(box, session, "xlsx_read", target="d:by_rep")
    assert read.ok and read.data["dataset"]["stale"] is True
    assert any("stale" in w for w in read.warnings)
    assert read.summary.endswith("(stale: its source has changed)")
    assert read.data["rows"] == before.data["rows"]           # never refreshed
    # a dataset derived from it goes stale with it
    assert session.dataset("d:by_region").stale
    assert call(box, session, "xlsx_read", target="d:by_region").data["dataset"]["stale"]
    # deriving from a stale dataset warns, and the new one is stale from the start
    derived = call(box, session, "xlsx_derive", **{"from": "d:by_rep"}, id="again",
                   measures=[{"fn": "count"}])
    assert derived.ok and any("stale" in w for w in derived.warnings)
    assert session.dataset("d:again").stale
    # a workbook has no history before 0.2 (no undo): once stale, stale for good, closed too
    assert call(box, session, "close_document", doc="d1").ok
    assert session.dataset("d:by_rep").stale and session.dataset("d:by_rep").closed
    # derived again from the workbook, as it is now: current
    session2 = _session(box, sales())
    derive_by_rep(box, session2)
    assert not session2.dataset("d:by_rep").stale


def test_a_change_to_another_workbook_leaves_it_current(box):
    session = _session(box, sales(), interleaved())
    derive_by_rep(box, session)
    assert call(box, session, "xlsx_touch", doc="d2").ok
    read = call(box, session, "xlsx_read", target="d:by_rep")
    assert read.data["dataset"]["stale"] is False and read.warnings == []


def test_a_value_shows_eight_ranges_and_pages_through_all_of_them(box):
    session = _session(box, interleaved())
    result = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Log"}, id="teams",
                  group_by=["Team"], measures=[{"fn": "sum", "column": "Hours"},
                                               {"fn": "share"}])
    assert result.ok, result.error
    hours = result.data["rows"][0][1]
    assert len(hours["sources"]) == SOURCES_SHOWN and hours["more"] == 15 - SOURCES_SHOWN
    assert result.data["totals"]["share"] == {"value": 30, "sources": ["Log!A2:A31"],
                                              "count": 30}
    red = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Log"}, id="red",
               where=[{"column": "Team", "op": "eq", "value": "Red"}],
               measures=[{"fn": "share", "column": "Hours"}])
    assert red.data["totals"]["share Hours"]["more"] == 15 - SOURCES_SHOWN
    assert len(red.data["totals"]["share Hours"]["sources"]) == SOURCES_SHOWN
    read = call(box, session, "xlsx_read", target="d:teams", sources=True)
    assert read.data["rows"][0][1]["more"] == 15 - SOURCES_SHOWN
    plain = call(box, session, "xlsx_read", target="d:teams")
    assert "sources" not in plain.data["rows"][0][1]
    given, cursor, pages = [], None, 0
    while True:
        page = call(box, session, "xlsx_read", target="d:teams/Red/sum Hours", sources=True,
                    limit=4, **({"cursor": cursor} if cursor else {}))
        assert page.ok, page.error
        assert page.total == 15 and page.data["value"]["v"] == math.fsum(
            r + 0.5 for r in range(1, 31, 2))
        given += page.data["sources"]
        pages += 1
        cursor = page.next_cursor
        if cursor is None:
            break
        assert page.data["cut"] == {"sources": 15 - len(given)}
    assert pages == 4 and given == [f"Log!B{r}" for r in range(2, 32, 2)]
    other = call(box, session, "xlsx_read", target="d:teams/Blue/sum Hours", sources=True,
                 cursor=page_cursor(box, session))
    assert not other.ok and err(other)["code"] == "invalid_arguments"


def _keys(sources: list[str]) -> list[str]:
    """The keys of key-set sources, in order."""
    out = []
    for text in sources:
        address = parse_address(text)
        out.extend(address.keys if address.keys is not None else [address.key])
    return out


def test_sources_by_key_show_eight_keys_and_page_through_all_of_them(box):
    session = _session(box, interleaved(keyed=True))
    result = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Log"}, id="teams",
                  group_by=["Team"], measures=[{"fn": "sum", "column": "Hours"},
                                               {"fn": "share"}])
    assert result.ok, result.error
    hours = result.data["rows"][0][1]
    note = f" {MARKER}"
    assert hours["sources"] == ["t:Log/{" + ",".join(f"n{r}{note}" for r in range(1, 17, 2))
                                + "}/Hours"]
    assert hours["more"] == 15 - SOURCES_SHOWN
    share = result.data["totals"]["share"]
    assert (share["count"], share["more"]) == (30, 30 - SOURCES_SHOWN)
    assert _keys(share["sources"]) == [f"n{r}{note}" for r in range(1, 9)]
    given, cursor, pages = [], None, 0
    while True:
        page = call(box, session, "xlsx_read", target="d:teams/Red/sum Hours", sources=True,
                    limit=4, **({"cursor": cursor} if cursor else {}))
        assert page.ok, page.error
        assert page.total == 15 and len(page.data["sources"]) == 1
        given += _keys(page.data["sources"])
        pages += 1
        cursor = page.next_cursor
        if cursor is None:
            break
        assert page.data["cut"] == {"sources": 15 - len(given)}
    assert pages == 4 and given == [f"n{r}{note}" for r in range(1, 31, 2)]
    # every key names the cell it was made from
    workbook = session.entry("d1").document
    cells = [workbook.cell(f"t:Log/n{r}{note}/Hours") for r in (1, 29)]
    assert [c.a for c in cells] == ["Log!B2", "Log!B30"]


def page_cursor(box, session) -> str:
    return call(box, session, "xlsx_read", target="d:teams/Red/sum Hours", sources=True,
                limit=4).next_cursor


def test_a_large_derive_is_bounded_and_continued_by_xlsx_read(box):
    book = WorkbookBuilder()
    sheet = book.sheet("Items")
    rows = 400
    sheet.rows("A1", [["Item", "Price"]] + [[f"item-{r:04d}", r * 1.25]
                                             for r in range(1, rows + 1)])
    sheet.table("Items", f"A1:B{rows + 1}", ["Item", "Price"], write_header=False)
    session = _session(box, book)
    result = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Items"}, id="items",
                  group_by=["Item"], measures=[{"fn": "sum", "column": "Price"}])
    assert result.ok and result.total == rows
    assert len(json.dumps(result.to_json(), ensure_ascii=False)) <= Limits().max_result_chars
    shown = len(result.data["rows"])
    assert 0 < shown <= Limits().max_list_items
    assert result.data["cut"]["rows"] == rows - shown
    assert result.data["next"]["target"] == "d:items"
    page = call(box, session, "xlsx_read", target="d:items", sources=True,
                cursor=result.data["next"]["cursor"])
    assert page.ok, page.error
    assert page.data["rows"][0][0]["a"] == f"d:items/#{shown + 1}/Item"
    assert len(json.dumps(page.to_json(), ensure_ascii=False)) <= Limits().max_result_chars


def test_xlsx_derive_refusals(box):
    session = _session(box, sales())
    derive_by_rep(box, session)
    cases = [
        ({"from": "d:nope", "id": "x"}, "not_found"),
        ({"from": "d:by_rep[total]", "id": "x"}, "invalid_arguments"),
        ({"from": "t:Sales", "id": "x"}, "invalid_arguments"),            # no doc
        ({"doc": "d1", "from": "t:Sales", "id": "2x"}, "invalid_arguments"),
        ({"doc": "d1", "from": "t:Sales", "id": "x",
          "measures": [{"fn": "sum", "column": "Rep"}]}, "refused"),
        ({"doc": "d1", "from": "t:Nope", "id": "x"}, "not_found"),
    ]
    for arguments, code in cases:
        result = call(box, session, "xlsx_derive", **arguments)
        assert not result.ok and err(result)["code"] == code, (arguments, err(result))
    missing = call(box, session, "xlsx_read", target="d:nope")
    assert err(missing)["code"] == "not_found" and err(missing)["valid_options"] == ["d:by_rep"]
    no_doc = call(box, session, "xlsx_read", target="Sales!A1")
    assert err(no_doc)["code"] == "invalid_arguments" and err(no_doc)["field"] == "doc"
    sources = call(box, session, "xlsx_read", doc="d1", target="Sales!A1", sources=True)
    assert err(sources)["field"] == "sources"
    key = call(box, session, "xlsx_read", target="d:by_rep/North")
    assert err(key)["code"] == "refused" and err(key)["details"]["code"] == "no_key_column"
    # nothing failed defined a dataset
    assert session.dataset_ids() == ["d:by_rep"]


def test_a_refused_column_of_dates_stored_as_text_names_the_issue_and_a_cell(box):
    session = _session(box, checks())
    measures = [{"fn": "min", "column": "Seen", "as": "earliest"}]
    refused = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Checks"}, id="seen",
                   measures=measures)
    error = err(refused)
    assert error["code"] == "refused" and error["field"] == "measures"
    assert error["details"] == {"code": "wrong_type", "column": "Seen", "column_type": "text",
                                "issue": "date_as_text",
                                "cells": {"count": 3, "first": ["Log!C2", "Log!C3", "Log!C5"]}}
    assert "Log!C2" in error["message"] and 'coerce ["date_as_text"]' in error["message"]
    # one step: the same call with the coercion it names
    read = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Checks"}, id="seen",
                measures=measures, coerce=["date_as_text"])
    assert read.ok, read.error
    assert read.data["types"]["earliest"] == "date"
    assert read.data["rows"][0][0]["v"] == "2026-02-11"


def test_the_limits_of_a_derive_are_limit_errors(box, monkeypatch):
    from xlsx_agent import derive as derive_module
    monkeypatch.setattr(derive_module, "MAX_GROUPS", 2)
    session = _session(box, sales())
    result = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Sales"}, id="x",
                  group_by=["Rep"])
    assert err(result)["code"] == "limit"


def test_no_cell_content_in_summaries_warnings_or_errors(box):
    session = _session(box, interleaved())
    results = [call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Log"}, id="notes",
                    group_by=["Note"]),
               call(box, session, "xlsx_read", target=f"d:notes/n1 {MARKER}", sources=True),
               call(box, session, "xlsx_read", target=f"d:notes/{MARKER}"),
               call(box, session, "xlsx_touch", doc="d1"),
               call(box, session, "xlsx_read", target="d:notes")]
    assert results[0].ok and results[1].ok and not results[2].ok
    assert MARKER in json.dumps(results[1].data)                 # content in results only
    for result in results:
        assert MARKER not in result.summary if result.ok else MARKER not in json.dumps(
            err(result))
        assert MARKER not in " ".join(result.warnings)
    for record in session.log:
        assert MARKER not in json.dumps(getattr(record, "summary", "") or "")


# -- datasets made in Python: dates ------------------------------------------------------------

from xlsx_agent.derive import derive  # noqa: E402

#: Made-up days, one blank among them.
OPENED = [{"site": "Ash", "opened": dt.date(2026, 3, 1), "at": dt.datetime(2026, 3, 1, 9, 30),
           "shift": dt.time(7, 15)},
          {"site": "Elm", "opened": None, "at": dt.datetime(2026, 4, 2, 17, 5, 30),
           "shift": dt.time(22, 0)},
          {"site": "Oak", "opened": dt.date(2026, 1, 20), "at": None, "shift": None}]


def test_python_dates_are_dates_without_a_label():
    data = Dataset.from_records(OPENED)
    measures = [{"fn": "min", "column": "opened", "as": "first_opened"},
                {"fn": "max", "column": "opened", "as": "last_opened"},
                {"fn": "first", "column": "opened", "as": "first_row"},
                {"fn": "max", "column": "at", "as": "latest"},
                {"fn": "min", "column": "shift", "as": "early"}]
    derived = derive(None, data, measures=measures)
    assert derived.rows == [{"first_opened": "2026-01-20", "last_opened": "2026-03-01",
                             "first_row": "2026-03-01", "latest": "2026-04-02T17:05:30",
                             "early": "07:15:00"}]
    assert derived.types == {"first_opened": "date", "last_opened": "date",
                             "first_row": "date", "latest": "datetime", "early": "time"}


def test_python_dates_read_back_as_iso_text_typed():
    table = DatasetTable(Dataset.from_records(OPENED), "sites")
    first = table.read("d:sites/#1")[0]
    assert [(v.v, v.t) for v in first] == [("Ash", "text"), ("2026-03-01", "date"),
                                           ("2026-03-01T09:30:00", "datetime"),
                                           ("07:15:00", "time")]
    assert json.dumps([v.as_dict() for v in first])            # plain JSON, no date objects
    when = dt.datetime(2026, 3, 1, 9, 30, 0, 999700, tzinfo=dt.timezone.utc)
    late = DatasetTable(Dataset.from_records([{"at": when}]), "late")
    assert late.read("d:late/#1/at")[0][0].v == "2026-03-01T09:30:01"   # to the millisecond


def test_dates_mixed_with_text_follow_the_majority():
    mostly = Dataset.from_records([{"d": dt.date(2026, 3, 1)}, {"d": dt.date(2026, 3, 9)},
                                   {"d": "soon"}])
    result = derive(None, mostly, measures=[{"fn": "max", "column": "d"}])
    assert result.rows == [{"max d": "2026-03-09"}]
    assert sum(c.count for c in result.excluded["max d"].values()) == 1
    few = Dataset.from_records([{"d": dt.date(2026, 3, 1)}, {"d": "soon"}, {"d": "later"}])
    with pytest.raises(DeriveError) as caught:
        derive(None, few, measures=[{"fn": "max", "column": "d"}])
    assert caught.value.code == "wrong_type"


def test_unlabelled_iso_text_is_text_and_the_refusal_says_how_to_read_it():
    rows = [{"site": "Ash", "opened": "2026-03-01"}, {"site": "Oak", "opened": "2026-01-20"}]
    measures = [{"fn": "min", "column": "opened"}]
    with pytest.raises(DeriveError) as caught:
        derive(None, Dataset.from_records(rows), measures=measures)
    error = caught.value
    assert error.code == "wrong_type" and error.issue == "date_as_text"
    assert "coerce [\"date_as_text\"]" in str(error)
    assert "types={'opened': \"date\"}" in str(error)
    labelled = derive(None, Dataset.from_records(rows, types={"opened": "date"}),
                      measures=measures)
    coerced = derive(None, Dataset.from_records(rows), measures=measures,
                     coerce=["date_as_text"])
    assert labelled.rows == coerced.rows == [{"min opened": "2026-01-20"}]
    # a column the caller says is text holds text: no flag, no way out offered
    with pytest.raises(DeriveError) as caught:
        derive(None, Dataset.from_records(rows, types={"opened": "text"}), measures=measures)
    assert caught.value.issue is None and "types=" not in str(caught.value)


def test_xlsx_derive_takes_a_python_dataset_of_dates(box):
    from ooxml_edit.tools import DatasetEntry
    session = _session(box)
    session.keep_datasets([DatasetEntry("sites", Dataset.from_records(OPENED))])
    result = call(box, session, "xlsx_derive", **{"from": "d:sites"}, id="span",
                  measures=[{"fn": "min", "column": "opened", "as": "since"},
                            {"fn": "max", "column": "opened", "as": "until"}])
    assert result.ok, result.error
    assert result.data["types"] == {"since": "date", "until": "date"}
    assert [v["v"] for v in result.data["rows"][0]] == ["2026-01-20", "2026-03-01"]
    read = call(box, session, "xlsx_read", target="d:sites/#1/opened")
    assert read.ok and read.data["rows"] == [[{"a": "d:sites/#1/opened", "v": "2026-03-01",
                                               "t": "date"}]]
