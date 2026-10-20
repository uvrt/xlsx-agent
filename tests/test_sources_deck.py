"""A workbook's tables in a deck (X1, ooxml-edit's O3), across formats: one session holds a
made-up workbook and a deck, and pptx-agent's ``ppt_add_table`` and the shared
``edit_chart`` take their rows from ``t:checks``.  The result names the workbook cells
each placed value came from -- by key where the table has one, else by cell address,
merged -- and the workbook is only read: same version, same bytes.

pptx-agent is a test-only dependency (the ``test`` extra; CI installs it from its
repository, as it does ooxml-edit).
"""

from __future__ import annotations

import datetime as dt
import json

import pytest
from ooxml_common.locales import Locale
from ooxml_edit.tools import Dataset, DatasetEntry, Limits, Toolbox, strict_problems
from test_sources import book

pptx_agent = pytest.importorskip("pptx_agent")
pptx_tools = pytest.importorskip("pptx_agent.tools")

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
BOX = {"x": 40, "y": 60, "w": 600, "h": 300}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*pptx_tools.TOOLS, *TOOLS], formats=[pptx_tools.FORMAT, FORMAT],
                 groups=[*pptx_tools.GROUPS, *GROUPS], workers=0) as box:
        yield box


def _session(toolbox, **options):
    session = toolbox.session(clock=CLOCK, **options)
    deck = pptx_agent.Document.new()
    deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")                           # d1
    session.open(book(), "checks.xlsx")                                  # d2
    return session


@pytest.fixture
def session(toolbox):
    return _session(toolbox)


def call(toolbox, session, name, **arguments):
    return toolbox.dispatch(session, name, arguments)


def ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
    return result


def failed(result, code):
    assert not result.ok, json.dumps(result.to_json(), indent=1, default=str)
    assert result.error.code == code, result.to_json()
    return result.error


def chart(toolbox, session, **arguments):
    return call(toolbox, session, "edit_chart",
                **{"doc": "d1", "target": "s:256", "action": "add", "chart_type": "column",
                   "box": BOX, **arguments})


def test_a_table_and_a_chart_from_one_workbook_table(toolbox, session):
    workbook = session.entry("d2")
    version, before = workbook.version, workbook.document.to_bytes()

    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "t:checks", "from_columns": ["code", "rating", "due"]})).data
    assert table["cells"] == [["code", "rating", "due"], ["A1", "3.5", "1/5/2026"],
                              ["A2", "4", "2/10/2026"], ["B1", "2.5", "3/15/2026"]]
    placed = table["from"]
    assert placed["source"] == "d2/t:checks" and placed["doc"] == "d2"
    assert placed["version"] == version
    assert [c["sources"] for c in placed["columns"]] == [
        ["t:checks/{A1,A2,B1}/code"], ["t:checks/{A1,A2,B1}/rating"],
        ["t:checks/{A1,A2,B1}/due"]]

    drawn = ok(chart(toolbox, session, **{"from": "t:checks",
                                          "columns": ["code", "rating"]})).data
    assert drawn["series"] == [{"name": "rating", "values": [3.5, 4, 2.5]}]
    placed = drawn["from"]
    assert placed["categories"] == {"column": "code", "rows": "1-3",
                                    "sources": ["t:checks/{A1,A2,B1}/code"]}
    assert placed["series"] == [{"name": "rating", "column": "rating", "rows": "1-3",
                                 "sources": ["t:checks/{A1,A2,B1}/rating"]}]

    assert workbook.version == version                     # only read
    assert workbook.document.to_bytes() == before
    deck = session.entry("d1").document.to_bytes()
    assert b"t:checks" not in deck                         # provenance stays in the result


def test_a_key_set_places_its_rows_in_its_order(toolbox, session):
    drawn = ok(chart(toolbox, session, **{"from": "t:checks/{B1,A1}",
                                          "columns": ["code", "rating"]})).data
    assert drawn["categories"] == ["B1", "A1"]
    assert drawn["from"]["series"][0]["sources"] == ["t:checks/{B1,A1}/rating"]


def test_a_row_range_places_those_rows_and_names_their_cells(toolbox, session):
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "t:checks/#2:", "from_columns": ["code", "rating"]})).data
    assert table["cells"] == [["code", "rating"], ["A2", "4"], ["B1", "2.5"]]
    assert table["from"]["source"] == "d2/t:checks/#2:"
    assert table["from"]["columns"][1]["sources"] == ["t:checks/{A2,B1}/rating"]
    error = failed(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                        **{"from": "t:checks/#2:#5"}), "not_found")
    assert error.valid_options == ["t:checks/#2:#3"]


def test_a_table_without_a_key_names_its_cells_merged(toolbox, session):
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "d2/t:visits", "from_columns": ["area", "rating"]})).data
    assert [c["sources"] for c in table["from"]["columns"]] == [
        ["'Visit plan'!B3:B5"], ["'Visit plan'!C3:C5"]]
    assert table["cells"][2] == ["north", "4.5"]           # text, as stored


def test_text_in_a_series_is_refused_never_coerced(toolbox, session):
    error = failed(chart(toolbox, session, **{"from": "t:visits",
                                              "columns": ["area", "rating"]}),
                   "invalid_arguments")
    assert "'Visit plan'!C4" in error.message and "4.5" not in error.message


def test_the_workbooks_refusals_reach_the_deck_tools(toolbox, session):
    error = failed(chart(toolbox, session, **{"from": "t:chekcs",
                                              "columns": ["code", "rating"]}), "not_found")
    assert "t:checks" in error.valid_options
    error = failed(chart(toolbox, session, **{"from": "t:checks/A1/rating",
                                              "columns": ["code", "rating"]}),
                   "invalid_arguments")
    assert "t:checks[rating]" in error.valid_options


def test_a_table_over_the_sessions_limit_is_refused(toolbox):
    session = _session(toolbox, limits=Limits(max_dataset_cells=10))
    error = failed(chart(toolbox, session, **{"from": "t:checks",
                                              "columns": ["code", "rating"]}), "limit")
    assert "15 values" in error.message
    ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
            **{"from": "t:checks/*/rating"}))                # one column: 3 values


#: A dataset made in Python, its dates Python's (made up).
SITES = [{"site": "Ash", "opened": dt.date(2026, 3, 1), "at": dt.datetime(2026, 4, 15, 9, 30),
          "visits": 4},
         {"site": "Oak", "opened": dt.date(2026, 1, 20), "at": None, "visits": 7}]
DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")


def _keep_sites(session):
    session.keep_datasets([DatasetEntry("sites", Dataset.from_records(SITES))])


@pytest.mark.parametrize("locale, cells", [
    (None, [["Ash", "3/1/2026", "4/15/2026 09:30", "4"], ["Oak", "1/20/2026", "", "7"]]),
    (DUTCH, [["Ash", "1/3/2026", "15/4/2026 09:30", "4"], ["Oak", "20/1/2026", "", "7"]]),
])
def test_python_dates_are_placed_in_the_sessions_locale(toolbox, locale, cells):
    session = _session(toolbox, locale=locale)
    _keep_sites(session)
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "d:sites"})).data
    assert table["cells"] == [["site", "opened", "at", "visits"], *cells]
    drawn = ok(chart(toolbox, session, **{"from": "d:sites",
                                          "columns": ["opened", "visits"]})).data
    assert drawn["categories"] == [row[1] for row in cells]   # a category, not a blank


@pytest.mark.parametrize("locale, cells", [(None, ["1/20/2026", "3/1/2026"]),
                                           (DUTCH, ["20/1/2026", "1/3/2026"])])
def test_a_derive_of_python_dates_places_its_dates_in_the_locale(toolbox, locale, cells):
    session = _session(toolbox, locale=locale)
    _keep_sites(session)
    ok(call(toolbox, session, "xlsx_derive", **{"from": "d:sites"}, id="span",
            measures=[{"fn": "min", "column": "opened", "as": "since"},
                      {"fn": "first", "column": "opened", "as": "first"}]))
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "d:span"})).data
    assert table["cells"] == [["since", "first"], cells]
    # The derive's own result keeps ISO 8601: only placing writes the locale's.
    assert session.dataset("span").dataset.rows == (("2026-01-20", "2026-03-01"),)


#: The checks' due dates (book(), a date column) as each locale writes them.
DUE = {"en-US": ["1/5/2026", "2/10/2026", "3/15/2026"],
       "nl-BE": ["5/1/2026", "10/2/2026", "15/3/2026"]}


@pytest.mark.parametrize("locale, key", [(None, "en-US"), (DUTCH, "nl-BE")])
def test_a_sheets_date_column_reads_as_python_dates_do(toolbox, locale, key):
    session = _session(toolbox, locale=locale)
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "t:checks", "from_columns": ["due"]})).data
    assert [row[0] for row in table["cells"][1:]] == DUE[key]
    assert table["from"]["columns"][0]["sources"] == ["t:checks/{A1,A2,B1}/due"]
    drawn = ok(chart(toolbox, session, **{"from": "t:checks", "columns": ["due", "rating"]}))
    assert drawn.data["categories"] == DUE[key]
    # An older deck's chart, its categories ISO 8601 and in another order: matched by date.
    older = ok(chart(toolbox, session, categories=["2026-03-15", "2026-01-05", "2026-02-10"],
                     data=[{"name": "rating", "values": [0, 0, 0]}])).created[0]
    for target, values in ((drawn.created[0], [3.5, 4, 2.5]), (older, [2.5, 3.5, 4])):
        ok(call(toolbox, session, "edit_chart", doc="d1", target=target, action="set_values",
                series="rating", **{"from": "t:checks", "columns": ["due", "rating"]}))
        read = ok(call(toolbox, session, "edit_chart", doc="d1", target=target,
                       action="read"))
        assert read.data["series"][0]["values"] == values


def test_the_workbook_definitions_are_unchanged_and_strict_clean():
    assert strict_problems(TOOLS) == []
    assert not any(tool.sources for tool in TOOLS)         # no workbook tool takes from


# -- formula freshness across the formats (X3) -------------------------------------------------


def test_a_deck_table_from_edited_formula_cells_is_placed_stale_with_its_caches(toolbox):
    """The shop's item ids and line values are formulas: once a cell of the workbook is
    written, every cache is invalidated, and the deck's table says so -- stale, with the
    counts -- while a dataset derived before the edit goes stale and comes back with undo."""
    from shop import Shop

    session = toolbox.session(clock=CLOCK)
    deck = pptx_agent.Document.new()
    deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")                           # d1
    session.open(Shop().to_bytes(), "stock.xlsx")                        # d2
    source = {"from": "t:Stock!A9", "from_columns": ["Id", "Value"]}
    fresh = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **source)).data["from"]
    # every formula cell of the rows read: the ids, the line values and the tax rates
    assert fresh["cache"] == {"cached_unverified": 15} and "stale" not in fresh
    ok(call(toolbox, session, "xlsx_derive", doc="d2", **{"from": "t:Stock!A9"}, id="before",
            group_by=["Category"], measures=[{"fn": "sum", "column": "Value", "as": "total"}]))
    ok(call(toolbox, session, "xlsx_set_cells", doc="d2",
            cells=[{"address": "Stock!D10", "number": 7}]))
    placed = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                     **source)).data["from"]
    assert placed["stale"] is True and placed["cache"] == {"invalidated": 15}
    after = ok(call(toolbox, session, "xlsx_derive", doc="d2", **{"from": "t:Stock!A9"},
                    id="after", group_by=["Category"],
                    measures=[{"fn": "sum", "column": "Value", "as": "total"}]))
    assert after.data["cache"]["invalidated"] >= 5
    assert any("d:after is stale" in w and "invalidated" in w for w in after.warnings)
    assert session.dataset("d:before").stale and session.dataset("d:after").stale
    drawn = ok(chart(toolbox, session, **{"from": "d:before", "columns": ["Category", "total"]}))
    assert drawn.data["from"]["stale"] is True
    ok(call(toolbox, session, "undo", doc="d2"))
    assert not session.dataset("d:before").stale      # undo back: current again
    assert session.dataset("d:after").stale           # computed from invalidated caches
