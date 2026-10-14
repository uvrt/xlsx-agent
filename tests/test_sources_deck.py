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
    assert table["cells"] == [["code", "rating", "due"], ["A1", "3.5", "2026-01-05"],
                              ["A2", "4", "2026-02-10"], ["B1", "2.5", "2026-03-15"]]
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


def test_a_derive_of_python_dates_places_its_iso_text(toolbox):
    session = _session(toolbox, locale=DUTCH)
    _keep_sites(session)
    ok(call(toolbox, session, "xlsx_derive", **{"from": "d:sites"}, id="span",
            measures=[{"fn": "min", "column": "opened", "as": "since"},
                      {"fn": "first", "column": "opened", "as": "first"}]))
    table = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    **{"from": "d:span"})).data
    assert table["cells"] == [["since", "first"], ["2026-01-20", "2026-03-01"]]


def test_the_workbook_definitions_are_unchanged_and_strict_clean():
    assert strict_problems(TOOLS) == []
    assert not any(tool.sources for tool in TOOLS)         # no workbook tool takes from
