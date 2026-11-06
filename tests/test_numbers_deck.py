"""A workbook's values for a check of numbers in text (round 36, R28; ooxml-edit 0.83.0).

The xlsx format lists every Excel table and declared table as a value source
(``value_sources``), and a table's dataset says each column's number format
(``Dataset.formats``) -- a derived dataset too, for the columns that come straight from a
source's cells.  End to end: a made-up deck whose text quotes a made-up workbook, checked
with ``include: ["numbers"]`` -- a slip found with the nearest values, a typed figure
matched within the rounding the table shows, its fix written by reference and then left
out.  pptx-agent is a test-only dependency (as in :mod:`test_sources_deck`).
"""

from __future__ import annotations

import datetime as dt
import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.docformat import value_sources

pptx_agent = pytest.importorskip("pptx_agent")
pptx_tools = pytest.importorskip("pptx_agent.tools")
if getattr(pptx_tools.FORMAT, "number_texts", None) is None:     # pragma: no cover
    pytest.skip("needs pptx-agent 0.1.17 (a deck's texts for numbers)", allow_module_level=True)

CLOCK = lambda: dt.datetime(2026, 11, 6, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
COLUMNS = ("topic", "mean", "share", "count")
ROWS = [("door", 2.66, 0.4512, 41), ("roof", 3.149, 0.2, 12), ("wall", 1.5, 0.3488, 30)]


def book(*, keyed: bool = False) -> bytes:
    """``t:Scores``: a mean shown ``0.0``, a share ``0.0%``, a count in General."""
    built = WorkbookBuilder()
    sheet = built.sheet("Scores")
    one, percent = built.style(num_fmt="0.0"), built.style(num_fmt="0.0%")
    sheet.rows("A1", [COLUMNS])
    for row, (topic, mean, share, count) in enumerate(ROWS, start=2):
        sheet.cell(f"A{row}", topic)
        sheet.cell(f"B{row}", mean, style=one)
        sheet.cell(f"C{row}", share, style=percent)
        sheet.cell(f"D{row}", count if not keyed else "n/a")   # keyed: topic alone is unique
    sheet.table("Scores", "A1:D4", COLUMNS, write_header=False)
    return built.to_bytes()


# -- the formats and the sources -----------------------------------------------------------------


def test_every_table_is_a_value_source_and_its_dataset_says_its_formats():
    workbook = Workbook.open(book())
    assert value_sources(workbook) == ["t:Scores"]
    data = workbook.dataset("t:Scores")
    assert data.formats == (None, "0.0", "0.0%", None)          # General is None
    assert workbook.dataset("t:Scores[mean]").formats == ("0.0",)


def test_a_derived_column_straight_from_a_cell_keeps_its_format():
    workbook = Workbook.open(book())
    top = workbook.derive(from_="t:Scores", measures=[
        {"fn": "max", "column": "mean", "as": "top"}, {"fn": "mean", "column": "mean", "as": "avg"}])
    assert top.dataset().formats == ("0.0", None)               # a mean is computed: none
    grouped = workbook.derive(from_="t:Scores", group_by=["topic"],
                              measures=[{"fn": "first", "column": "share", "as": "s"}])
    assert grouped.dataset().formats == (None, "0.0%")
    again = workbook.derive(from_=grouped, measures=[{"fn": "max", "column": "s", "as": "m"}])
    assert again.dataset().formats == ("0.0%",)                 # from a dataset's formats
    joined = workbook.derive(from_="t:Scores", join={"from": "t:Scores", "on": ["topic", "topic"],
                                                     "columns": ["mean"]})
    assert joined.columns[-1] == "mean (Scores)"
    assert joined.dataset().formats == (None, "0.0", "0.0%", None, "0.0")


def test_mixed_formats_in_a_column_say_none():
    built = WorkbookBuilder()
    sheet = built.sheet("S")
    sheet.rows("A1", [("k", "v")])
    sheet.cell("A2", "a")
    sheet.cell("B2", 1.25, style=built.style(num_fmt="0.0"))
    sheet.cell("A3", "b")
    sheet.cell("B3", 2.5, style=built.style(num_fmt="0.00"))
    sheet.table("Mixed", "A1:B3", ("k", "v"), write_header=False)
    assert Workbook.open(built.to_bytes()).dataset("t:Mixed").formats == (None, None)


# -- a deck's text against the workbook ----------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*pptx_tools.TOOLS, *TOOLS], formats=[pptx_tools.FORMAT, FORMAT],
                 groups=[*pptx_tools.GROUPS, *GROUPS], workers=0) as box:
        yield box


def call(toolbox, session, name, **arguments):
    result = toolbox.dispatch(session, name, arguments)
    assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
    return result


def _session(toolbox, text, **workbook):
    session = toolbox.session(clock=CLOCK)
    deck = pptx_agent.Document.new()
    slide = deck.add_slide("Blank")
    slide.add_textbox(508000, 762000, 7620000, 2540000, text)
    session.open(deck.to_bytes(), "deck.pptx")                           # d1
    session.open(book(**workbook), "scores.xlsx")                        # d2
    return session


def numbers(toolbox, session):
    return call(toolbox, session, "check", doc="d1", include=["numbers"]).data["numbers"]


def test_a_slip_and_a_typed_figure_against_the_open_workbook(toolbox):
    session = _session(toolbox, "The door rated 2.7 of 5, named by 45.1% of answers; the roof 3.2.")
    found = numbers(toolbox, session)
    assert found["values"] == 9 and (found["unverified"], found["typed"]) == (1, 2)
    text = found["texts"][0]
    assert text["unverified"][0]["text"] == "3.2"
    assert text["unverified"][0]["near"][0] == {"value": "3.1", "from": "t:Scores/#2/mean",
                                                "sources": ["Scores!B3"]}
    typed = {entry["text"]: entry for entry in text["typed"]}
    # The table has no one key column (topic and count are both unique): rows by position.
    assert typed["2.7"]["from"] == "t:Scores/#1/mean"
    assert typed["45.1%"]["from"] == "t:Scores/#1/share"
    fixes = [entry["fix"] for entry in text["typed"]]
    call(toolbox, session, "batch", ops=[{"tool": f["tool"], "arguments": f["arguments"]}
                                         for f in fixes])
    after = numbers(toolbox, session)
    assert (after["unverified"], after["typed"]) == (1, 0)      # written by reference now
    assert session.entry("d2").version == 0                     # the workbook only read


def test_a_keyed_table_is_quoted_by_key(toolbox):
    session = _session(toolbox, "The wall rated 1.5.", keyed=True)
    typed = numbers(toolbox, session)["texts"][0]["typed"][0]
    assert typed["from"] == "t:Scores/wall/mean"
    assert typed["fix"]["arguments"]["replace"].endswith("{{t:Scores/wall/mean:0.0}}.")
