"""Round 31, R7: rows traced to the rows they rest on (:mod:`xlsx_agent.trace`) and the
``xlsx_trace`` tool (:mod:`xlsx_agent.tools.trace`).

Every workbook is made up here: a table of parts whose notes name other parts, a declared
table of tickets whose "Parts" cell lists the parts each concerns (a declared split), and a
table of crews with no column that can be a key.
"""

from __future__ import annotations

import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Result, Toolbox, shared

from xlsx_agent import Workbook
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format
from xlsx_agent.trace import MAX_DEPTH

TICKETS = {"name": "tickets", "sheets": "Tickets", "key": "Ticket",
           "fields": [{"header": "Parts", "clean": {"split": ";"}}]}

LONG = "see P-2 " + "x" * 3000


def book(notes4: str = "P-40 and p-1 are other codes; Alpha, 7") -> WorkbookBuilder:
    out = WorkbookBuilder()
    parts = out.sheet("Parts")
    parts.rows("A1", [["Code", "Name", "Group", "Weight", "Notes"],
                      ["P-1", "rotor", "core", 4, "needs P-2"],
                      ["P-2", "gasket", "trim", 2, "see P-3 and T-1"],
                      ["P-3", "rotor", "core", 1, None],
                      ["P-4", "clip", "core", 1, notes4]])
    parts.table("Parts", "A1:E5", ["Code", "Name", "Group", "Weight", "Notes"])
    tickets = out.sheet("Tickets")
    tickets.rows("A1", [["Ticket", "Parts", "State"], ["T-1", "P-1; P-3", "open"],
                        ["T-2", "P-1", "done"], ["T-3", "P-9", "open"]])
    crews = out.sheet("Crews")
    crews.rows("A1", [["Crew", "Part"], ["blue", "P-4"], ["blue", "P-4"]])
    crews.table("Crews", "A1:B3", ["Crew", "Part"])
    return out


@pytest.fixture
def wb() -> Workbook:
    opened = Workbook.open(book().to_bytes())
    opened.declare_tables([TICKETS])
    return opened


def _links(trace):
    return [(link.row, link.why, link.by, link.column, link.depth) for link in trace.links]


# -- the library -------------------------------------------------------------------------------


def test_a_row_names_rows_by_key_and_those_name_more_to_depth(wb):
    found = wb.trace("t:tickets/T-1")
    (trace,) = found.traces
    assert trace.row == "t:tickets/T-1"
    assert _links(trace) == [
        ("t:Parts/P-1", "names", None, "Parts", 1),
        ("t:Parts/P-3", "names", None, "Parts", 1),
        ("t:Parts/P-2", "names", "t:Parts/P-1", "Notes", 2),
        ("t:Parts/P-2", "lists", None, "Notes", 0)]
    # every value in full, each row once, blanks left out, a split's items as a list
    assert found.values("t:tickets/T-1") == {"Ticket": "T-1", "Parts": ["P-1", "P-3"],
                                             "State": "open"}
    assert found.values("t:Parts/P-3") == {"Code": "P-3", "Name": "rotor", "Group": "core",
                                           "Weight": 1}
    assert [link.as_dict() for link in trace.links][2] == {
        "row": "t:Parts/P-2", "why": "names", "by": "t:Parts/P-1", "column": "Notes",
        "depth": 2}


def test_depth_bounds_names_and_lists_name_the_rows_listing_a_key(wb):
    (trace,) = wb.trace("t:Parts/P-1", depth=1).traces
    assert _links(trace) == [
        ("t:Parts/P-2", "names", None, "Notes", 1),
        ("t:tickets/T-1", "lists", None, "Parts", 0),
        ("t:tickets/T-2", "lists", None, "Parts", 0)]
    (trace,) = wb.trace("t:Parts/P-1", depth=0).traces
    assert [link.why for link in trace.links] == ["lists", "lists"]
    with pytest.raises(ValueError, match=f"depth is 0 to {MAX_DEPTH}"):
        wb.trace("t:Parts/P-1", depth=MAX_DEPTH + 1)


def test_a_key_is_found_whole_case_counting_and_only_with_a_letter_and_a_digit(wb):
    """P-40 is not P-4, p-1 not P-1, and a row never links to itself (P-4's notes would
    name it); words and numbers are no keys looked for in text."""
    (trace,) = wb.trace("t:Parts/P-4").traces
    assert [link.why for link in trace.links] == ["lists", "lists"]     # the crews' rows
    other = Workbook.open(book(notes4="ask P-1, then P-3.").to_bytes())
    (trace,) = other.trace("t:Parts/P-4", depth=1).traces
    assert [link.row for link in trace.links if link.why == "names"] == [
        "t:Parts/P-1", "t:Parts/P-3"]


def test_rows_by_key_set_range_table_and_cell_and_a_table_without_a_key(wb):
    rows = lambda given: [t.row for t in wb.trace(given, depth=0).traces]  # noqa: E731
    assert rows("t:Parts/{P-3,P-1}") == ["t:Parts/P-3", "t:Parts/P-1"]
    assert rows("t:Parts/#2:#3") == ["t:Parts/P-2", "t:Parts/P-3"]
    assert rows("t:Parts") == ["t:Parts/P-1", "t:Parts/P-2", "t:Parts/P-3", "t:Parts/P-4"]
    assert rows("Parts!E3") == ["t:Parts/P-2"]      # a citation of a cell in a table row
    # a table with no key column: its rows by position; their text still names keys
    (trace,) = wb.trace("t:Crews/#1").traces
    assert trace.row == "t:Crews/#1" and [link.row for link in trace.links] == ["t:Parts/P-4"]
    (trace,) = wb.trace("t:Parts/P-4", depth=0).traces
    assert [(link.row, link.why) for link in trace.links] == [
        ("t:Crews/#1", "lists"), ("t:Crews/#2", "lists")]


def test_an_edit_or_a_declaration_is_seen_by_the_next_trace():
    opened = Workbook.open(book().to_bytes())
    (trace,) = opened.trace("t:Parts/P-3", depth=0).traces
    assert [link.row for link in trace.links] == ["t:Parts/P-2"]   # tickets not declared yet
    opened.declare_tables([TICKETS])
    (trace,) = opened.trace("t:Parts/P-3", depth=0).traces
    assert [link.row for link in trace.links] == ["t:Parts/P-2", "t:tickets/T-1"]


# -- the tool ----------------------------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[xlsx_format(tables={"tables": [TICKETS]})], groups=GROUPS,
                 workers=0) as toolbox:
        yield toolbox


def _session(box, data: bytes | None = None):
    session = box.session()
    session.open(data or book().to_bytes(), "parts.xlsx")
    return session


def test_xlsx_trace_is_on_demand_in_its_own_group(box):
    tool = box.tools["xlsx_trace"]
    assert tool.group == "xlsx_trace" and not tool.mutates
    core = [d["name"] for d in box.definitions("anthropic", groups="core")
            if not d.get("defer_loading")]
    assert "xlsx_trace" not in core


def test_a_dataset_row_rests_on_its_source_rows_and_what_they_name(box):
    session = _session(box)
    made = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "by_group", "group_by": ["Group"],
        "measures": [{"fn": "sum", "column": "Weight", "as": "weight"}]})
    assert made.ok, made.to_json()
    result = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["d:by_group/#2"],
                                                  "depth": 1})
    assert result.ok, result.to_json()
    (traced,) = result.data["traced"]
    assert traced["row"] == "d:by_group/#2"
    assert traced["links"] == [
        {"row": "t:Parts/P-2", "why": "source", "columns": ["Group", "weight"]},
        {"row": "t:Parts/P-3", "why": "names", "by": "t:Parts/P-2", "column": "Notes",
         "depth": 1},
        {"row": "t:tickets/T-1", "why": "names", "by": "t:Parts/P-2", "column": "Notes",
         "depth": 1},
        {"row": "t:Parts/P-1", "why": "lists", "by": "t:Parts/P-2", "column": "Notes"}]
    rows = result.data["rows"]
    assert rows["d:by_group/#2"] == {"Group": "trim", "weight": 2}
    assert rows["t:Parts/P-2"]["Notes"] == "see P-3 and T-1"
    assert set(rows) == {"d:by_group/#2", "t:Parts/P-2", "t:Parts/P-3", "t:tickets/T-1",
                         "t:Parts/P-1"}
    assert "left_out" not in result.data and "untraced" not in result.data
    assert "trim" not in result.summary and "Traced 1 of 1 row(s)" in result.summary
    # a whole dataset is its rows
    every = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["d:by_group"]})
    assert [t["row"] for t in every.data["traced"]] == ["d:by_group/#1", "d:by_group/#2"]


def test_a_dataset_is_traced_in_its_own_workbook_only(box):
    session = _session(box)
    session.open(book().to_bytes(), "other.xlsx")
    assert box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "by_group", "group_by": ["Group"]}).ok
    refused = box.dispatch(session, "xlsx_trace", {"doc": "d2", "rows": ["d:by_group/#1"]})
    assert refused.error.code == "invalid_arguments" and refused.error.field == "doc"
    assert refused.error.valid_options == ["d1"]
    missing = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["d:nothing"]})
    assert missing.error.code == "not_found" and missing.error.field == "rows[0]"


def test_every_value_in_full_and_what_did_not_fit_is_listed_to_ask_for(box):
    session = _session(box, book(notes4=LONG).to_bytes())
    whole = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Parts/P-4"]})
    assert whole.data["rows"]["t:Parts/P-4"]["Notes"] == LONG      # never cut
    small = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Parts"],
                                                 "max_chars": 1000})
    assert small.ok, small.to_json()
    data = small.data
    assert [t["row"] for t in data["traced"]] == ["t:Parts/P-1"]
    assert data["untraced"] == ["t:Parts/{P-2,P-3,P-4}"]
    assert data["left_out"] == ["t:Parts/{P-2,P-3}", "t:tickets/{T-1,T-2}"]
    assert data["next"].startswith("left_out: read them with xlsx_read")
    assert len(json.dumps(data)) < 1000
    # what is listed is asked for as listed
    read = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": data["left_out"][0]})
    assert read.ok and read.total == 2
    again = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": data["untraced"]})
    assert [t["row"] for t in again.data["traced"]] == ["t:Parts/P-2", "t:Parts/P-3",
                                                        "t:Parts/P-4"]


def test_rows_left_untraced_are_given_back_as_key_sets_and_row_ranges():
    from xlsx_agent.tools.trace import _key_sets, _runs

    assert _key_sets(_runs(["t:Crews/#2", "t:Crews/#3", "t:Crews/#4", "d:by_group/#5",
                            "t:Parts/P-1", "t:Parts/P-3", "Parts!A9"])) == [
        "t:Crews/#2:#4", "d:by_group/#5", "t:Parts/{P-1,P-3}", "Parts!A9"]


def test_a_column_is_no_rows_and_an_edit_is_seen(box):
    session = _session(box)
    refused = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Parts[Notes]"]})
    assert not refused.ok and "is a column: name rows" in refused.error.message
    outside = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["Parts!H40"]})
    assert "lies in no one table's data rows" in outside.error.message
    missing = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Parts/P-99"]})
    assert missing.error.code == "not_found" and "P-99" not in missing.error.message
    written = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Parts!E5", "text": "now with P-1"}]})
    assert written.ok, written.to_json()
    (traced,) = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Parts/P-4"],
                                                     "depth": 1}).data["traced"]
    assert traced["links"][0] == {"row": "t:Parts/P-1", "why": "names", "column": "Notes",
                                  "depth": 1}
    assert box.dispatch(session, "undo", {"doc": "d1"}).ok
    (traced,) = box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["t:Parts/P-4"],
                                                     "depth": 1}).data["traced"]
    assert [link["why"] for link in traced["links"]] == ["lists", "lists"]


# -- round 31, R5e: another kind's field at its default --------------------------------------


def _edit_version() -> tuple[int, ...]:
    import ooxml_edit

    return tuple(int(part) for part in ooxml_edit.__version__.split(".")[:3])


@pytest.mark.skipif(_edit_version() < (0, 77, 0), reason="needs ooxml-edit 0.77 (R5e)")
def test_find_text_on_a_workbook_ignores_a_documents_stories_at_its_default():
    """A toolbox that also serves documents is sent find_text with stories, and a strict
    decoder writes it; on a workbook, body and all ask nothing (ooxml-edit 0.77.0)."""
    other = shared.handler("find_text", kind="docx")(
        lambda call, doc, text, regex=False, slides=None, range=None, stories=None,
        cursor=None: Result(summary="none"))
    with Toolbox([*TOOLS, other], formats=[xlsx_format()], groups=GROUPS, workers=0) as box:
        session = _session(box)
        found = box.dispatch(session, "find_text", {"doc": "d1", "text": "rotor",
                                                    "stories": "body", "slides": [],
                                                    "regex": False})
        assert found.ok, found.to_json()
        assert found.total == 2
        assert found.warnings == [
            "slides ignored: it applies to pptx, not to this xlsx document",
            "stories ignored: it applies to pptx and docx, not to this xlsx document"]
        refused = box.dispatch(session, "find_text", {"doc": "d1", "text": "rotor",
                                                      "stories": "layouts"})
        assert refused.error.field == "stories"
