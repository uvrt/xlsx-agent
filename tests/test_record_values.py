"""Round 53, R108: values a caller judged from a column's text, one per row by key, kept as a
traced dataset (:mod:`xlsx_agent.record`, the deferred ``xlsx_record_values``).

Every workbook is made up (``shelf_book``): a "Shelf" table whose ``notes`` say how far each
item got, worded many ways -- English, French, a rich-text cell with a bold run -- and rows
with blank notes and notes of white space only.  The caller records an ``outcome`` per key.

Tool-design rules: 1 (a table address with a column read as the whole table, a whole-number
key matched as its digits), 2 (every refusal names each problem by place, the rows by
position and the loosely equal keys' rows), 3 (recording the same values again is no error),
4 (``not_recorded`` and ``counts`` capped, the rest counted, with the read that gives them),
5 (one call for every row).
"""

from __future__ import annotations

import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox
from shelf_book import JUDGED, ROWS, shelf_book

from xlsx_agent import RecordError, Workbook
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.record import LISTED

CALL = {"doc": "d1", "table": "t:Shelf", "key": "Ref", "source": "notes", "column": "outcome",
        "id": "outcome"}

#: Strings the refusals and summaries must never carry: cell content and the values given.
CONTENT = ("brackets", "Outcome", "Résultat", "terminé", "supplier", "partly done",
           "not started", "half printed", "S-0")


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, data: bytes | None = None):
    session = box.session()
    session.open(data or shelf_book().to_bytes(), "shelf.xlsx")
    return session


def ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1, default=str, ensure_ascii=False)
    return result


def record(box, session, values, **arguments):
    return box.dispatch(session, "xlsx_record_values", {**CALL, **arguments, "values": values})


# -- the library -------------------------------------------------------------------------------


def test_a_row_per_key_in_table_order_each_value_sourced_by_the_cell_judged():
    wb = Workbook.open(shelf_book().to_bytes())
    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                                values=list(reversed(JUDGED)))
    data = recorded.dataset
    assert data.columns == ("Ref", "outcome", "note") and data.keys == 1
    assert data.types == ("text", "text", "text")
    assert [row[0] for row in data.rows] == ["S-01", "S-02", "S-03", "S-04", "S-05", "S-08"]
    assert data.rows[3] == ("S-04", "done", "French: terminé")
    assert data.rows[0] == ("S-01", "done", None)
    # the key's own cell; the value and its note the cell judged (the rich one too: S-03)
    assert data.sources[2] == (("Shelf!A4",), ("Shelf!C4",), ("Shelf!C4",))
    assert recorded.not_recorded == (("S-09", "Shelf!C10"), ("S-010", "Shelf!C11"))
    assert recorded.no_key == ()
    assert dict(recorded.counts) == {"done": 3, "partly done": 2, "not started": 1}


def test_sources_are_by_key_where_the_table_has_one():
    out = WorkbookBuilder()
    sheet = out.sheet("Shelf")
    sheet.rows("A1", [["Ref", "Bay", "notes"], ["S-1", 4, "half"], ["S-2", 4, "half"]])
    sheet.table("Shelf", "A1:C3", ["Ref", "Bay", "notes"])
    wb = Workbook.open(out.to_bytes())
    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                                values=[{"key": "S-2", "value": "partly done"}])
    assert recorded.dataset.sources == ((("t:Shelf/S-2/Ref",), ("t:Shelf/S-2/notes",),
                                         ("t:Shelf/S-2/notes",)),)


def test_rows_with_text_and_no_key_are_listed_by_position():
    rows = [r[:] for r in ROWS]
    rows[8] = [None, "hooks", "no progress yet"]
    wb = Workbook.open(shelf_book(rows).to_bytes())
    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                                values=JUDGED)
    assert recorded.no_key == ("t:Shelf/#9",)
    # Ref has a blank now, so Item is the table's one key: the source is named by it
    assert recorded.not_recorded == (("S-010", "t:Shelf/shelf liners/notes"),)


def test_a_key_column_where_two_rows_share_a_key_is_refused_with_the_rows():
    rows = [r[:] for r in ROWS]
    rows[9][0] = "S-01"
    wb = Workbook.open(shelf_book(rows).to_bytes())
    with pytest.raises(RecordError) as caught:
        wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                         values=JUDGED)
    assert caught.value.code == "duplicate_keys" and caught.value.field == "key"
    assert caught.value.valid_options == ["t:Shelf/#1/Ref", "t:Shelf/#10/Ref"]


# -- the tool ----------------------------------------------------------------------------------


def test_xlsx_record_values_is_on_demand_in_its_own_group(box):
    tool = box.tools["xlsx_record_values"]
    assert tool.group == "xlsx_record" and not tool.mutates
    sent = {d["name"]: d for d in box.definitions("anthropic", defer=False)}
    assert not sent["xlsx_record_values"].get("strict")
    core = [d["name"] for d in box.definitions("anthropic", groups="core")
            if not d.get("defer_loading")]
    assert "xlsx_record_values" not in core


def test_the_result_counts_lists_not_recorded_and_the_read_that_gives_them(box):
    session = _session(box)
    version = session.entry("d1").version
    result = ok(record(box, session, JUDGED))
    data = result.data
    assert data["dataset"] == "d:outcome" and data["from"] == "t:Shelf"
    assert data["columns"] == ["Ref", "outcome", "note"] and data["recorded"] == 6
    assert data["counts"] == {"done": 3, "partly done": 2, "not started": 1}
    assert data["not_recorded"] == {
        "count": 2, "keys": ["S-09", "S-010"],
        "read": {"doc": "d1", "target": "t:Shelf/{S-09,S-010}", "key": "Ref",
                 "columns": ["Ref", "notes"]}}
    assert "replaced" not in data and not result.warnings
    assert result.summary == ("Recorded 6 value(s) as d:outcome from t:Shelf; 2 row(s) with "
                              "text in notes not recorded")
    # the read it gives works as given, and the workbook was only read
    read = ok(box.dispatch(session, "xlsx_read", data["not_recorded"]["read"]))
    assert [[c["v"] for c in row] for row in read.data["rows"]] == [
        ["S-09", "no progress yet"], ["S-010", "Mostly there: 8 of 10 bays."]]
    assert session.entry("d1").version == version


def test_the_dataset_reads_derives_and_traces_to_the_cell_judged(box):
    session = _session(box)
    ok(record(box, session, JUDGED))
    one = ok(box.dispatch(session, "xlsx_read", {"target": "d:outcome/S-04/outcome",
                                                 "sources": True}))
    assert one.data["value"]["v"] == "done" and one.data["sources"] == ["Shelf!C5"]
    derived = ok(box.dispatch(session, "xlsx_derive", {
        "from": "d:outcome", "id": "by_outcome", "group_by": ["outcome"],
        "sort": [{"column": "count", "desc": True}]}))
    assert [[c["v"] for c in row] for row in derived.data["rows"]] == [
        ["done", 3], ["partly done", 2], ["not started", 1]]
    assert derived.data["rows"][0][1]["sources"] == ["Shelf!C2", "Shelf!C5", "Shelf!C9"]
    where = ok(box.dispatch(session, "xlsx_derive", {
        "from": "d:outcome", "id": "done", "where": [{"column": "outcome", "op": "eq",
                                                      "value": "done"}]}))
    assert where.data["groups"] == 3
    traced = ok(box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["d:outcome/S-03"]}))
    (row,) = traced.data["traced"]
    assert row["links"] == [{"row": "t:Shelf/#3", "why": "source",
                             "columns": ["Ref", "outcome", "note"]}]
    assert traced.data["rows"]["t:Shelf/#3"]["notes"] == "Outcome: not started - supplier unknown"


def test_recording_again_replaces_and_says_so_and_an_edit_makes_it_stale(box):
    session = _session(box)
    ok(record(box, session, JUDGED))
    again = ok(record(box, session, JUDGED[:2]))
    assert again.data["replaced"] is True and again.data["recorded"] == 2
    assert again.warnings == ["d:outcome replaced: it holds this call's values only"]
    assert again.data["not_recorded"]["count"] == 6
    ok(box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Shelf!C3", "text": "All labels printed. Outcome: done"}]}))
    read = ok(box.dispatch(session, "xlsx_read", {"target": "d:outcome"}))
    assert read.data["dataset"]["stale"] is True


def test_every_problem_is_refused_at_once_by_place_with_the_rows_to_read(box):
    session = _session(box)
    bad = [{"key": "s-1", "value": "done"},            # loosely S-01
           {"key": "S-02", "value": "partly done"},
           {"key": "S-02", "value": "done"},           # twice
           {"key": "S-06", "value": "done"},           # blank notes
           {"key": "S-07", "value": "done"},           # notes of white space only
           {"key": "S-09", "value": " "},              # an empty value
           {"key": "S-10", "value": "done"},           # loosely S-010
           {"key": "X-99", "value": "done"}]           # nothing like it
    refused = record(box, session, bad)
    error = refused.error
    assert error.code == "not_found" and error.field == "values"
    assert error.details["code"] == "unknown_key"
    assert error.details["problems"] == [
        {"code": "unknown_key", "item": 0, "like": ["t:Shelf/#1/Ref"]},
        {"code": "key_twice", "item": 2, "row": "t:Shelf/#2/Ref", "first": 1},
        {"code": "empty_source", "item": 3, "row": "t:Shelf/#6/notes"},
        {"code": "empty_source", "item": 4, "row": "t:Shelf/#7/notes"},
        {"code": "empty_value", "item": 5, "row": "t:Shelf/#9/Ref"},
        {"code": "unknown_key", "item": 6, "like": ["t:Shelf/#10/Ref"]},
        {"code": "unknown_key", "item": 7, "like": []}]
    assert error.valid_options == ["t:Shelf/#1/Ref", "t:Shelf/#10/Ref"]
    assert "values[0], values[6], values[7]" in error.message
    assert "nothing was recorded" in error.message
    assert not any(text in error.message for text in CONTENT)
    assert "d:outcome" not in session.dataset_ids()
    # one kind alone gives its own code
    twice = record(box, session, [JUDGED[0], JUDGED[0]])
    assert twice.error.code == "invalid_arguments" and twice.error.details["code"] == "key_twice"
    empty = record(box, session, [{"key": "S-06", "value": "done"}])
    assert empty.error.code == "refused" and empty.error.details["code"] == "empty_source"


def test_a_taken_column_name_and_an_unknown_column_are_refused(box):
    session = _session(box)
    for column in ("note", "Ref", "ref"):
        refused = record(box, session, JUDGED, column=column)
        assert refused.error.code == "invalid_arguments" and refused.error.field == "column"
    missing = record(box, session, JUDGED, source="remarks")
    assert missing.error.code == "not_found" and "notes" in missing.error.valid_options
    duplicated = WorkbookBuilder()
    sheet = duplicated.sheet("Shelf")
    sheet.rows("A1", [["Ref", "notes"], ["S-1", "a"], ["S-1", "b"]])
    sheet.table("Shelf", "A1:B3", ["Ref", "notes"])
    other = _session(box, duplicated.to_bytes())
    refused = record(box, other, [{"key": "S-1", "value": "done"}])
    assert refused.error.code == "refused" and refused.error.details["code"] == "duplicate_keys"
    assert refused.error.valid_options == ["t:Shelf/#1/Ref", "t:Shelf/#2/Ref"]


def test_intent_a_column_address_and_a_whole_number_key(box):
    """Rule 1: t:Shelf[notes] is read as the table, with a note; a key column of whole numbers
    matches a key sent as its digits; recording the same values twice is no error (rule 3)."""
    session = _session(box)
    result = ok(record(box, session, JUDGED, table="t:Shelf[notes]"))
    assert result.data["from"] == "t:Shelf" and result.data["recorded"] == 6
    assert result.warnings == ["table: t:Shelf[notes] read as the whole table t:Shelf; key "
                               "and source name its columns"]
    numbers = WorkbookBuilder()
    sheet = numbers.sheet("Bays")
    sheet.rows("A1", [["Bay", "notes"], [101, "done"], [102, "half"]])
    sheet.table("Bays", "A1:B3", ["Bay", "notes"])
    other = _session(box, numbers.to_bytes())
    made = ok(record(box, other, [{"key": "102", "value": "partly done"}], table="t:Bays",
                     key="Bay"))
    assert made.data["recorded"] == 1 and made.data["not_recorded"]["keys"] == ["101"]
    row = ok(box.dispatch(other, "xlsx_read", {"target": "d:outcome/102"}))
    assert [c["v"] for c in row.data["rows"][0]] == [102, "partly done", None]
    same = ok(record(box, other, [{"key": "102", "value": "partly done"}], table="t:Bays",
                     key="Bay"))
    assert same.data["recorded"] == 1 and same.data["replaced"] is True


def test_not_recorded_and_counts_are_capped_and_counted(box):
    """Rule 4 and 5: 130 rows in one call; what the reply has no room for is counted."""
    rows = [[f"K-{i:03}", f"item {i}", f"note {i}. Outcome: grade {i}"] for i in range(130)]
    session = _session(box, shelf_book(rows).to_bytes())
    values = [{"key": f"K-{i:03}", "value": f"grade {i}"} for i in range(0, 130, 2)]
    result = ok(record(box, session, values))
    data = result.data
    assert data["recorded"] == 65
    assert len(data["counts"]) == LISTED and data["counts_more"] == 65 - LISTED
    assert data["not_recorded"]["count"] == 65 and len(data["not_recorded"]["keys"]) == LISTED
    assert data["not_recorded"]["keys"][:2] == ["K-001", "K-003"]
    assert len(json.dumps(result.to_json())) < 6000
    assert not any(text in result.summary for text in CONTENT)


def test_a_recorded_value_goes_into_a_word_document_by_reference():
    docx_agent = pytest.importorskip("docx_agent")
    docx_tools = pytest.importorskip("docx_agent.tools")
    with Toolbox([*docx_tools.TOOLS, *TOOLS], formats=[docx_tools.FORMAT, FORMAT],
                 groups=[*docx_tools.GROUPS, *GROUPS], workers=0) as both:
        session = both.session()
        session.open(shelf_book().to_bytes(), "shelf.xlsx")                       # d1
        session.open(docx_agent.Document.new().to_bytes(), "report.docx")         # d2
        ok(record(both, session, JUDGED))
        made = ok(both.dispatch(session, "word_insert_markdown", {
            "doc": "d2", "at": "end",
            "markdown": "Dividers: {{d:outcome/S-04/outcome}} ({{d:outcome/S-04/note}})."}))
        document = session.entry("d2").document
        texts = [document.paragraph(b).text for b in made.data["blocks"]]
        assert texts == ["Dividers: done (French: terminé)."]
        assert json.dumps(made.data["facts"], ensure_ascii=False).count("Shelf!C5") >= 1
