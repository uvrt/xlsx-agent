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

Round 54 (R113 and the answers to round 53): a "section" heading row and an "item" row
sharing a reference, a spacer row repeating one (``SECTIONS``); rules 1 (rows without text
sharing a key left out, numbers as text, add to nothing), 2 (a shared key named with both
rows, a near miss with its key text, add's mismatch saying what differs, a stale add), 3 (a
key sent again with its value under add), 4 (a query for the rows left where a key repeats),
5 (two chunks with add).

Round 55 (R118): a key two or more rows with text hold refuses that value only, the rest of
the call recorded (``_parts_book``: 60 rows, three sharing a key, a 40-value call with the
shared key at values[28], replacing and with add); rules 2 (the refused value by place with
its rows and the way out), 4 (``refused`` capped and counted), 5 (one call, the exception
listed).
"""

from __future__ import annotations

import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox
from shelf_book import JUDGED, ROWS, SECTIONS, shelf_book

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


def test_a_key_two_rows_with_text_hold_refuses_that_value_only_naming_both_rows():
    """Round 54, R113, made per item in round 55 (R118): duplicates are judged among rows
    with source text only; a key named that two such rows hold names no one row, so that
    value alone is refused (``refused``, both rows named) and the others are recorded."""
    rows = [r[:] for r in ROWS]
    rows[9][0] = "S-01"
    wb = Workbook.open(shelf_book(rows).to_bytes())
    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                                values=JUDGED)
    assert recorded.refused == ({"code": "key_shared", "item": 0, "key": "S-01",
                                 "rows": ["t:Shelf/#1/Ref", "t:Shelf/#10/Ref"]},)
    assert recorded.recorded == 5 and recorded.shared == ("t:Shelf/#1", "t:Shelf/#10")
    assert [k for k, _ in recorded.not_recorded] == ["S-01", "S-09", "S-01"]
    # without the shared key: the same rows recorded, nothing refused
    again = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                             values=JUDGED[1:])
    assert again.dataset.rows == recorded.dataset.rows and again.refused == ()
    # a call whose every value names a shared key records nothing: refused, rows named
    with pytest.raises(RecordError) as caught:
        wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                         values=JUDGED[:1])
    assert caught.value.code == "key_shared" and caught.value.field == "values"
    assert caught.value.problems == list(recorded.refused)
    assert "S-01" not in caught.value.message
    # refused for another problem, the shared key is named with it (every problem at once)
    with pytest.raises(RecordError) as caught:
        wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                         values=[*JUDGED, {"key": "X-99", "value": "done"}])
    assert [(p["code"], p["item"]) for p in caught.value.problems] == [
        ("key_shared", 0), ("unknown_key", 6)]


def test_heading_and_spacer_rows_repeating_a_key_are_left_out():
    """R113: a section heading and a spacer row share an item's reference; only the item row
    has notes, so the key names it; the rows left out are counted in a note."""
    wb = Workbook.open(shelf_book(SECTIONS).to_bytes())
    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="grade",
                                values=[{"key": "S-04", "value": "4"},
                                        {"key": "S-05", "value": "3"},
                                        {"key": "S-08", "value": "3.5"}])
    data = recorded.dataset
    assert [row[:2] for row in data.rows] == [("S-04", "4"), ("S-05", "3"), ("S-08", "3.5")]
    # each value sourced by the item row's notes (the key column repeats: by address)
    assert [src[1] for src in data.sources] == [("Shelf!C3",), ("Shelf!C6",), ("Shelf!C12",)]
    assert recorded.ignored == ("t:Shelf/#1", "t:Shelf/#4", "t:Shelf/#6", "t:Shelf/#10")
    assert recorded.notes == ("4 row(s) without text in notes share a key with a row "
                              "recorded: left out (keys are matched among rows with text)",)
    assert recorded.not_recorded == (("S-06", "Shelf!C9"), ("S-07", "Shelf!C10"))
    assert recorded.repeated == {"S-06"} and recorded.shared == ()


# -- adding, numbers (round 54) ----------------------------------------------------------------


def test_add_to_extends_same_value_again_is_no_error_another_is_refused():
    wb = Workbook.open(shelf_book().to_bytes())
    first = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                             values=JUDGED[:3])
    more = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                            values=[*JUDGED[3:], {"key": "S-01", "value": "done"}],
                            add_to=first.dataset)
    assert more.recorded == 6 and more.added == 3 and more.kept == 2
    assert more.dataset.rows == wb.record_values(
        "t:Shelf", key="Ref", source="notes", column="outcome", values=JUDGED).dataset.rows
    with pytest.raises(RecordError) as caught:
        wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                         values=[{"key": "S-02", "value": "done"}], add_to=more.dataset)
    assert caught.value.problems == [{"code": "already_recorded", "item": 0, "key": "S-02",
                                      "row": "t:Shelf/#2/Ref"}]
    for wrong, said in ((dict(column="result"), "column 'outcome' (this call 'result')"),
                        (dict(type="number"), "type 'text' (this call 'number')")):
        arguments = {"key": "Ref", "source": "notes", "column": "outcome", **wrong}
        with pytest.raises(RecordError) as caught:
            wb.record_values("t:Shelf", values=[{"key": "S-09", "value": "1"}],
                             add_to=first.dataset, **arguments)
        assert caught.value.code == "add_mismatch" and said in caught.value.message


def test_type_number_takes_numbers_and_plain_number_text_only():
    wb = Workbook.open(shelf_book(SECTIONS).to_bytes())
    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="grade",
                                type="number", values=[
                                    {"key": "S-04", "value": 4}, {"key": "S-05", "value": "3"},
                                    {"key": "S-06", "value": 1.0},
                                    {"key": "S-08", "value": " 3.5 "}])
    assert [row[1] for row in recorded.dataset.rows] == [4, 3, 1.0, 3.5]
    assert recorded.dataset.types[1] == "number"
    with pytest.raises(RecordError) as caught:
        wb.record_values("t:Shelf", key="Ref", source="notes", column="grade", type="number",
                         values=[{"key": "S-04", "value": "four"},
                                 {"key": "S-05", "value": "3,5"},
                                 {"key": "S-06", "value": True},
                                 {"key": "S-07", "value": "1 000"},
                                 {"key": "S-08", "value": ""}])
    assert [p["code"] for p in caught.value.problems] == ["not_number"] * 4 + ["empty_value"]
    assert "four" not in caught.value.message


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
    assert again.warnings == ["d:outcome replaced: it holds this call's values only (add: "
                              "true adds to it)"]
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
        {"code": "unknown_key", "item": 0, "like": [{"row": "t:Shelf/#1/Ref", "key": "S-01"}]},
        {"code": "key_twice", "item": 2, "row": "t:Shelf/#2/Ref", "first": 1},
        {"code": "empty_source", "item": 3, "row": "t:Shelf/#6/notes"},
        {"code": "empty_source", "item": 4, "row": "t:Shelf/#7/notes"},
        {"code": "empty_value", "item": 5, "row": "t:Shelf/#9/Ref"},
        {"code": "unknown_key", "item": 6,
         "like": [{"row": "t:Shelf/#10/Ref", "key": "S-010"}]},
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
    assert refused.error.code == "refused" and refused.error.details["code"] == "key_shared"
    assert refused.error.details["problems"][0]["rows"] == ["t:Shelf/#1/Ref", "t:Shelf/#2/Ref"]


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


# -- round 54: repeated keys, add, numbers, near-miss keys (the tool) --------------------------


def test_a_workbook_repeating_keys_on_headings_records_and_reads_what_was_left(box):
    """R113 (rule 1, 4): before, the whole call was refused duplicate_keys; now each named key
    finds the one row with notes. The rows left to judge come with a call that reads them --
    a query, since a key shared with a heading names no one row."""
    session = _session(box, shelf_book(SECTIONS).to_bytes())
    result = ok(record(box, session, [{"key": "S-04", "value": "4"},
                                      {"key": "S-05", "value": "3"}], column="grade",
                       id="grade"))
    assert result.data["recorded"] == 2
    assert result.warnings == ["3 row(s) without text in notes share a key with a row "
                               "recorded: left out (keys are matched among rows with text)"]
    left = result.data["not_recorded"]
    assert left["keys"] == ["S-06", "S-07", "S-08"] and "read" not in left
    query = ok(box.dispatch(session, "xlsx_query", left["query"]))
    # (not_blank: the S-08 heading whose notes are white space only comes too)
    assert [[c["a"] for c in row] for row in query.data["rows"]] == [
        ["Shelf!A9", "Shelf!C9"], ["Shelf!A10", "Shelf!C10"], ["Shelf!A11", "Shelf!C11"],
        ["Shelf!A12", "Shelf!C12"]]
    traced = ok(box.dispatch(session, "xlsx_trace", {"doc": "d1", "rows": ["d:grade/S-05"]}))
    assert traced.data["traced"][0]["links"][0]["row"] == "t:Shelf/#5"
    assert not any(text in result.summary for text in CONTENT)


def test_add_in_two_chunks_and_a_conflict_named_by_key(box):
    """Add mode: the second chunk extends d:grade; a key sent again with its value is no error
    (rule 3); with another value it is refused naming the key; add to nothing says so."""
    session = _session(box, shelf_book(SECTIONS).to_bytes())
    first = ok(record(box, session, [{"key": "S-04", "value": 4}, {"key": "S-05", "value": 3}],
                      column="grade", id="grade", type="number", add=True))
    assert first.warnings[0] == "add: there is no d:grade yet; recorded as a new dataset"
    second = ok(record(box, session, [{"key": "S-06", "value": 1}, {"key": "S-07", "value": 5},
                                      {"key": "S-04", "value": 4}],
                       column="grade", id="grade", type="number", add=True))
    assert second.data["recorded"] == 4 and second.data["added"] == 2
    assert "replaced" not in second.data and "not_recorded" in second.data
    read = ok(box.dispatch(session, "xlsx_read", {"target": "d:grade"}))
    assert [[c["v"] for c in row][:2] for row in read.data["rows"]] == [
        ["S-04", 4], ["S-05", 3], ["S-06", 1], ["S-07", 5]]
    clash = record(box, session, [{"key": "S-04", "value": 2}], column="grade", id="grade",
                   type="number", add=True)
    assert clash.error.code == "refused"
    assert clash.error.details["problems"] == [{"code": "already_recorded", "item": 0,
                                                "key": "S-04", "row": "t:Shelf/#2/Ref"}]
    assert "S-04" not in clash.error.message
    other = record(box, session, [{"key": "S-08", "value": 3.5}], column="grade", id="grade",
                   add=True)            # type text this time
    assert other.error.code == "refused" and other.error.field == "add"
    assert other.error.details["code"] == "add_mismatch"
    ok(box.dispatch(session, "xlsx_derive", {"from": "d:grade", "id": "by_grade",
                                             "group_by": ["grade"]}))
    not_ours = record(box, session, [{"key": "S-08", "value": 3}], column="grade",
                      id="by_grade", type="number", add=True)
    assert not_ours.error.details["code"] == "add_mismatch"
    assert ok(box.dispatch(session, "xlsx_read", {"target": "d:grade"})).total == 4


def test_number_values_are_averaged_by_derive(box):
    session = _session(box, shelf_book(SECTIONS).to_bytes())
    ok(record(box, session, [{"key": "S-04", "value": 4}, {"key": "S-05", "value": "3"},
                             {"key": "S-06", "value": 1}, {"key": "S-07", "value": 5},
                             {"key": "S-08", "value": 3.5}],
              column="grade", id="grade", type="number"))
    mean = ok(box.dispatch(session, "xlsx_derive", {
        "from": "d:grade", "id": "mean", "measures": [{"fn": "mean", "column": "grade"}]}))
    cell = mean.data["rows"][0][0]
    assert cell["v"] == 3.3 and cell["t"] == "number"
    assert cell["sources"] == ["Shelf!C3", "Shelf!C6", "Shelf!C9:C10", "Shelf!C12"]
    wrong = record(box, session, [{"key": "S-04", "value": "four"}, {"key": "S-05", "value": True}],
                   column="grade", id="grade", type="number")
    assert wrong.error.code == "invalid_arguments"


def test_a_near_miss_gives_the_key_as_held(box):
    """The answer to round 53: a refused key's near misses carry the key text (key cells only)
    in the result the model reads; the message and summary keep to places."""
    session = _session(box, shelf_book(SECTIONS).to_bytes())
    refused = record(box, session, [{"key": "s 7", "value": "5"}, {"key": "S-004", "value": "4"}],
                     column="grade", id="grade")
    assert refused.error.details["problems"] == [
        {"code": "unknown_key", "item": 0, "like": [{"row": "t:Shelf/#9/Ref", "key": "S-07"}]},
        {"code": "unknown_key", "item": 1, "like": [{"row": "t:Shelf/#2/Ref", "key": "S-04"}]}]
    assert refused.error.valid_options == ["t:Shelf/#9/Ref", "t:Shelf/#2/Ref"]
    assert not any(text in refused.error.message for text in CONTENT)


def test_add_after_a_recorded_key_changed_in_the_workbook_is_refused_with_the_way_out(box):
    session = _session(box, shelf_book(SECTIONS).to_bytes())
    ok(record(box, session, [{"key": "S-04", "value": 4}], column="grade", id="grade",
              type="number"))
    ok(box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
        {"address": "Shelf!A3", "text": "S-44"}]}))
    stale = record(box, session, [{"key": "S-05", "value": 3}], column="grade", id="grade",
                   type="number", add=True)
    assert stale.error.code == "refused" and stale.error.details["code"] == "add_stale"
    assert "without add" in stale.error.message


# -- round 55, R118: a shared key refuses its value only ---------------------------------------


def _parts_book() -> bytes:
    """A made-up 'Parts' table of 60 rows, each with notes; rows #54, #55 and #56 share the
    key P-54 (the shape of the call refused in production: three rows with text, one key)."""
    rows: list[list[object]] = [["Ref", "Bin", "notes"]]
    for n in range(1, 61):
        rows.append(["P-54" if n in (54, 55, 56) else f"P-{n:02d}", f"bin {n}",
                     f"checked, item {n}"])
    out = WorkbookBuilder()
    sheet = out.sheet("Parts")
    sheet.rows("A1", rows)
    sheet.table("Parts", "A1:C61", ["Ref", "Bin", "notes"])
    return out.to_bytes()


#: 40 values, the 29th (values[28]) on the shared key.
PARTS = [{"key": k, "value": "done"} for k in
         [f"P-{n:02d}" for n in range(1, 29)] + ["P-54"] + [f"P-{n:02d}" for n in range(29, 40)]]
PARTS_CALL = {"table": "t:Parts", "key": "Ref", "source": "notes", "column": "outcome",
              "id": "outcome"}
SHARED = {"code": "key_shared", "item": 28, "key": "P-54",
          "rows": ["t:Parts/#54/Ref", "t:Parts/#55/Ref", "t:Parts/#56/Ref"]}


def test_a_shared_key_in_a_40_value_call_refuses_that_value_and_records_the_rest(box):
    """R118 (rules 2, 5): on 0.1.39 the whole call was refused ("nothing was recorded",
    key_shared, values[28]); now the other 39 values are recorded, values[28] is in
    ``refused`` as a refusal's problem would be, naming its three rows, and a warning says
    where it is and the way out -- by place, never by key."""
    assert len(PARTS) == 40 and PARTS[28]["key"] == "P-54"
    session = _session(box, _parts_book())
    result = ok(record(box, session, PARTS, **PARTS_CALL))
    assert result.data["recorded"] == 39 and result.data["refused"] == [SHARED]
    assert result.data["shared_keys"] == {"count": 3, "rows": ["t:Parts/#54", "t:Parts/#55",
                                                                "t:Parts/#56"]}
    assert result.warnings == [
        "1 value(s) not recorded, each key held by two or more rows with text to judge: "
        "values[28]; the rest are recorded. each refused item's rows are the rows holding the "
        "key with text in 'notes': give each its own key in the workbook, or leave the key out"]
    assert "1 value(s) refused" in result.summary
    assert not any("P-" in text for text in [result.summary, *result.warnings])
    read = ok(box.dispatch(session, "xlsx_read", {"target": "d:outcome"}))
    keys = [row[0]["v"] for row in read.data["rows"]]
    assert len(keys) == 39 and "P-54" not in keys
    # the shared rows are left to judge, and the query that reads them names the key once
    left = result.data["not_recorded"]
    assert left["keys"].count("P-54") == 3
    assert left["query"]["where"][0]["values"].count("P-54") == 1


def test_a_shared_key_with_add_refuses_that_value_and_adds_the_rest(box):
    """R118 with add: true: the second chunk holds the shared key (its values[8]); the rest
    of the chunk is added to the first, the shared key refused by its place in this call."""
    session = _session(box, _parts_book())
    first = ok(record(box, session, PARTS[:20], **PARTS_CALL))
    assert first.data["recorded"] == 20 and "refused" not in first.data
    second = ok(record(box, session, PARTS[20:], add=True, **PARTS_CALL))
    assert second.data["recorded"] == 39 and second.data["added"] == 19
    assert second.data["refused"] == [{**SHARED, "item": 8}]
    assert "values[8]" in second.warnings[0]
    # the shared key alone, added: nothing to record, so refused, its rows named
    alone = record(box, session, [PARTS[28]], add=True, **PARTS_CALL)
    assert alone.error.code == "refused" and alone.error.details["code"] == "key_shared"
    assert alone.error.details["problems"] == [{**SHARED, "item": 0}]
    assert "nothing was recorded" in alone.error.message
    assert ok(box.dispatch(session, "xlsx_read", {"target": "d:outcome"})).total == 39


def test_refused_values_are_capped_and_counted(box):
    """Rule 4: ``refused`` lists at most LISTED values, the rest counted."""
    rows: list[list[object]] = [["Ref", "notes"]]
    for n in range(LISTED + 3):
        rows += [[f"K{n}", "a"], [f"K{n}", "b"]]
    rows.append(["single", "c"])
    out = WorkbookBuilder()
    sheet = out.sheet("Pairs")
    sheet.rows("A1", rows)
    sheet.table("Pairs", f"A1:B{len(rows)}", ["Ref", "notes"])
    session = _session(box, out.to_bytes())
    values = [{"key": f"K{n}", "value": "x"} for n in range(LISTED + 3)]
    result = ok(record(box, session, [*values, {"key": "single", "value": "x"}],
                       table="t:Pairs", key="Ref", source="notes", column="outcome",
                       id="pairs"))
    assert result.data["recorded"] == 1
    assert len(result.data["refused"]) == LISTED and result.data["refused_more"] == 3
    assert "and 3 more" in result.warnings[-1]
