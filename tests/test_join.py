"""Round 25, X1: ``join`` in a derive -- another table's or dataset's columns added to each row
of a table, by key, each added value with the cells it came from (:mod:`xlsx_agent.join`).

Every workbook is made up here: a table of items, and a table of issues whose "Items" cell
lists the items each issue concerns (a declared split), plus a table of owners with one
row per item.
"""

from __future__ import annotations

import json

import pytest
from build import WorkbookBuilder
from ooxml_edit.tools import Dataset, Toolbox

from xlsx_agent import DeriveError, NameNotFound, Workbook
from xlsx_agent.derive import MAX_ROWS
from xlsx_agent.join import JoinedTable
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

ISSUES = {"name": "issues", "sheets": "Issues", "key": "Issue",
          "fields": [{"header": "Items", "clean": {"split": ";"}}]}


def book() -> WorkbookBuilder:
    out = WorkbookBuilder()
    items = out.sheet("Items")
    items.rows("A1", [["Ref", "Name", "Status", "Weight"], ["I-1", "pump", "ok", 4],
                      ["I-2", "valve", "late", 2], ["I-3", "hose", "ok", 1],
                      ["I-4", "seal", "ok", 1]])
    items.table("Items", "A1:D5", ["Ref", "Name", "Status", "Weight"])
    issues = out.sheet("Issues")
    issues.rows("A1", [["Issue", "Items", "Status"], ["Q-1", "I-1; I-3", "open"],
                       ["Q-2", "I-1", "closed"], ["Q-3", "I-9", "open"],
                       ["Q-4", "i-2", "open"]])
    owners = out.sheet("Owners")
    owners.rows("A1", [["Item", "Owner"], ["I-3", "Kim"], ["I-1", "Ash"], ["I-2", None]])
    owners.table("Owners", "A1:B4", ["Item", "Owner"])
    return out


@pytest.fixture(scope="module")
def wb() -> Workbook:
    opened = Workbook.open(book().to_bytes())
    opened.declare_tables([ISSUES])
    return opened


JOIN = {"from": "t:issues", "on": ["Ref", "Items"]}


def sources(derived, row: int, column: str) -> list[str]:
    return [str(s) for s in derived.provenance[row][column].sources]


# -- the rows ----------------------------------------------------------------------------------


def test_every_row_kept_several_matches_a_list_cell_unmatched_blank(wb):
    joined = wb.derive("t:Items", join=JOIN, id="items_issues")
    assert joined.columns == ("Ref", "Name", "Status", "Weight", "Issue", "Status (issues)")
    assert joined.types["Issue"] == "list" and joined.types["Status (issues)"] == "list"
    assert [(r["Ref"], r["Issue"], r["Status (issues)"]) for r in joined.rows] == [
        ("I-1", ("Q-1", "Q-2"), ("open", "closed")),
        ("I-2", None, None),                    # "i-2" is another key: case counts
        ("I-3", ("Q-1",), ("open",)),           # one match in a list column: a list of one
        ("I-4", None, None)]
    assert joined.joined.as_dict() == {
        "from": "t:issues", "on": ["Ref", "Items"],
        "columns": {"Issue": "Issue", "Status (issues)": "Status"},
        "matched": 2, "unmatched": 2, "lists": True}
    assert joined.matched == joined.used == joined.groups == 4


def test_each_joined_value_lists_the_cells_it_came_from(wb):
    joined = wb.derive("t:Items", join=JOIN)
    assert sources(joined, 0, "Issue") == ["t:issues/{Q-1,Q-2}/Issue"]
    assert sources(joined, 0, "Status (issues)") == ["t:issues/{Q-1,Q-2}/Status"]
    assert sources(joined, 1, "Issue") == []
    assert sources(joined, 0, "Name") == ["Items!B2"]
    assert joined.provenance[1]["Issue"].count == 0
    data = joined.dataset()
    assert data.sources[0][4] == ("t:issues/{Q-1,Q-2}/Issue",)
    # placed in a deck or document a list is its items, one per line
    assert data.rows[0][4] == "Q-1\nQ-2" and data.types[4] == "text"


def test_one_match_at_most_keeps_the_values_as_they_are(wb):
    joined = wb.derive("t:Items", join={"from": "t:Owners", "on": ["Ref", "Item"]})
    assert joined.columns[-1] == "Owner" and joined.types["Owner"] == "text"
    assert [r["Owner"] for r in joined.rows] == ["Ash", None, "Kim", None]
    assert sources(joined, 0, "Owner") == ["t:Owners/I-1/Owner"]      # by its key
    assert joined.joined.lists is False and joined.joined.matched == 3


def test_a_list_key_on_the_left_matches_by_any_item(wb):
    joined = wb.derive("t:issues", join={"from": "t:Items", "on": ["Items", "Ref"],
                                         "columns": ["Name"]})
    assert [r["Name"] for r in joined.rows] == [("pump", "hose"), ("pump",), None, None]


def test_columns_picks_and_orders_and_renames_on_a_clash(wb):
    joined = wb.derive("t:Items", join={**JOIN, "columns": ["status", "Issue"]})
    assert joined.columns[4:] == ("Status (issues)", "Issue")
    assert joined.joined.columns == {"Status (issues)": "Status", "Issue": "Issue"}


def test_keys_match_by_type_and_value():
    data = Dataset.from_records([{"k": 1, "v": "one"}, {"k": "2", "v": "two"},
                                 {"k": 3.0, "v": "three"}])
    left = Dataset.from_records([{"k": 1.0}, {"k": 2}, {"k": 3}, {"k": None}])
    joined = Workbook.open(book().to_bytes()).derive(left, join={"from": data,
                                                                 "on": ["k", "k"]})
    assert [r["v"] for r in joined.rows] == ["one", None, "three", None]


# -- the pipeline ------------------------------------------------------------------------------


def test_the_join_comes_first_then_where_groups_and_measures(wb):
    grouped = wb.derive("t:Items", join={**JOIN, "columns": ["Issue"]},
                        where=[{"column": "Status", "op": "eq", "value": "ok"}],
                        group_by=["Status"],
                        measures=[{"fn": "count", "column": "Issue", "as": "with issues"},
                                  {"fn": "count", "as": "items"},
                                  {"fn": "sum", "column": "Weight", "as": "weight"}])
    assert grouped.rows == [{"Status": "ok", "with issues": 2, "items": 3, "weight": 6.0}]
    assert sources(grouped, 0, "with issues") == ["t:issues/{Q-1,Q-2}/Issue"]
    assert sources(grouped, 0, "weight") == ["Items!D2", "Items!D4:D5"]
    assert grouped.joined.columns == {"Issue": "Issue"}
    filtered = wb.derive("t:Items", join={**JOIN, "columns": ["Issue"]},
                         where=[{"column": "Issue", "op": "eq", "value": "Q-2"}])
    assert [r["Ref"] for r in filtered.rows] == ["I-1"]


def test_a_joined_list_column_is_not_grouped_by(wb):
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Items", join=JOIN, group_by=["Issue"])
    assert caught.value.code == "wrong_type"


def test_the_rows_sort_and_top_and_round_is_refused(wb):
    joined = wb.derive("t:Items", join={"from": "t:Owners", "on": ["Ref", "Item"]},
                       sort=[{"column": "Owner"}], top=3)
    assert [r["Ref"] for r in joined.rows] == ["I-1", "I-3", "I-2"]   # blanks last
    assert joined.groups == 4
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Items", join=JOIN, round=1)
    assert caught.value.code == "bad_round"


def test_too_many_rows_are_refused_with_the_way_out():
    big = WorkbookBuilder()
    sheet = big.sheet("Log")
    sheet.rows("A1", [["Ref", "N"]] + [[f"R-{n}", n] for n in range(MAX_ROWS + 1)])
    sheet.table("Log", f"A1:B{MAX_ROWS + 2}", ["Ref", "N"])
    wb = Workbook.open(big.to_bytes())
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Log", join={"from": "t:Log", "on": ["Ref", "Ref"], "columns": ["N"]})
    assert caught.value.code == "too_many_rows" and "top" in str(caught.value)
    kept = wb.derive("t:Log", join={"from": "t:Log", "on": ["Ref", "Ref"], "columns": ["N"]},
                     top=2)
    assert [r["N (Log)"] for r in kept.rows] == [0, 1]


@pytest.mark.parametrize("join, code", [
    ({"from": "t:issues"}, "bad_join"),
    ({"from": "t:issues", "on": ["Ref"]}, "bad_join"),
    ({"from": "t:issues", "on": "Ref"}, "bad_join"),
    ({"from": "t:issues", "on": ["Ref", "Items"], "columns": []}, "bad_join"),
    ({"from": "t:issues", "on": ["Ref", "Items"], "how": "left"}, "bad_join"),
    ({"from": "t:issues[Items]", "on": ["Ref", "Items"]}, "bad_join"),
    ({"from": "t:Owners", "on": ["Ref", "Item"], "columns": ["Item"]}, None),
])
def test_a_join_that_is_not_one_is_refused(wb, join, code):
    if code is None:   # the key alone: allowed when named
        assert wb.derive("t:Items", join=join).columns[-1] == "Item"
        return
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Items", join=join)
    assert caught.value.code == code and caught.value.field == "join"


def test_an_unknown_column_names_the_ones_there_are(wb):
    with pytest.raises(NameNotFound):
        wb.derive("t:Items", join={"from": "t:issues", "on": ["Ref", "Item"]})


def test_a_dataset_on_either_side(wb):
    by_status = wb.derive("t:Items", group_by=["Status"],
                          measures=[{"fn": "count", "as": "n"}], id="by_status")
    joined = wb.derive("t:Items", join={"from": by_status, "on": ["Status", "Status"]})
    assert [r["n"] for r in joined.rows] == [3, 1, 3, 3]
    assert sources(joined, 0, "n") == ["Items!C2", "Items!C4:C5"]   # the original cells
    back = wb.derive(by_status, join={"from": "t:Items", "on": ["Status", "Status"],
                                      "columns": ["Ref"]})
    assert [r["Ref"] for r in back.rows] == [("I-1", "I-3", "I-4"), ("I-2",)]
    again = wb.derive(joined, group_by=["Status"], measures=[{"fn": "max", "column": "n"}])
    assert again.rows == [{"Status": "ok", "max n": 3}, {"Status": "late", "max n": 1}]
    assert isinstance(JoinedTable, type)


# -- the tool ----------------------------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[xlsx_format(tables=[ISSUES])], groups=GROUPS,
                 workers=0) as toolbox:
        yield toolbox


def _session(box):
    session = box.session()
    blob = session.add_blob(book().to_bytes(), name="book.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_derive_joins_a_table_and_keeps_the_rows_as_a_dataset(box):
    session = _session(box)
    result = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Items", "id": "items_issues",
        "join": {"from": "t:issues", "on": ["Ref", "Items"], "columns": ["Issue"]}})
    assert result.ok, result.error
    assert result.data["joined"] == {"from": "t:issues", "on": ["Ref", "Items"],
                                     "columns": {"Issue": "Issue"}, "matched": 2,
                                     "unmatched": 2, "lists": True}
    first = json.loads(json.dumps(result.data))["rows"][0][4]
    assert first == {"a": "d:items_issues/#1/Issue", "v": ["Q-1", "Q-2"], "t": "list",
                     "sources": ["t:issues/{Q-1,Q-2}/Issue"]}

    read = box.dispatch(session, "xlsx_read", {"target": "d:items_issues/#3/Issue"})
    assert read.ok and json.loads(json.dumps(read.data))["rows"] == [[{"a": "d:items_issues/#3/Issue",
                                               "v": ["Q-1"], "t": "list"}]]


def test_xlsx_derive_joins_datasets(box):
    session = _session(box)
    assert box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Owners", "id": "owners", "group_by": ["Item", "Owner"]}).ok
    result = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Items", "id": "with_owner",
        "join": {"from": "d:owners", "on": ["Ref", "Item"], "unused": 1}})
    assert not result.ok and result.error.field == "join.unused"    # an unknown key
    result = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Items", "id": "with_owner",
        "join": {"from": "d:owners", "on": ["Ref", "Item"], "columns": ["Owner"]}})
    assert result.ok, result.error
    assert [row[-1]["v"] for row in result.data["rows"]] == ["Ash", None, "Kim", None]
    assert result.data["rows"][0][-1]["sources"] == ["t:Owners/I-1/Owner"]
    # a dataset joined with a table needs the table's workbook
    refused = box.dispatch(session, "xlsx_derive", {
        "from": "d:owners", "id": "x", "join": {"from": "t:Items", "on": ["Item", "Ref"]}})
    assert not refused.ok and refused.error.field == "doc"
    taken = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "d:owners", "id": "x",
        "join": {"from": "t:Items", "on": ["Item", "Ref"], "columns": ["Name"]}})
    assert taken.ok, taken.error
    assert [row[-1]["v"] for row in taken.data["rows"]] == ["hose", "pump", "valve"]


def two_books(box):
    """``d1`` the items, ``d2`` a second copy whose issues and owners are the joined side."""
    session = box.session()
    for name in ("items.xlsx", "issues.xlsx"):
        blob = session.add_blob(book().to_bytes(), name=name)
        assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def _set(box, session, doc, address, text):
    written = box.dispatch(session, "xlsx_set_cells", {"doc": doc, "cells": [
        {"address": address, "text": text}]})
    assert written.ok, written.error


def test_a_joined_dataset_goes_stale_when_the_joined_side_changes(box):
    session = two_books(box)
    assert box.dispatch(session, "xlsx_derive", {
        "doc": "d2", "from": "t:Owners", "id": "owners", "group_by": ["Item", "Owner"]}).ok
    joined = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Items", "id": "with_owner",
        "join": {"from": "d:owners", "on": ["Ref", "Item"], "columns": ["Owner"]}})
    assert joined.ok, joined.error
    entry = session.dataset("d:with_owner")
    assert not entry.stale and entry.doc == "d1"
    _set(box, session, "d2", "Owners!B2", "Lee")               # the joined side only
    assert entry.stale
    read = box.dispatch(session, "xlsx_read", {"target": "d:with_owner"})
    assert read.data["dataset"]["stale"] is True
    assert any("a source it also read" in w for w in read.warnings)
    assert box.dispatch(session, "undo", {"doc": "d2"}).ok
    assert not entry.stale


def test_a_dataset_joined_with_a_table_goes_stale_when_that_workbook_changes(box):
    session = two_books(box)
    assert box.dispatch(session, "xlsx_derive", {
        "doc": "d2", "from": "t:Owners", "id": "owners", "group_by": ["Item"]}).ok
    joined = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "d:owners", "id": "named",
        "join": {"from": "t:Items", "on": ["Item", "Ref"], "columns": ["Name"]}})
    assert joined.ok, joined.error
    entry = session.dataset("d:named")
    assert not entry.stale and entry.doc == "d2"               # from's source
    _set(box, session, "d1", "Items!B2", "pumps")              # the joined table's workbook
    assert entry.stale
