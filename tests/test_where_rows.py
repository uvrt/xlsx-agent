"""Round 51, R104: a derive with ``where`` alone -- no ``group_by``, ``measures``, ``bins``,
``pivot``, ``join`` or ``fill`` -- keeps the matching rows (a filter's meaning: design rule
1), every column, in the table's order, each value's sources its own cell; the count stays
``measures=[{"fn": "count"}]`` (rule 2: a refusal past the row limit names it).  Every other
form is as it was.  Made-up workbooks only (tests/build.py)."""

from __future__ import annotations

import pytest

from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from test_datasets import _session, call, err
from test_keys import book as checks_book, CHECKS
from test_truth_filters import ROWS, items

from xlsx_agent import Workbook
from xlsx_agent.derive import MAX_ROWS, DeriveError
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

KEPT = [{"column": "Kept", "op": "is_true"}]


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return Workbook.open(items().to_bytes())


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


# -- rule 1: where alone is a filter: the rows it keeps -----------------------------------------


def test_where_alone_keeps_the_matching_rows_every_column_in_order(wb):
    ds = wb.derive("t:Items", where=KEPT)
    assert ds.columns == tuple(ROWS[0])
    assert ds.rows == [{"Code": "K7-12.3", "Kept": True, "Qty": 3},
                       {"Code": "K7-13.1", "Kept": True, "Qty": 5}]
    assert ds.types == {"Code": "text", "Kept": "bool", "Qty": "number"}
    assert ds.matched == ds.used == 2
    # the text cell in the TRUE/FALSE column is counted, as for any filter
    assert ds.not_compared["Kept"].count == 1
    # each value's sources are its own cell
    assert [str(s) for s in ds.provenance[1]["Qty"].sources] == ["Items!C4"]
    assert ds.joined is None and not ds.filled


def test_sort_and_top_apply_to_the_kept_rows(wb):
    ds = wb.derive("t:Items", where=[{"column": "Qty", "op": "gte", "value": "2"}],
                   sort=[("Qty", "desc")], top=2)
    assert [row["Code"] for row in ds.rows] == ["K7-14.1", "K7-13.1"]
    assert ds.matched == 5


def test_a_keyed_tables_rows_name_their_sources_by_key():
    wb = Workbook.open(checks_book(CHECKS))
    ds = wb.derive("t:checks", where=[{"column": "area", "op": "eq", "value": "east"}])
    assert [row["code"] for row in ds.rows] == ["C1", "C2"]
    assert ds.key_columns == 1
    assert [str(s) for s in ds.provenance[0]["rating"].sources] == ["t:checks/C1/rating"]


def test_a_dataset_filtered_alone_keeps_its_rows_and_their_original_sources(wb):
    every = wb.derive("t:Items", where=[{"column": "Qty", "op": "gte", "value": "1"}])
    assert len(every.rows) == len(ROWS) - 1
    ds = wb.derive(every, where=KEPT)
    assert [row["Code"] for row in ds.rows] == ["K7-12.3", "K7-13.1"]
    assert [str(s) for s in ds.provenance[0]["Code"].sources] == ["Items!A2"]


def test_no_matching_row_is_no_rows(wb):
    ds = wb.derive("t:Items", where=[{"column": "Code", "op": "eq", "value": "none"}])
    assert ds.rows == [] and ds.matched == 0
    assert ds.columns == tuple(ROWS[0])


# -- the count: one measure away; every other form unchanged ------------------------------------


def test_the_count_is_the_count_measure(wb):
    assert wb.derive("t:Items", where=KEPT, measures=[{"fn": "count"}]).rows == [{"count": 2}]
    assert wb.derive("t:Items", where=KEPT,
                     measures=[{"fn": "count", "as": "n"}]).rows == [{"n": 2}]


def test_without_where_a_derive_still_counts_every_row(wb):
    assert wb.derive("t:Items").rows == [{"count": len(ROWS) - 1}]
    # an empty where (a placeholder) is no filter: the count, as before
    assert wb.derive("t:Items", where=[]).rows == [{"count": len(ROWS) - 1}]


def test_where_with_groups_measures_bins_or_a_pivot_is_as_before(wb):
    grouped = wb.derive("t:Items", where=KEPT, group_by=["Kept"])
    assert grouped.rows == [{"Kept": True, "count": 2}]
    summed = wb.derive("t:Items", where=KEPT, measures=[{"fn": "sum", "column": "Qty"}])
    assert summed.rows == [{"sum Qty": 8}]
    binned = wb.derive("t:Items", where=KEPT,
                       bins={"column": "Qty", "edges": ["0", "4", "10"], "as": "band"})
    assert [row["count"] for row in binned.rows] == [1, 1]
    pivoted = wb.derive("t:Items", where=[{"column": "Qty", "op": "gte", "value": "1"}],
                        group_by=["Kept"], pivot="Code")
    assert len(pivoted.rows) == 4      # TRUE, FALSE, blank and the text cell's group


def test_where_with_a_join_or_a_fill_still_gives_their_rows():
    built = WorkbookBuilder()
    sheet = built.sheet("Items")
    sheet.rows("A1", ROWS)
    sheet.table("Items", f"A1:C{len(ROWS)}", ROWS[0], write_header=False)
    owners = [["Code", "Owner"], ["K7-12.3", "lab"], ["K7-13.1", "yard"]]
    other = built.sheet("Owners")
    other.rows("A1", owners)
    other.table("Owners", "A1:B3", owners[0], write_header=False)
    wb = Workbook.open(built.to_bytes())
    joined = wb.derive("t:Items", where=KEPT,
                       join={"from": "t:Owners", "on": ["Code", "Code"]})
    assert [row["Owner"] for row in joined.rows] == ["lab", "yard"]
    assert joined.joined is not None
    filled = wb.derive("t:Items", where=KEPT,
                       fill=[{"column": "Qty", "from": "Code"}])
    assert [row["Qty"] for row in filled.rows] == [3, 5]
    assert filled.filled


# -- rule 2: a refusal carries the way out -----------------------------------------------------


def test_round_is_refused_naming_measures(wb):
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Items", where=KEPT, round=2)
    assert caught.value.code == "bad_round"
    assert "where without group_by or measures gives its rows' cells as stored" \
        in str(caught.value)


def test_more_rows_than_a_result_holds_is_refused_naming_top_and_the_count():
    built = WorkbookBuilder()
    sheet = built.sheet("Log")
    rows = [["Ref", "Qty"]] + [[f"R{n}", n] for n in range(1, MAX_ROWS + 3)]
    sheet.rows("A1", rows)
    sheet.table("Log", f"A1:B{len(rows)}", rows[0], write_header=False)
    wb = Workbook.open(built.to_bytes())
    where = [{"column": "Qty", "op": "gte", "value": "1"}]
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Log", where=where)
    assert caught.value.code == "too_many_rows"
    assert "give top" in str(caught.value)
    assert "count them: measures [{fn: count}]" in str(caught.value)
    assert len(wb.derive("t:Log", where=where, top=MAX_ROWS).rows) == MAX_ROWS
    assert wb.derive("t:Log", where=where,
                     measures=[{"fn": "count"}]).rows == [{"count": MAX_ROWS + 2}]


# -- the tool ----------------------------------------------------------------------------------


def test_the_tool_keeps_the_rows_as_a_dataset_read_again_with_their_sources(box):
    session = _session(box, items())
    result = call(box, session, "xlsx_derive", doc="d1", id="kept", where=KEPT,
                  **{"from": "t:Items"})
    assert result.ok, result.error
    assert result.summary == "Derived d:kept from t:Items: 2 rows"
    assert result.data["columns"] == ROWS[0]
    assert [[v["v"] for v in row] for row in result.data["rows"]] == \
        [["K7-12.3", True, 3], ["K7-13.1", True, 5]]
    assert result.data["rows"][1][2] == {"a": "d:kept/#2/Qty", "v": 5, "t": "number",
                                         "sources": ["Items!C4"]}
    read = call(box, session, "xlsx_read", target="d:kept/#1/Code")
    assert read.ok, read.error
    counted = call(box, session, "xlsx_derive", doc="d1", id="kept_n", where=KEPT,
                   measures=[{"fn": "count"}], **{"from": "t:Items"})
    assert counted.ok and counted.data["rows"][0][0]["v"] == 2


def test_the_tool_says_how_to_count(box):
    definition = next(d for d in box.definitions("anthropic", groups="core")
                      if d.get("name") == "xlsx_derive")
    assert "where/join/fill without group_by, measures, bins, pivot: the rows" \
        in definition["description"]
    assert "a count: measures [{fn: count}]" in definition["description"]


def test_the_tool_refuses_round_with_the_way_out(box):
    session = _session(box, items())
    result = call(box, session, "xlsx_derive", doc="d1", id="kept", where=KEPT, round=1,
                  **{"from": "t:Items"})
    assert not result.ok
    assert "round rounds measures" in err(result)["message"]
