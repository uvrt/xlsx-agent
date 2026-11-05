"""Round 35, R22: ``fill`` in a derive -- a column's blank cells given another column's value
of the same row, each filled value with the cells it came from and marked as filled
(:mod:`xlsx_agent.fill`).

Every workbook is made up here: a table of parts whose ``Limit`` is often blank and whose
``Default limit`` holds the value to use then (a number, text that reads as one, a word,
an error, a formula without a cached value, nothing), a ``Team limit`` tried first, and a
table of groups whose limit a part's group looks up.
"""

from __future__ import annotations

import json
from datetime import date

import pytest
from build import Error, Formula, WorkbookBuilder
from ooxml_edit.tools import Dataset, Toolbox

from xlsx_agent import DeriveError, Workbook
from xlsx_agent.dataset import DatasetTable
from xlsx_agent.derive import derive
from xlsx_agent.fill import FilledTable, fills
from xlsx_agent.tools import GROUPS, TOOLS, format as xlsx_format

COLUMNS = ["Code", "Name", "Limit", "Default limit", "Team limit", "Group", "Due", "Due text"]


def book() -> WorkbookBuilder:
    out = WorkbookBuilder()
    parts = out.sheet("Parts")
    parts.rows("A1", [COLUMNS,
                      ["P-1", "bolt", 10, 40, None, "g1", date(2026, 3, 1), None],
                      ["P-2", "nut", None, 40, 15, "g1", None, "2026-04-02"],
                      ["P-3", "gear", None, "25", None, "g2", None, "soon"],
                      ["P-4", "cog", None, None, None, "g2", None, None],
                      ["P-5", "pin", None, "many", None, "g2", None, None],
                      ["P-6", "axle", None, Error("#N/A"), None, "g3", None, None],
                      ["P-7", "bolt", None, Formula("D2*2"), None, "g3", None, None]])
    parts.table("Parts", "A1:H8", COLUMNS)
    groups = out.sheet("Groups")
    groups.rows("A1", [["Group", "Group limit"], ["g1", 30], ["g2", 20]])
    groups.table("Groups", "A1:B3", ["Group", "Group limit"])
    return out


@pytest.fixture(scope="module")
def wb() -> Workbook:
    opened = Workbook.open(book().to_bytes())
    opened.declare_key("t:Parts", "Code")
    return opened


FILL = [{"column": "Limit", "from": "Default limit", "coerce": "number"}]


def column(derived, name: str) -> list:
    return [row[name] for row in derived.rows]


# -- the rows ----------------------------------------------------------------------------------


def test_blanks_take_the_same_rows_value_and_nothing_else_changes(wb):
    filled = wb.derive("t:Parts", fill=FILL, id="limits")
    assert filled.columns == tuple(COLUMNS)
    # P-1 keeps its own 10; "25" is read as a number; the rest stay blank
    assert column(filled, "Limit") == [10, 40, 25.0, None, None, None, None]
    assert column(filled, "Default limit")[:3] == [40, 40, "25"]     # the source unchanged
    assert filled.matched == filled.used == filled.groups == 7
    assert filled.types["Limit"] == "number"


def test_each_filled_value_keeps_its_source_and_is_marked(wb):
    filled = wb.derive("t:Parts", fill=FILL)
    where = [p["Limit"] for p in filled.provenance]
    assert where[0].filled is None and [str(s) for s in where[0].sources] == [
        "t:Parts/P-1/Limit"]
    assert where[1].filled == "Default limit"
    assert [str(s) for s in where[1].sources] == ["t:Parts/P-2/Default limit"]
    assert where[2].as_dict() == {"sources": ["t:Parts/P-3/Default limit"], "count": 1,
                                  "filled": "Default limit"}
    assert where[3].filled is None and where[3].count == 0    # left blank: not marked


def test_the_counts_say_what_was_filled_and_why_the_rest_was_not(wb):
    filled = wb.derive("t:Parts", fill=FILL)
    assert [f.as_dict() for f in filled.filled] == [{
        "column": "Limit", "from": "Default limit", "coerce": "number",
        "filled": {"count": 2, "first": ["Parts!C3", "Parts!C4"]},
        "coerced": {"count": 1, "first": ["Parts!C4"]},
        "unfilled": {"blank": {"count": 1, "first": ["Parts!C5"]},
                     "error": {"count": 1, "first": ["Parts!C7"]},
                     "no_cached_value": {"count": 1, "first": ["Parts!C8"]},
                     "wrong_type": {"count": 1, "first": ["Parts!C6"]}}}]
    assert filled.as_dict()["filled"][0]["filled"]["count"] == 2


def test_without_coerce_a_value_is_filled_as_it_is(wb):
    filled = wb.derive("t:Parts", fill=[{"column": "Limit", "from": "Default limit"}])
    assert column(filled, "Limit")[:5] == [10, 40, "25", None, "many"]
    report = filled.filled[0]
    assert report.coerce is None and report.filled.count == 3 and not report.coerced.count
    assert set(report.unfilled) == {"blank", "error", "no_cached_value"}


def test_coerce_date_reads_iso_text_and_refuses_other_text(wb):
    filled = wb.derive("t:Parts", fill=[{"column": "Due", "from": "Due text",
                                         "coerce": "date"}])
    assert column(filled, "Due")[:3] == ["2026-03-01", "2026-04-02", None]
    assert filled.types["Due"] == "date"
    report = filled.filled[0].as_dict()
    assert report["filled"]["first"] == ["Parts!G3"]
    assert report["unfilled"]["wrong_type"]["first"] == ["Parts!G4"]


def test_several_fills_apply_in_order(wb):
    filled = wb.derive("t:Parts", fill=[
        {"column": "Limit", "from": "Team limit"},
        {"column": "Limit", "from": "Default limit", "coerce": "number"}])
    assert column(filled, "Limit")[:3] == [10, 15, 25.0]      # P-2 took Team limit first
    assert [f.filled.count for f in filled.filled] == [1, 1]
    assert filled.provenance[1]["Limit"].filled == "Team limit"
    assert filled.provenance[2]["Limit"].filled == "Default limit"


def test_a_value_filled_into_a_column_used_as_from_keeps_its_origin(wb):
    filled = wb.derive("t:Parts", fill=[
        {"column": "Team limit", "from": "Default limit", "coerce": "number"},
        {"column": "Limit", "from": "Team limit"}])
    assert filled.provenance[2]["Limit"].filled == "Default limit"
    assert [str(s) for s in filled.provenance[2]["Limit"].sources] == [
        "t:Parts/P-3/Default limit"]


def test_from_a_looked_up_table_join_then_fill(wb):
    filled = wb.derive("t:Parts", join={"from": "t:Groups", "on": ["Group", "Group"]},
                       fill=[{"column": "Limit", "from": "Group limit"}])
    assert column(filled, "Limit") == [10, 30, 20, 20, 20, None, None]
    assert [str(s) for s in filled.provenance[1]["Limit"].sources] == ["Groups!B2"]
    assert filled.provenance[1]["Limit"].filled == "Group limit"
    assert filled.joined is not None and filled.joined.matched == 5
    assert filled.filled[0].unfilled["blank"].count == 2       # g3 is in no row of Groups


def test_where_and_measures_read_the_filled_rows(wb):
    grouped = wb.derive("t:Parts", fill=FILL, where={"Limit": {"gte": 10}},
                        group_by=["Group"],
                        measures=[{"fn": "sum", "column": "Limit", "as": "total"}])
    assert grouped.rows == [{"Group": "g1", "total": 50.0}, {"Group": "g2", "total": 25.0}]
    assert [str(s) for s in grouped.provenance[0]["total"].sources] == [
        "t:Parts/P-1/Limit", "t:Parts/P-2/Default limit"]
    assert grouped.filled[0].filled.count == 2              # over every row, before where


def test_a_fill_never_overwrites_and_a_column_without_blanks_is_no_error(wb):
    filled = wb.derive("t:Parts", fill=[{"column": "Name", "from": "Code"}])
    assert column(filled, "Name")[:2] == ["bolt", "nut"]
    assert filled.filled[0].filled.count == 0 and not filled.filled[0].unfilled


def test_a_fill_from_itself_does_nothing_and_says_so(wb):
    filled = wb.derive("t:Parts", fill=[{"column": "Limit", "from": "limit"}])
    assert column(filled, "Limit")[1] is None
    assert filled.filled[0].note == ("Limit is filled from itself (names match case "
                                     "aside): nothing to fill")


def test_empty_text_is_blank_and_a_dataset_is_filled_too():
    data = Dataset(("Item", "Size", "Fallback"),
                   (("a", "", "1,250.5"), ("b", 3, 4), ("c", None, "x")),
                   sources=((("S!A2",), ("S!B2",), ("S!C2",)),
                            (("S!A3",), ("S!B3",), ("S!C3",)),
                            (("S!A4",), ("S!B4",), ("S!C4",))))
    filled = derive(None, data, fill={"column": "Size", "from": "Fallback",
                                      "coerce": "number_as_text"})
    assert column(filled, "Size") == [1250.5, 3, None]
    assert [str(s) for s in filled.provenance[0]["Size"].sources] == ["S!C2"]
    assert filled.filled[0].coerce == "number"           # the derive's name taken
    assert filled.filled[0].unfilled["wrong_type"].first == ("d:derived/#3/Size",)


def test_round_is_refused_on_rows_and_top_and_sort_apply(wb):
    with pytest.raises(DeriveError) as refused:
        wb.derive("t:Parts", fill=FILL, round=1)
    assert refused.value.code == "bad_round" and "a fill" in refused.value.message
    top = wb.derive("t:Parts", fill=FILL, sort=[("Limit", "desc")], top=2)
    assert column(top, "Code") == ["P-2", "P-3"]
    assert top.provenance[0]["Limit"].filled == "Default limit"


# -- the form ----------------------------------------------------------------------------------


@pytest.mark.parametrize("fill, says", [
    ([{"column": "Limit"}], "needs both column"),
    ([{"column": "Limit", "from": "Default limit", "coerce": "text"}], "number or date"),
    ([{"column": "Limit", "from": "Default limit", "with": 1}], "{column, from, coerce}"),
    ("Limit", "fill is a list"),
])
def test_a_fill_of_the_wrong_form_is_refused_with_the_form_that_works(wb, fill, says):
    with pytest.raises(DeriveError) as refused:
        wb.derive("t:Parts", fill=fill)
    assert refused.value.code == "bad_fill" and refused.value.field == "fill"
    assert says in refused.value.message


def test_an_unknown_column_is_refused_with_the_columns(wb):
    with pytest.raises(DeriveError) as refused:
        wb.derive("t:Parts", fill=[{"column": "Limt", "from": "Default limit"}])
    assert refused.value.code == "bad_fill"
    assert "Limit" in refused.value.valid_options
    assert "Default limit" in refused.value.message      # the columns are listed
    with pytest.raises(DeriveError) as refused:
        wb.derive("t:Parts", fill=[{"column": "Limit", "from": "Group limit"}])
    assert "after a join" in refused.value.message


def test_a_fill_is_bounded_as_a_join_is(wb, monkeypatch):
    import xlsx_agent.fill as fill_module

    monkeypatch.setattr(fill_module, "MAX_JOINED", 20)
    with pytest.raises(DeriveError) as refused:
        wb.derive("t:Parts", fill=FILL)
    assert refused.value.code == "too_many_cells" and refused.value.field == "fill"


def test_placeholders_a_strict_call_sends_are_dropped():
    assert fills([]) == [] and fills(None) == []
    assert fills([{"column": "", "from": "", "coerce": ""}]) == []
    assert fills([{"column": "A", "from": "B", "coerce": None}]) == [
        {"column": "A", "from": "B", "coerce": None}]
    assert fills({"column": "A", "from": "B", "coerce": "date_as_text"})[0]["coerce"] == "date"


def test_an_empty_fill_is_a_plain_derive(wb):
    plain = wb.derive("t:Parts", fill=[], group_by=["Group"])
    assert plain.filled == () and plain.rows == wb.derive("t:Parts", group_by=["Group"]).rows


def test_a_filled_table_is_a_joined_table_the_engine_reads(wb):
    from xlsx_agent.fill import filled_table

    table = filled_table(wb, wb.region("t:Parts"), FILL)
    assert isinstance(table, FilledTable) and table.joined is None
    assert table.marks == {(2, 2): "Default limit", (3, 2): "Default limit"}


# -- the tool ----------------------------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[xlsx_format()],
                 groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box):
    session = box.session()
    blob = session.add_blob(book().to_bytes(), name="parts.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_derive_fills_and_marks_each_filled_value(box):
    session = _session(box)
    result = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "limits",
        "fill": [{"column": "Limit", "from": "Default limit", "coerce": "number"}]})
    assert result.ok, result.error
    data = json.loads(json.dumps(result.data))
    assert data["filled"][0]["filled"] == {"count": 2, "first": ["Parts!C3", "Parts!C4"]}
    assert data["filled"][0]["unfilled"]["wrong_type"]["count"] == 1
    assert data["rows"][1][2] == {"a": "d:limits/#2/Limit", "v": 40, "t": "number",
                                  "sources": ["t:Parts/P-2/Default limit"],
                                  "filled": "Default limit"}
    assert "filled" not in data["rows"][0][2]
    read = box.dispatch(session, "xlsx_read", {"target": "d:limits/#3/Limit"})
    assert read.ok and read.data["rows"][0][0]["filled"] == "Default limit"
    again = box.dispatch(session, "xlsx_read", {"target": "d:limits/#3/Limit",
                                                "sources": True})
    assert again.ok and again.data["sources"] == ["t:Parts/P-3/Default limit"]


def test_xlsx_derive_takes_a_strict_calls_placeholders_and_refuses_with_the_way_out(box):
    session = _session(box)
    taken = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "plain", "group_by": ["Group"],
        "fill": [{"column": "", "from": "", "coerce": None}]})
    assert taken.ok, taken.error
    assert "filled" not in taken.data
    refused = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "x",
        "fill": [{"column": "Limt", "from": "Default limit"}]})
    assert not refused.ok and refused.error.field == "fill"
    assert "Limit" in refused.error.valid_options
    refused = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "x",
        "fill": [{"column": "Limit", "from": "Default limit", "coerce": "text"}]})
    assert not refused.ok and refused.error.field.startswith("fill")


def test_xlsx_derive_fills_a_dataset_from_a_looked_up_one(box):
    session = _session(box)
    assert box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Groups", "id": "groups",
        "group_by": ["Group"], "measures": [{"fn": "max", "column": "Group limit",
                                             "as": "limit"}]}).ok
    join = {"from": "d:groups", "on": ["Group", "Group"], "columns": ["limit"]}
    # the joined "limit" is renamed beside Limit: "limit" names Limit itself, and says so
    same = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "limits", "join": join,
        "fill": [{"column": "Limit", "from": "limit"}]})
    assert same.ok, same.error
    assert same.data["filled"][0]["note"].endswith("the joined limit is limit (groups)")
    result = box.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Parts", "id": "limits", "join": join,
        "fill": [{"column": "Limit", "from": "limit (groups)"}]})
    assert result.ok, result.error
    limits = [row[2]["v"] for row in result.data["rows"]]
    assert limits == [10, 30, 20, 20, 20, None, None]
    assert result.data["rows"][1][2]["sources"] == ["Groups!B2"]
    assert result.data["joined"]["matched"] == 5


def test_a_dataset_value_without_a_fill_is_not_marked(wb):
    table = DatasetTable(wb.derive("t:Parts", group_by=["Group"]), "by_group")
    assert table.filled_from == {} and "filled" not in table.value(1, 0).as_dict()
