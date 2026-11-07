"""Round 40, R43: ``contains_any`` (text holding any of a list), ``loose`` (text compared
ignoring case and every white space character), and results that say when matching was
loosened; in ``xlsx_query``, of a dataset, and in ``xlsx_derive``'s filters.  Made-up
workbooks only (tests/build.py)."""

from __future__ import annotations

import pytest

from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from test_datasets import _session, call, err

from xlsx_agent import QueryError, Workbook
from xlsx_agent.query import OPS, OPS_BY_TYPE, loose_text
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

ROWS = [["Code", "Name", "Qty", "Tags"],
        ["AB.12", "Bolt", 4, "steel"],
        ["ab. 12", "Nut", 2, "Steel"],
        ["CD.7", "Washer", 9, "zinc"],
        ["CD\u00a0.7", "Pin", 1, "brass"],
        ["EF.1", "Stra\u00dfe", 3, "steel"],
        ["XY\u200b.9", "Rivet", 5, "\tzinc "]]


def parts() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Parts")
    sheet.rows("A1", ROWS)
    sheet.table("Parts", f"A1:D{len(ROWS)}", ROWS[0], write_header=False)
    return book


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return Workbook.open(parts().to_bytes())


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _codes(result):
    return [row[0].v for row in result.rows]


@pytest.mark.parametrize("text, loose", [
    ("AB. 12", "ab.12"),
    (" a\tb\nc\r\u00a0d\u2003e\u3000f\u200bg\ufeffh ", "abcdefgh"),
    ("Stra\u00dfe", "strasse"),                     # full case folding
    ("Caf\u00e9", "caf\u00e9"),                     # an accent counts
    ("A-B", "a-b"),                                 # punctuation counts
])
def test_loose_text_drops_white_space_and_folds_case(text, loose):
    assert loose_text(text) == loose


def test_contains_any_is_a_text_test_of_a_list(wb):
    assert "contains_any" in OPS
    assert "contains_any" in OPS_BY_TYPE["text"] and "contains_any" in OPS_BY_TYPE["list"]
    assert "contains_any" not in OPS_BY_TYPE["number"]
    table = wb.table("t:Parts")
    found = table.query(["Code"], where=[{"column": "Code", "op": "contains_any",
                                          "values": ["ab", "xy"]}])
    assert _codes(found) == ["AB.12", "ab. 12", "XY\u200b.9"]      # case ignored, as contains
    with pytest.raises(QueryError) as caught:
        table.query(where=[{"column": "Qty", "op": "contains_any", "values": ["1"]}])
    assert caught.value.code == "wrong_type" and "contains_any" not in caught.value.valid_options
    for bad in ([], [""], "ab"):
        with pytest.raises(QueryError) as caught:
            table.query(where=[{"column": "Code", "op": "contains_any", "values": bad}])
        assert caught.value.code == "bad_filter"


def test_contains_given_a_list_is_contains_any_said(wb):
    table = wb.table("t:Parts")
    found = table.query(["Code"], where={"Code": {"contains": ["ab", "xy"]}})
    assert _codes(found) == ["AB.12", "ab. 12", "XY\u200b.9"]


def test_loose_matches_ignoring_case_and_white_space_and_says_which_rows(wb):
    table = wb.table("t:Parts")
    exact = table.query(["Code"], where=[{"column": "Code", "op": "in",
                                          "values": ["AB.12", "CD.7", "XY.9"]}])
    assert _codes(exact) == ["AB.12", "CD.7"]
    loose = table.query(["Code"], where=[{"column": "Code", "op": "in", "loose": True,
                                          "values": ["AB.12", "CD.7", "XY.9"]}])
    assert _codes(loose) == ["AB.12", "ab. 12", "CD.7", "CD\u00a0.7", "XY\u200b.9"]
    assert _codes(table.query(["Code"], where={"Code": {"eq": "ab.12", "loose": True}})) \
        == ["AB.12", "ab. 12"]
    assert _codes(table.query(["Code"], where={"Code": {"contains": "D.7", "loose": True}})) \
        == ["CD.7", "CD\u00a0.7"]
    assert _codes(table.query(["Code"], where={"Name": {"eq": "STRASSE", "loose": True}})) \
        == ["EF.1"]


def test_matches_count_the_rows_matched_only_loosely(wb):
    from xlsx_agent.query import matches

    region = wb.table("t:Parts")
    found = matches(wb, region, where=[{"column": "Code", "op": "in", "loose": True,
                                        "values": ["AB.12", "CD.7"]}])
    assert found.loosened["Code"].count == 2
    assert found.loosened["Code"].first == ("Parts!A3", "Parts!A5")
    same = matches(wb, region, where=[{"column": "Code", "op": "eq", "value": "CD.7",
                                       "loose": True}, {"column": "Qty", "op": "gte",
                                                        "value": "5"}])
    assert same.loosened["Code"].count == 0         # loose changed nothing: said, as 0
    assert matches(wb, region, where={"Code": {"eq": "CD.7"}}).loosened == {}


def test_xlsx_query_says_when_matching_was_loosened(box):
    session = _session(box, parts())
    result = call(box, session, "xlsx_query", doc="d1", table="t:Parts", columns=["Code"],
                  where=[{"column": "Code", "op": "contains_any", "loose": True,
                          "values": ["ab.12", "xy.9"]}])
    assert result.ok, result.error
    assert result.total == 3
    assert result.data["loosened"] == {"Code": {"count": 2, "first": ["Parts!A3",
                                                                      "Parts!A7"]}}
    listed = call(box, session, "xlsx_query", doc="d1", table="t:Parts", columns=["Code"],
                  where=[{"column": "Code", "op": "contains", "values": ["ab", "xy"]}])
    assert listed.total == 3
    assert listed.warnings == ["contains on Code was given a list: read as contains_any, any "
                               "of them"]
    number = call(box, session, "xlsx_query", doc="d1", table="t:Parts",
                  where=[{"column": "Qty", "op": "eq", "value": "4", "loose": True}])
    assert number.ok and number.total == 1 and "loosened" not in number.data
    assert number.warnings == ["loose on Qty changes nothing: it loosens text tests (eq, in, "
                               "contains, contains_any), and this is eq on a number column"]


def test_no_exact_match_says_how_many_loose_would_find(box):
    session = _session(box, parts())
    result = call(box, session, "xlsx_query", doc="d1", table="t:Parts",
                  where=[{"column": "Code", "op": "eq", "value": "AB .12"}])
    assert result.ok and result.total == 0
    assert result.warnings == ["no row matches exactly; 2 would with loose: true on where[0] "
                               "(text compared ignoring case and white space, on Code)"]
    nothing = call(box, session, "xlsx_query", doc="d1", table="t:Parts",
                   where=[{"column": "Code", "op": "eq", "value": "ZZ.0"}])
    assert nothing.total == 0 and nothing.warnings == []


def test_a_dataset_is_queried_loosely_as_a_table(box):
    session = _session(box, parts())
    assert call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Parts"}, id="by_tag",
                group_by=["Tags"], measures=[{"fn": "sum", "column": "Qty", "as": "q"}]).ok
    result = call(box, session, "xlsx_query", table="d:by_tag",
                  where=[{"column": "Tags", "op": "in", "loose": True,
                          "values": ["steel", "zinc"]}])
    assert result.ok, result.error
    assert [row[0]["v"] for row in result.data["rows"]] == ["steel", "Steel", "zinc",
                                                            "\tzinc "]
    assert result.data["loosened"]["Tags"]["count"] == 2
    none = call(box, session, "xlsx_query", table="d:by_tag",
                where=[{"column": "Tags", "op": "eq", "value": "ZINC"}])
    assert none.total == 0 and none.warnings[0].startswith("no row matches exactly; 2 would")


def test_xlsx_derive_filters_loosely_and_says_so(box):
    session = _session(box, parts())
    result = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Parts"}, id="ab",
                  where=[{"column": "Code", "op": "eq", "value": "ab . 12", "loose": True}],
                  measures=[{"fn": "sum", "column": "Qty", "as": "q"}])
    assert result.ok, result.error
    assert result.data["matched"] == 2
    assert result.data["loosened"] == {"Code": {"count": 2, "first": ["Parts!A2",
                                                                      "Parts!A3"]}}
    listed = call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Parts"}, id="cd",
                  where=[{"column": "Code", "op": "contains", "values": ["cd", "ef"]}],
                  measures=[{"fn": "count", "as": "n"}])
    assert listed.ok and listed.data["matched"] == 3
    assert listed.warnings == ["contains on Code was given a list: read as contains_any, any "
                               "of them"]


def test_loose_is_true_or_false(box):
    session = _session(box, parts())
    refused = call(box, session, "xlsx_query", doc="d1", table="t:Parts",
                   where=[{"column": "Code", "op": "eq", "value": "x", "loose": "yes"}])
    assert not refused.ok and err(refused)["code"] == "invalid_arguments"
