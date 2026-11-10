"""Round 49, R95: ``is_true`` and ``is_false``, the filters of a TRUE/FALSE column that take
no value (a strict caller's schema has no boolean value to give), and ``eq`` without a value
refused naming them; in the library, ``xlsx_query``, ``xlsx_derive`` and of a dataset.
Made-up workbooks only (tests/build.py)."""

from __future__ import annotations

import pytest

from build import WorkbookBuilder
from ooxml_edit.tools import Toolbox

from test_datasets import _session, call, err

from xlsx_agent import QueryError, Workbook
from xlsx_agent.query import OPS, OPS_BY_TYPE
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

ROWS = [["Code", "Kept", "Qty"],
        ["K7-12.3", True, 3],
        ["K7-12.4", False, 4],
        ["K7-13.1", True, 5],
        ["K7-13.2", None, 1],          # blank: neither TRUE nor FALSE
        ["K7-14.0", "yes", 2],         # text in the TRUE/FALSE column: not compared
        ["K7-14.1", False, 6]]


def items() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Items")
    sheet.rows("A1", ROWS)
    sheet.table("Items", f"A1:C{len(ROWS)}", ROWS[0], write_header=False)
    return book


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return Workbook.open(items().to_bytes())


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _codes(result):
    return [row[0].v for row in result.rows]


def _filter(op, **more):
    return [{"column": "Kept", "op": op, **more}]


# -- the tests ---------------------------------------------------------------------------------


def test_is_true_and_is_false_are_tests_of_a_true_false_column_only():
    assert "is_true" in OPS and "is_false" in OPS
    assert {"is_true", "is_false"} <= set(OPS_BY_TYPE["bool"])
    for other in ("text", "number", "date", "error", "list", "blank"):
        assert "is_true" not in OPS_BY_TYPE[other] and "is_false" not in OPS_BY_TYPE[other]


def test_is_true_and_is_false_find_the_true_and_the_false_cells(wb):
    table = wb.table("t:Items")
    true = table.query(["Code"], where=_filter("is_true"))
    false = table.query(["Code"], where=_filter("is_false"))
    assert _codes(true) == ["K7-12.3", "K7-13.1"]
    assert _codes(false) == ["K7-12.4", "K7-14.1"]
    # a blank is neither, and the text cell is counted as another type's, as for eq
    assert true.not_compared["Kept"].count == 1 and false.not_compared["Kept"].count == 1
    # the same rows as eq with the value as text, or as Python's
    assert _codes(table.query(["Code"], where=_filter("eq", value="true"))) == _codes(true)
    assert _codes(table.query(["Code"], where={"Kept": {"eq": False}})) == _codes(false)


def test_the_mapping_form_writes_them_as_blank_is(wb):
    table = wb.table("t:Items")
    assert _codes(table.query(["Code"], where={"Kept": {"is_true": True}})) == \
        ["K7-12.3", "K7-13.1"]
    with pytest.raises(QueryError) as caught:
        table.query(where={"Kept": {"is_true": False}})
    assert caught.value.code == "bad_filter"
    assert str(caught.value) == "is_true on Kept is written is_true: true; for the other, " \
        "use is_false"


def test_derive_filters_with_them(wb):
    derived = wb.derive("t:Items", where=_filter("is_false"),
                        measures=[{"fn": "sum", "column": "Qty", "as": "qty"}])
    assert derived.rows == [{"qty": 10}]
    assert derived.not_compared["Kept"].count == 1


# -- rule 2: a refusal carries the way out ----------------------------------------------------


@pytest.mark.parametrize("where", [_filter("eq", value=None), _filter("eq"),
                                   {"Kept": {"eq": None}}])
def test_eq_without_a_value_on_a_true_false_column_names_is_true_and_is_false(wb, where):
    with pytest.raises(QueryError) as caught:
        wb.table("t:Items").query(where=where)
    assert caught.value.code == "bad_filter"
    assert str(caught.value) == (
        'eq on Kept needs a value: Kept holds TRUE/FALSE, so use is_true or is_false '
        '(no value), or eq with value "true" or "false"')
    assert list(caught.value.valid_options) == ["is_true", "is_false"]


def test_eq_without_a_value_on_another_column_names_blank(wb):
    with pytest.raises(QueryError) as caught:
        wb.table("t:Items").query(where=[{"column": "Qty", "op": "eq", "value": None}])
    assert str(caught.value) == "eq on Qty needs a value; for its empty cells use blank"
    assert "blank" in caught.value.valid_options


def test_on_another_column_type_they_are_refused_with_the_tests_it_takes(wb):
    with pytest.raises(QueryError) as caught:
        wb.table("t:Items").query(where=[{"column": "Code", "op": "is_true"}])
    assert caught.value.code == "wrong_type"
    assert "holds no TRUE/FALSE values" in str(caught.value)
    assert "eq matches its text exactly" in str(caught.value)
    assert "eq" in caught.value.valid_options and "is_true" not in caught.value.valid_options


def test_a_value_that_says_the_other_is_refused_naming_it(wb):
    with pytest.raises(QueryError) as caught:
        wb.table("t:Items").query(where=_filter("is_true", value="false"))
    assert str(caught.value) == ("is_true on Kept takes no value, and this one says the "
                                 "other: use is_false, with no value")
    assert list(caught.value.valid_options) == ["is_false"]
    with pytest.raises(QueryError) as caught:
        wb.table("t:Items").query(where=_filter("is_false", value="maybe"))
    assert str(caught.value) == "is_false on Kept takes no value: drop value"
    with pytest.raises(QueryError) as caught:
        wb.table("t:Items").query(where=_filter("is_false", values=["false"]))
    assert str(caught.value) == "is_false on Kept takes no value: drop values"


# -- rule 1: intent accepted, said ------------------------------------------------------------


@pytest.mark.parametrize("value", ["true", " TRUE ", True, ""])
def test_a_value_that_says_the_same_or_a_placeholder_changes_nothing_said(box, value):
    session = _session(box, items())
    result = call(box, session, "xlsx_query", doc="d1", table="t:Items",
                  where=_filter("is_true", value=value) if value is not True
                  else _filter("is_true"))
    assert result.ok, result.error
    assert [row[0]["v"] for row in result.data["rows"]] == ["K7-12.3", "K7-13.1"]
    if value is not True:
        assert result.warnings == ["is_true on Kept takes no value: its value changes "
                                   "nothing"]


# -- the tools ---------------------------------------------------------------------------------


def test_the_tools_take_them_and_a_dataset_too(box):
    session = _session(box, items())
    derived = call(box, session, "xlsx_derive", doc="d1", id="kept_qty",
                   group_by=["Code", "Kept"], **{"from": "t:Items"},
                   where=_filter("not_blank"))
    assert derived.ok, derived.error
    found = call(box, session, "xlsx_query", table="d:kept_qty", where=_filter("is_true"))
    assert found.ok, found.error
    assert [row[0]["v"] for row in found.data["rows"]] == ["K7-12.3", "K7-13.1"]
    summed = call(box, session, "xlsx_derive", doc="d1", id="false_qty", **{"from": "t:Items"},
                  where=_filter("is_false"), measures=[{"fn": "sum", "column": "Qty"}])
    assert summed.ok, summed.error
    assert summed.data["rows"][0][0]["v"] == 10


def test_the_tool_refusal_for_eq_with_null_names_them(box):
    session = _session(box, items())
    for where in (_filter("eq", value=None), _filter("eq")):
        result = call(box, session, "xlsx_derive", doc="d1", id="kept", where=where,
                      **{"from": "t:Items"})
        assert not result.ok
        error = err(result)
        assert error["code"] == "invalid_arguments"
        assert "use is_true or is_false" in error["message"]
        assert error["valid_options"] == ["is_true", "is_false"]


def test_a_json_boolean_value_is_still_the_shared_checks_refusal(box):
    """The schema's value is a string (the common strict subset has no unions); the shared
    argument check refuses a JSON true before the filter reads it."""
    session = _session(box, items())
    result = call(box, session, "xlsx_query", doc="d1", table="t:Items",
                  where=_filter("eq", value=True))
    assert not result.ok and err(result)["message"] == "where[0].value must be a string"
