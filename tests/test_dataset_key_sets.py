"""Round 40, R42: a dataset's rows named by a key set, as a table's are
(``d:by_region/{South,North}/total``), with the same lookups and refusals; a dataset with
no key column refused with the way out -- the column to key on, and the call that works.
Made-up workbooks only (tests/build.py)."""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Dataset, Toolbox

from test_datasets import BY_REP, _session, call, err
from test_derive import sales

from xlsx_agent import DatasetTable, NameNotFound, RegionError, Workbook
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

BY_REGION = {"group_by": ["Region"], "measures": [{"fn": "sum", "column": "Amount",
                                                   "as": "total"}], "round": 2}


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return Workbook.open(sales().to_bytes())


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _values(rows):
    return [[v.v for v in row] for row in rows]


# -- the library -------------------------------------------------------------------------------


def test_a_key_set_names_its_rows_in_its_order(wb):
    table = DatasetTable(wb.derive("t:Sales", id="by_region", **BY_REGION))
    selection = table.select("d:by_region/{South,North}/total")
    assert selection.text == "d:by_region/{South,North}/total"
    assert selection.kind == "values" and selection.order == (2, 1) and selection.rows == 2
    rows = table.read("d:by_region/{South,North}")
    assert [[v.a for v in row] for row in rows] == [
        ["d:by_region/#2/Region", "d:by_region/#2/total"],
        ["d:by_region/#1/Region", "d:by_region/#1/total"]]
    # each row as the key alone names it
    assert _values(rows) == _values(table.read("d:by_region/South")) \
        + _values(table.read("d:by_region/North"))
    assert table.select("d:by_region/{East,North}/*").kind == "rows"


def test_a_key_set_on_a_dataset_refuses_as_on_a_table(wb):
    table = DatasetTable(wb.derive("t:Sales", id="by_region", **BY_REGION))
    with pytest.raises(NameNotFound) as caught:
        table.read("d:by_region/{South,Nrth}")
    # the second key, by its place, never its text; the near row by position
    assert "key 2 of the set" in caught.value.message and "Nrth" not in caught.value.message
    assert caught.value.valid_options == ["d:by_region/#1"]
    twice = DatasetTable(Dataset(("k", "v"), [("a", 1), ("a", 2), ("b", 3)], keys=1), "twice")
    with pytest.raises(RegionError) as shared:
        twice.read("d:twice/{a,b}")
    assert shared.value.code == "duplicate_key"
    assert shared.value.valid_options == ["d:twice/#1", "d:twice/#2"]


def test_a_dataset_without_one_key_column_says_which_column_can_be_the_key(wb):
    by_rep = DatasetTable(wb.derive("t:Sales", id="by_rep", **BY_REP))
    for address in ("d:by_rep/{Ann,Bob}", "d:by_rep/Ann"):
        with pytest.raises(RegionError) as caught:
            by_rep.read(address)
        error = caught.value
        assert error.code == "no_key_column"
        assert error.valid_options == by_rep.key_columns() == ["Rep"]
        assert 'key: "Rep"' in error.message and "d:by_rep/#1:#6" in error.message
        assert "Ann" not in error.message
    # keyed on the column the refusal named
    rows = by_rep.read("d:by_rep/{Ann,Bob}/total", key="Rep")
    assert [[v.a for v in row] for row in rows] == [["d:by_rep/#1/total"],
                                                    ["d:by_rep/#2/total"]]
    plain = DatasetTable(Dataset(("v",), [(1.5,), (1.5,)]), "plain")
    with pytest.raises(RegionError) as caught:
        plain.read("d:plain/{a,b}")
    assert caught.value.valid_options == [] and "no column holds a key" in caught.value.message


# -- the tools ---------------------------------------------------------------------------------


def test_xlsx_read_takes_a_key_set_of_a_dataset_paged(box):
    session = _session(box, sales())
    assert call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Sales"}, id="by_region",
                **BY_REGION).ok
    result = call(box, session, "xlsx_read", target="d:by_region/{South,East,North}/total",
                  limit=2)
    assert result.ok, result.error
    assert result.data["target"] == "d:by_region/{South,East,North}/total"
    assert [row[0]["a"] for row in result.data["rows"]] == ["d:by_region/#2/total",
                                                            "d:by_region/#3/total"]
    assert result.total == 3 and result.data["cut"] == {"rows": 1}
    more = call(box, session, "xlsx_read", target="d:by_region/{South,East,North}/total",
                limit=2, cursor=result.next_cursor)
    assert [row[0]["a"] for row in more.data["rows"]] == ["d:by_region/#1/total"]
    picked = call(box, session, "xlsx_read", target="d:by_region/{East,South}",
                  columns=["total"])
    assert [row[0]["a"] for row in picked.data["rows"]] == ["d:by_region/#3/total",
                                                            "d:by_region/#2/total"]
    traced = call(box, session, "xlsx_trace", doc="d1", rows=["d:by_region/{South,North}"])
    assert traced.ok, traced.error
    assert [t["row"] for t in traced.data["traced"]] == ["d:by_region/#2", "d:by_region/#1"]


def test_xlsx_read_refuses_a_keyless_dataset_with_the_call_that_works(box):
    session = _session(box, sales())
    assert call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Sales"}, id="by_rep",
                **BY_REP).ok
    refused = err(call(box, session, "xlsx_read", target="d:by_rep/{Ann,Bob}"))
    assert refused["code"] == "refused" and refused["field"] == "target"
    assert refused["valid_options"] == ["Rep"]
    works = refused["details"]["call"]
    assert works == {"target": "d:by_rep/{Ann,Bob}", "key": "Rep"}
    assert 'The call that works: xlsx_read {"target": "d:by_rep/{Ann,Bob}", "key": "Rep"}' \
        in refused["message"]
    result = call(box, session, "xlsx_read", **works)
    assert result.ok and result.total == 2
    unknown = err(call(box, session, "xlsx_read", target="d:by_rep/{Ann,Bob}", key="Agent"))
    assert unknown["code"] == "not_found" and unknown["field"] == "key"
    assert "Rep" in unknown["valid_options"]


def test_a_key_with_no_row_by_key_is_a_no_op_said(box):
    session = _session(box, sales())
    assert call(box, session, "xlsx_derive", doc="d1", **{"from": "t:Sales"}, id="by_rep",
                **BY_REP).ok
    result = call(box, session, "xlsx_read", target="d:by_rep[total]", key="Rep")
    assert result.ok and result.total == 6
    assert result.warnings == ["key 'Rep' not used: it names the column holding the keys of "
                               "a row by key or a key set (d:by_rep/{a,b}), and "
                               "d:by_rep[total] names none"]
