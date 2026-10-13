"""R8: ``table.query`` and ``xlsx_query``: every filter, the refusal of a filter its column's
type does not take, the sort, and the page; the library first, then through a real
ooxml-edit Toolbox, bounded and paged on large fixtures.

Every workbook is built with tests/build.py.
"""

import datetime as dt
import json
import logging
from datetime import date, datetime, time

import pytest

from build import Error, Formula, IsoDate, WorkbookBuilder
from ooxml_edit.tools import Limits, Toolbox
from test_catalogue import MARKER, ROWS, WIDE, warehouse

from xlsx_agent import NameNotFound, QueryError, Workbook
from xlsx_agent.query import OPS, OPS_BY_TYPE, matches
from xlsx_agent.read.types import TYPES, day_serial, serial_day
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS
from xlsx_agent.tools.read import READ_ROWS

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731

COLUMNS = ["SKU", "Category", "Price", "Share", "Added", "At", "Opens", "Active", "Note"]


def goods(date1904: bool = False) -> WorkbookBuilder:
    """A table with a column of every type a filter compares: text, number (one number
    stored as text, one blank, one formula without a cached value), percent, date (one with
    a time of day), datetime, time, bool, and a column with no values."""
    book = WorkbookBuilder(date1904=date1904)
    pct = book.style(num_fmt="0%")
    sheet = book.sheet("Stock")
    sheet.rows("A1", [COLUMNS])
    rows = [
        ["SKU-1", "Hardware", 12.5, 0.25, date(2026, 1, 5), datetime(2026, 1, 5, 9, 30),
         time(9, 0), True],
        ["SKU-2", "hardware", "7.25", 0.5, datetime(2026, 1, 6, 18, 0),
         datetime(2026, 1, 6, 12, 0), time(10, 30), False],
        ["SKU-3", "Tools", 3, 0.1, date(2026, 2, 1), datetime(2026, 2, 1, 0, 0),
         time(12, 0), True],
        ["SKU-4", "Garden", None, 1, date(2026, 3, 1), datetime(2026, 3, 1, 23, 59, 59),
         time(17, 45), False],
        ["SKU-5", f"Hand tools {MARKER}", Formula("C2*2"), 0.75, date(2026, 1, 5),
         datetime(2026, 1, 5, 9, 30), time(9, 0), True],
        ["SKU-6", "Tools", 3.0, 0.2, date(2025, 12, 31), datetime(2025, 12, 31, 8, 0),
         time(8, 0), False],
    ]
    for r, values in enumerate(rows, start=2):
        for c, value in enumerate(values):
            if value is None:
                continue
            ref = f"{chr(65 + c)}{r}"
            if c == 3:
                sheet.cell(ref, value, style=pct)
            elif c == 4 and isinstance(value, datetime):
                sheet.cell(ref, value, style=book.style(num_fmt="yyyy-mm-dd"))
            else:
                sheet.cell(ref, value)
    sheet.table("Goods", "A1:I7", COLUMNS, write_header=False)
    return book


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


@pytest.fixture(scope="module")
def table():
    return _open(goods()).table("t:Goods")


def _skus(result) -> list[str]:
    return [row[0].v for row in result.rows]


# -- the library: every filter ----------------------------------------------------------------


@pytest.mark.parametrize("where,expected", [
    # text: eq and in exactly (case counts); contains ignores case
    ({"Category": {"eq": "Hardware"}}, ["SKU-1"]),
    ({"Category": {"eq": "hardware"}}, ["SKU-2"]),
    ({"Category": {"eq": "Hardware "}}, []),
    ({"Category": {"in": ["Tools", "Garden", "Paint"]}}, ["SKU-3", "SKU-4", "SKU-6"]),
    ({"Category": {"contains": "HARD"}}, ["SKU-1", "SKU-2"]),
    ({"Category": {"contains": "tool"}}, ["SKU-3", "SKU-5", "SKU-6"]),
    # numbers: as stored, 3 and 3.0 alike; text in its JSON form; a number as text never
    ({"Price": {"eq": 3}}, ["SKU-3", "SKU-6"]),
    ({"Price": {"eq": "3.0"}}, ["SKU-3", "SKU-6"]),
    ({"Price": {"in": [12.5, "7.25"]}}, ["SKU-1"]),
    ({"Price": {"gte": 5}}, ["SKU-1"]),
    ({"Price": {"lte": "3"}}, ["SKU-3", "SKU-6"]),
    ({"Price": {"gte": 3, "lte": 12.5}}, ["SKU-1", "SKU-3", "SKU-6"]),
    # percent: the stored fraction
    ({"Share": {"gte": 0.5}}, ["SKU-2", "SKU-4", "SKU-5"]),
    ({"Share": {"eq": "0.25"}}, ["SKU-1"]),
    # dates: the day as the reader shows it; ISO text or a serial
    ({"Added": {"eq": "2026-01-05"}}, ["SKU-1", "SKU-5"]),
    ({"Added": {"eq": "2026-01-06"}}, ["SKU-2"]),        # shown as its day, time aside
    ({"Added": {"lte": "2026-01-06"}}, ["SKU-1", "SKU-2", "SKU-5", "SKU-6"]),
    ({"Added": {"gte": "2026-02-01", "lte": "2026-03-01"}}, ["SKU-3", "SKU-4"]),
    ({"Added": {"eq": 46027}}, ["SKU-1", "SKU-5"]),       # 2026-01-05's serial
    ({"Added": {"in": ["2026-03-01", "2025-12-31"]}}, ["SKU-4", "SKU-6"]),
    # datetime: the moment; a date alone is its midnight
    ({"At": {"gte": "2026-01-06"}}, ["SKU-2", "SKU-3", "SKU-4"]),
    ({"At": {"eq": "2026-01-05T09:30:00"}}, ["SKU-1", "SKU-5"]),
    ({"At": {"eq": "2026-01-05T09:30"}}, ["SKU-1", "SKU-5"]),
    ({"At": {"lte": "2026-02-01T00:00:00"}}, ["SKU-1", "SKU-2", "SKU-3", "SKU-5", "SKU-6"]),
    # time: the time of day
    ({"Opens": {"gte": "10:00", "lte": "12:00:00"}}, ["SKU-2", "SKU-3"]),
    ({"Opens": {"eq": 0.375}}, ["SKU-1", "SKU-5"]),          # 09:00 as a serial
    # bool
    ({"Active": {"eq": True}}, ["SKU-1", "SKU-3", "SKU-5"]),
    ({"Active": {"eq": "false"}}, ["SKU-2", "SKU-4", "SKU-6"]),
    # blank: no value, a formula without a cached value among them
    ({"Price": {"blank": True}}, ["SKU-4", "SKU-5"]),
    ({"Price": {"not_blank": True}}, ["SKU-1", "SKU-2", "SKU-3", "SKU-6"]),
    ({"Note": {"blank": True}}, ["SKU-1", "SKU-2", "SKU-3", "SKU-4", "SKU-5", "SKU-6"]),
    ({"Note": {"not_blank": True}}, []),
    # tests combine with and, across columns and within one
    ({"Category": {"contains": "tools"}, "Active": {"eq": False}}, ["SKU-6"]),
    ({"category": {"eq": "Tools"}, "Price": {"eq": 3}, "Added": {"lte": "2025-12-31"}},
     ["SKU-6"]),
], ids=lambda value: json.dumps(value) if isinstance(value, dict) else None)
def test_each_filter(table, where, expected):
    assert _skus(table.query(["SKU"], where=where)) == expected


def test_the_list_form_is_the_mapping_form(table):
    listed = table.query(where=[{"column": "Category", "op": "in", "values": ["Tools"]},
                                {"column": "Price", "op": "gte", "value": "3"},
                                {"column": "Note", "op": "blank"}])
    mapped = table.query(where={"Category": {"in": ["Tools"]}, "Price": {"gte": 3},
                                "Note": {"blank": True}})
    assert listed == mapped and _skus(listed) == ["SKU-3", "SKU-6"]


def test_a_result_gives_addresses_positions_types_and_the_count(table):
    result = table.query(["Price", "SKU"], where={"Category": {"eq": "Tools"}}, sort=["SKU"])
    assert result.address == "t:Goods" and result.columns == ("Price", "SKU")
    assert [[v.a for v in row] for row in result.rows] == [["Stock!C4", "Stock!A4"],
                                                           ["Stock!C7", "Stock!A7"]]
    assert result.positions == (3, 6) and result.matches == 2
    assert result.types == {"Category": "text", "SKU": "text"}
    data = result.as_dict()
    assert data["rows"][0] == [{"a": "Stock!C4", "v": 3, "t": "number"},
                               {"a": "Stock!A4", "v": "SKU-3", "t": "text"}]
    every = table.query()
    assert every.columns == tuple(COLUMNS) and every.matches == 6 and every.types == {}


def test_cells_of_another_type_never_match_and_are_counted(table):
    result = table.query(["SKU"], where={"Price": {"gte": 0}})
    assert _skus(result) == ["SKU-1", "SKU-3", "SKU-6"]      # "7.25" is text: not compared
    assert result.types == {"Price": "number"}
    assert result.not_compared["Price"].as_dict() == {"count": 1, "first": ["Stock!C3"]}
    assert table.query(where={"Price": {"blank": True}}).not_compared == {}


# -- the library: refusals --------------------------------------------------------------------


@pytest.mark.parametrize("where,column,column_type", [
    ({"Price": {"gte": "cheap"}}, "Price", "number"),        # text on a number column
    ({"Price": {"eq": True}}, "Price", "number"),
    ({"Price": {"in": [1, "two"]}}, "Price", "number"),
    ({"Price": {"contains": "1"}}, "Price", "number"),
    ({"Category": {"gte": "M"}}, "Category", "text"),         # text is not ordered here
    ({"Category": {"eq": 5}}, "Category", "text"),
    ({"Added": {"gte": "05/01/2026"}}, "Added", "date"),
    ({"Added": {"eq": "2026-02-30"}}, "Added", "date"),
    ({"Opens": {"eq": "2026-01-05"}}, "Opens", "time"),
    ({"Active": {"gte": True}}, "Active", "bool"),
    ({"Active": {"eq": "yes"}}, "Active", "bool"),
    ({"Note": {"eq": "x"}}, "Note", "blank"),                 # a column with no values
], ids=lambda value: json.dumps(value) if isinstance(value, dict) else None)
def test_a_wrong_type_filter_is_refused_naming_the_type(table, where, column, column_type):
    with pytest.raises(QueryError) as caught:
        table.query(where=where)
    error = caught.value
    assert (error.code, error.column, error.column_type) == ("wrong_type", column, column_type)
    assert error.valid_options == list(OPS_BY_TYPE[column_type])
    assert f"{column} is a {column_type} column" in str(error)
    for value in ("cheap", "two", "05/01/2026", "yes"):
        assert value not in str(error)


@pytest.mark.parametrize("where", [
    {"Price": {"between": [1, 2]}},
    {"Price": {"blank": False}},
    {"Price": {"in": []}},
    {"Price": {"in": "12.5"}},
    {"Price": {"eq": None}},
    {"Price": {}},
    {"Category": {"contains": ""}},
    [{"column": "Price", "op": "in", "value": "1"}],
    [{"column": "Price", "op": "blank", "value": "1"}],
    [{"column": "Price", "op": "eq", "values": ["1"]}],
    [{"column": "Price", "op": "eq", "value": "1", "extra": 1}],
    "Price",
], ids=str)
def test_a_malformed_filter_is_refused(table, where):
    with pytest.raises(QueryError) as caught:
        table.query(where=where)
    assert caught.value.code == "bad_filter"


def test_unknown_columns_and_bad_sorts_are_refused(table):
    with pytest.raises(NameNotFound) as caught:
        table.query(where={"Prise": {"eq": 1}})
    assert caught.value.valid_options == ["Price"]
    with pytest.raises(NameNotFound):
        table.query(["SKU", "Nope"])
    with pytest.raises(NameNotFound):
        table.query(sort=["Nope"])
    for sort in ([("Price", "down")], [3], [{"column": "Price", "order": "desc"}]):
        with pytest.raises(QueryError) as bad:
            table.query(sort=sort)
        assert bad.value.code == "bad_sort"
    with pytest.raises(ValueError):
        table.query(limit=-1)


def test_every_type_has_its_tests():
    assert set(OPS_BY_TYPE) == set(TYPES)
    assert all(set(ops) <= set(OPS) for ops in OPS_BY_TYPE.values())


# -- the library: sort, limit, offset ---------------------------------------------------------


def test_sorting_puts_other_types_then_blanks_last_either_way(table):
    up = table.query(["SKU"], sort=["Price"])
    assert _skus(up) == ["SKU-3", "SKU-6", "SKU-1", "SKU-2", "SKU-4", "SKU-5"]
    down = table.query(["SKU"], sort=[("Price", "desc")])
    assert _skus(down) == ["SKU-1", "SKU-3", "SKU-6", "SKU-2", "SKU-4", "SKU-5"]
    assert down.types == {"Price": "number"}


def test_sorting_by_several_keys_text_case_aside_and_dates(table):
    result = table.query(["SKU"], sort=[("Active", "desc"), {"column": "Added", "desc": True}])
    assert _skus(result) == ["SKU-3", "SKU-1", "SKU-5", "SKU-4", "SKU-2", "SKU-6"]
    text = table.query(["SKU"], sort=["Category", ("SKU", "desc")])
    assert _skus(text) == ["SKU-4", "SKU-5", "SKU-1", "SKU-2", "SKU-6", "SKU-3"]
    timed = table.query(["SKU"], where={"Active": {"eq": True}}, sort=[("At", "desc")])
    assert _skus(timed) == ["SKU-3", "SKU-1", "SKU-5"]       # ties keep the table's order


def test_limit_and_offset_give_a_page_of_all_the_matches(table):
    result = table.query(["SKU"], sort=[("Share", "desc")], limit=2, offset=1)
    assert _skus(result) == ["SKU-5", "SKU-2"] and result.matches == 6
    assert result.positions == (5, 2)
    assert table.query(limit=0).rows == () and table.query(limit=0).matches == 6
    assert table.query(offset=10).rows == ()


def test_pages_read_again_past_the_kept_cells_are_the_same(table, monkeypatch):
    arguments = dict(columns=["SKU", "Added"], where={"Price": {"not_blank": True}},
                     sort=[("Added", "desc")])
    kept = table.query(**arguments)
    assert matches(table.workbook, table, where=arguments["where"],
                   columns=["SKU"]).values is not None
    monkeypatch.setattr("xlsx_agent.query.KEEP_CELLS", 3)
    assert matches(table.workbook, table, where=arguments["where"],
                   columns=["SKU"]).values is None
    again = table.query(**arguments)
    assert again == kept and _skus(again) == ["SKU-3", "SKU-2", "SKU-1", "SKU-6"]


# -- the library: dates as the reader reads them ----------------------------------------------


def test_the_1904_date_system_and_iso_cells_compare_as_the_reader_shows_them():
    wb = _open(goods(date1904=True))
    added = wb.table("t:Goods").query(["Added"], where={"Added": {"eq": "2026-01-05"}})
    assert [row[0].v for row in added.rows] == ["2026-01-05", "2026-01-05"]
    serial = added.rows[0][0].serial
    assert serial == day_serial(2026, 1, 5, True)
    assert wb.table("t:Goods").query(where={"Added": {"eq": serial}}).matches == 2

    book = WorkbookBuilder()
    sheet = book.sheet("Log")
    sheet.rows("A1", [["When", "Clock"], [IsoDate(date(2026, 1, 5)), IsoDate(time(9, 30))],
                      [IsoDate(date(2026, 1, 7)), IsoDate(time(18, 0))]])
    sheet.table("Log", "A1:B3", ["When", "Clock"], write_header=False)
    log = _open(book).table("t:Log")
    assert log.query(where={"When": {"gte": "2026-01-06"}}).positions == (2,)
    assert log.query(where={"When": {"lte": 46027}}).positions == (1,)
    assert log.query(where={"Clock": {"lte": "12:00"}}).positions == (1,)
    assert log.query(["When"], sort=[("When", "desc")]).positions == (2, 1)


def test_day_serial_is_serial_days_inverse():
    for system in (False, True):
        for serial in (1, 59, 60, 61, 1000, 46027, 2958465 - (1462 if system else 0)):
            text = serial_day(serial, system)
            year, month, day = (int(p) for p in text.split("-"))
            assert day_serial(year, month, day, system) == serial, (serial, system)
    assert day_serial(1900, 2, 29, True) is None and day_serial(2026, 2, 30, False) is None
    assert day_serial(1899, 12, 31, False) is None and day_serial(1903, 12, 31, True) is None


def test_an_elapsed_time_column_compares_lengths():
    book = WorkbookBuilder()
    span = book.style(num_fmt="[h]:mm:ss")
    sheet = book.sheet("Jobs")
    sheet.rows("A1", [["Job", "Took"]])
    for r, (job, took) in enumerate([("J-1", 1.5), ("J-2", 0.25), ("J-3", 2.0)], start=2):
        sheet.cell(f"A{r}", job)
        sheet.cell(f"B{r}", took, style=span)
    sheet.table("Jobs", "A1:B4", ["Job", "Took"], write_header=False)
    jobs = _open(book).table("t:Jobs")
    assert jobs.query(["Took"]).rows[0][0].v == "PT36H0M0S"
    assert jobs.query(where={"Took": {"gte": "PT36H"}}).positions == (1, 3)
    assert jobs.query(where={"Took": {"lte": 1.5}}).positions == (1, 2)


def test_an_error_column_is_matched_by_code():
    book = WorkbookBuilder()
    sheet = book.sheet("Checks")
    sheet.rows("A1", [["Check"], [Error("#N/A")], [Error("#REF!")], [Error("#N/A")], [1]])
    sheet.table("Checks", "A1:A5", ["Check"], write_header=False)
    checks = _open(book).table("t:Checks")
    assert checks.query(where={"Check": {"eq": "#N/A"}}).positions == (1, 3)
    assert checks.query(where={"Check": {"in": ["#REF!"]}}).positions == (2,)
    with pytest.raises(QueryError):
        checks.query(where={"Check": {"contains": "N/A"}})


def test_a_region_without_data_rows_matches_nothing():
    book = WorkbookBuilder()
    book.sheet("Empty").rows("A1", [["SKU", "Price"]])
    wb = _open(book)
    region = wb.declare_region("Empty", "Empty!A1:B1")
    found = matches(wb, region, where={"Price": {"gte": 1}})
    assert found.rows == () and region.query().matches == 0


# -- xlsx_query through a Toolbox -------------------------------------------------------------


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(toolbox, book: "WorkbookBuilder | bytes", name: str = "book.xlsx"):
    session = toolbox.session(clock=CLOCK)
    data = book if isinstance(book, bytes) else book.to_bytes()
    blob = session.add_blob(data, name=name)
    assert toolbox.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def _query(toolbox, session, **arguments):
    return toolbox.dispatch(session, "xlsx_query", {"doc": "d1", **arguments})


def _cells(result) -> list[dict]:
    return [cell for row in result.data["rows"] for cell in row]


def test_a_query_end_to_end(box):
    session = _session(box, goods())
    result = _query(box, session, table="t:Goods", columns=["SKU", "Price"],
                    where=[{"column": "Price", "op": "gte", "value": "3"},
                           {"column": "Category", "op": "contains", "value": "TOOL"}],
                    sort=[{"column": "SKU", "desc": True}])
    assert result.ok, result.error
    assert result.total == 2 and result.next_cursor is None
    assert result.data == {
        "target": "t:Goods", "sheet": {"name": "Stock", "id": 1}, "range": "Stock!A1:I7",
        "columns": ["SKU", "Price"],
        "types": {"Price": "number", "Category": "text", "SKU": "text"},
        "not_compared": {"Price": {"count": 1, "first": ["Stock!C3"]}},
        "rows": [[{"a": "Stock!A7", "v": "SKU-6", "t": "text"},
                  {"a": "Stock!C7", "v": 3, "t": "number"}],
                 [{"a": "Stock!A4", "v": "SKU-3", "t": "text"},
                  {"a": "Stock!C4", "v": 3, "t": "number"}]]}
    assert result.summary == "t:Goods: matching rows 1-2 of 2 (of 6)"
    none = _query(box, session, table="t:Goods", where=[{"column": "SKU", "op": "eq",
                                                          "value": "SKU-9"}])
    assert none.ok and none.total == 0 and none.data["rows"] == []
    assert none.summary == "t:Goods: no rows of 6 match"


@pytest.mark.parametrize("item,expected", [
    ({"column": "Category", "op": "eq", "value": "Hardware"}, ["Stock!A2"]),
    ({"column": "Category", "op": "in", "values": ["Tools", "Garden"]},
     ["Stock!A4", "Stock!A5", "Stock!A7"]),
    ({"column": "Price", "op": "gte", "value": "12.5"}, ["Stock!A2"]),
    ({"column": "Added", "op": "lte", "value": "2026-01-05"},
     ["Stock!A2", "Stock!A6", "Stock!A7"]),
    ({"column": "Category", "op": "contains", "value": "hand"}, ["Stock!A6"]),
    ({"column": "Price", "op": "blank"}, ["Stock!A5", "Stock!A6"]),
    ({"column": "Price", "op": "not_blank"}, ["Stock!A2", "Stock!A3", "Stock!A4", "Stock!A7"]),
    ({"column": "Active", "op": "eq", "value": "true"}, ["Stock!A2", "Stock!A4", "Stock!A6"]),
], ids=lambda value: value["op"] if isinstance(value, dict) else None)
def test_each_filter_through_the_tool(box, item, expected):
    session = _session(box, goods())
    result = _query(box, session, table="t:Goods", columns=["SKU"], where=[item])
    assert result.ok, result.error
    assert [c["a"] for c in _cells(result)] == expected and result.total == len(expected)


def test_a_wrong_type_filter_is_refused_through_the_tool(box):
    session = _session(box, goods())
    result = _query(box, session, table="t:Goods",
                    where=[{"column": "Price", "op": "gte", "value": "cheap"}])
    assert not result.ok
    error = result.error
    assert error.code == "refused" and error.field == "where"
    assert error.details == {"code": "wrong_type", "column": "Price", "column_type": "number"}
    assert error.valid_options == ["eq", "in", "gte", "lte", "blank", "not_blank"]
    assert "number column" in error.message and "cheap" not in error.message
    ordered = _query(box, session, table="t:Goods",
                     where=[{"column": "Category", "op": "lte", "value": "M"}])
    assert ordered.error.code == "refused" and "contains" in ordered.error.valid_options


@pytest.mark.parametrize("arguments,code", [
    ({"table": "t:Goods[Price]"}, "invalid_arguments"),
    ({"table": "t:Goods/#1"}, "invalid_arguments"),
    ({"table": "Stock!A1:I7"}, "invalid_arguments"),
    ({"table": "t:Nope"}, "not_found"),
    ({"table": "t:Goods", "columns": ["Nope"]}, "not_found"),
    ({"table": "t:Goods", "where": [{"column": "Nope", "op": "blank"}]}, "not_found"),
    ({"table": "t:Goods", "where": [{"column": "Price", "op": "between", "value": "1"}]},
     "invalid_arguments"),
    ({"table": "t:Goods", "where": [{"column": "Price", "op": "in"}]}, "invalid_arguments"),
    ({"table": "t:Goods", "where": [{"column": "Price", "op": "eq"}]}, "invalid_arguments"),
    ({"table": "t:Goods", "sort": [{"column": "Nope"}]}, "not_found"),
    ({"table": "t:Goods", "limit": READ_ROWS + 1}, "invalid_arguments"),
    ({"table": "t:Goods", "cursor": "9.0.00000000"}, "invalid_arguments"),
], ids=["column", "row", "range", "table", "columns", "where-column", "op", "in-values",
        "eq-value", "sort-column", "limit", "cursor"])
def test_query_errors_use_the_tool_error_mapping(box, arguments, code):
    session = _session(box, goods())
    result = _query(box, session, **arguments)
    assert not result.ok and result.error.code == code, result.error


# -- large fixtures: sort, limit and paging ---------------------------------------------------

LONG_ROWS = 20_000


@pytest.fixture(scope="module")
def ledger() -> bytes:
    """A table of :data:`LONG_ROWS` rows: an id, a region, an amount (each tenth row's
    stored as text, each 97th blank) and a date."""
    book = WorkbookBuilder()
    sheet = book.sheet("Ledger")
    regions = ["North", "South", "East", "West"]
    rows = [["Id", "Region", "Amount", "Booked"]]
    for r in range(1, LONG_ROWS + 1):
        amount = None if r % 97 == 0 else (f"{r % 1000}" if r % 10 == 0 else (r * 37) % 1000)
        rows.append([f"L-{r:05d}", regions[r % 4], amount,
                     date(2025, 1, 1) + dt.timedelta(days=r % 365)])
    sheet.rows("A1", rows)
    sheet.table("Ledger", f"A1:D{LONG_ROWS + 1}", rows[0], write_header=False)
    return book.to_bytes()


def _expected_ledger() -> list[tuple[str, int]]:
    """The North rows with an amount of 500 or more, by amount descending, ties by id."""
    out = []
    for r in range(1, LONG_ROWS + 1):
        if r % 4 != 0 or r % 97 == 0 or r % 10 == 0:
            continue
        amount = (r * 37) % 1000
        if amount >= 500:
            out.append((f"L-{r:05d}", amount))
    out.sort(key=lambda item: -item[1])
    return out


def _walk(toolbox, session, **arguments) -> list:
    results, cursor = [], None
    while True:
        result = _query(toolbox, session, **arguments, **({"cursor": cursor} if cursor else {}))
        assert result.ok, result.error
        results.append(result)
        cursor = result.next_cursor
        if cursor is None:
            return results


def test_sort_limit_and_paging_on_a_large_table(box, ledger):
    session = _session(box, ledger)
    arguments = dict(table="t:Ledger", columns=["Id", "Amount"],
                     where=[{"column": "Region", "op": "eq", "value": "North"},
                            {"column": "Amount", "op": "gte", "value": "500"}],
                     sort=[{"column": "Amount", "desc": True}])
    expected = _expected_ledger()
    first = _query(box, session, **arguments, limit=10)
    assert first.total == len(expected) and len(first.data["rows"]) == 10
    assert [(r[0]["v"], r[1]["v"]) for r in first.data["rows"]] == expected[:10]
    assert first.data["cut"] == {"rows": len(expected) - 10}
    assert first.summary == f"t:Ledger: matching rows 1-10 of {len(expected)} (of {LONG_ROWS}); " \
        "more with next_cursor"
    assert first.data["not_compared"]["Amount"]["count"] == \
        sum(1 for r in range(1, LONG_ROWS + 1) if r % 10 == 0 and r % 97)

    results = _walk(box, session, **arguments, limit=READ_ROWS)
    cells = [c["v"] for res in results for c in _cells(res)]
    assert list(zip(cells[::2], cells[1::2])) == expected and len(results) > 1
    for result in results:
        assert len(result.to_text()) <= box.limits.max_result_chars
        assert result.total == len(expected)
    # a cursor belongs to its query
    other = _query(box, session, **{**arguments, "sort": [{"column": "Id"}]},
                   cursor=first.next_cursor)
    assert other.error.code == "invalid_arguments" and other.error.field == "cursor"
    assert "North" not in other.error.message
    nothing = _query(box, session, table="t:Ledger", columns=["Id"], limit=READ_ROWS)
    shown = len(nothing.data["rows"])
    assert nothing.total == LONG_ROWS and 400 < shown <= READ_ROWS
    assert nothing.data["cut"]["rows"] == LONG_ROWS - shown
    assert len(nothing.to_text()) <= box.limits.max_result_chars


def test_ascending_dates_and_blanks_last_across_pages(box, ledger):
    session = _session(box, ledger)
    results = _walk(box, session, table="t:Ledger", columns=["Booked", "Amount"],
                    where=[{"column": "Booked", "op": "lte", "value": "2025-01-03"}],
                    sort=[{"column": "Amount"}], limit=40)
    rows = [r for res in results for r in res.data["rows"]]
    days = {"2025-01-01", "2025-01-02", "2025-01-03"}
    assert {r[0]["v"] for r in rows} == days
    kinds = [r[1]["t"] for r in rows]
    numbers = [r[1]["v"] for r in rows if r[1]["t"] == "number"]
    assert numbers == sorted(numbers)
    assert kinds == sorted(kinds, key=["number", "text", "blank"].index)
    assert results[0].total == len(rows) == sum(
        1 for r in range(1, LONG_ROWS + 1) if r % 365 in (0, 1, 2))


@pytest.mark.parametrize("limits", [Limits(), Limits(max_result_chars=3000, max_list_items=20)],
                         ids=["default", "tight"])
def test_a_wide_query_is_bounded_split_and_resumed(large_book, limits):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, limits=limits) as toolbox:
        session = _session(toolbox, large_book)
        results = _walk(toolbox, session, table="t:Products",
                        where=[{"column": "Products column 2", "op": "eq", "value": "Tools"}],
                        sort=[{"column": "Products column 3", "desc": True}], limit=READ_ROWS)
        addresses = [c["a"] for res in results for c in _cells(res)]
        rows = [r for r in range(1, ROWS + 1) if r % 6 == 1]
        rows.sort(key=lambda r: -(round(r * 1.25, 2)))
        assert addresses == [f"Stock!{chr(65 + k)}{r + 1}" for r in rows for k in range(WIDE)]
        for result in results:
            assert len(result.to_text()) <= limits.max_result_chars
            if result.next_cursor is not None:
                assert result.data["cut"]["chars"] is True and result.data["cut"]["rows"] > 0
        assert len(results) > 1
        if limits != Limits():
            assert any("cells" in r.data.get("cut", {}) for r in results)


@pytest.fixture(scope="module")
def large_book() -> bytes:
    return warehouse().to_bytes()


# -- content stays out of diagnostics ----------------------------------------------------------


def test_no_cell_content_or_filter_value_reaches_a_diagnostic(box, caplog):
    caplog.set_level(logging.DEBUG)
    session = _session(box, goods())
    results = [
        _query(box, session, table="t:Goods", where=[
            {"column": "Category", "op": "eq", "value": f"Hand tools {MARKER}"}]),
        _query(box, session, table="t:Goods", where=[
            {"column": "Category", "op": "contains", "value": MARKER}], sort=[{"column": "SKU"}]),
        _query(box, session, table="t:Goods", where=[
            {"column": "Price", "op": "gte", "value": MARKER}]),
        _query(box, session, table="t:Goods", where=[
            {"column": "Note", "op": "in", "values": [MARKER]}]),
        _query(box, session, table="t:Goods", columns=["Category"], limit=1),
    ]
    assert results[0].ok and results[0].total == 1 and MARKER in json.dumps(results[0].data)
    second = _query(box, session, table="t:Goods", columns=["Category"], limit=1,
                    where=[{"column": "SKU", "op": "eq", "value": MARKER}],
                    cursor=results[4].next_cursor)
    results.append(second)
    diagnostics = [r.getMessage() for r in caplog.records]
    diagnostics += [w for r in results for w in r.warnings] + [r.summary for r in results]
    diagnostics += [json.dumps(r.error.to_json()) for r in results if not r.ok]
    diagnostics += [json.dumps(record.__dict__, default=str) for record in session.log]
    assert len([r for r in results if not r.ok]) == 3
    assert diagnostics and not [d for d in diagnostics if MARKER in d]
