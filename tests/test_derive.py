"""R10, the derive engine (library only): ``where``, ``group_by``, every measure,
``round``, provenance and exclusions; then ``bins``, ``pivot``, ``sort`` and ``top``.

The provenance test runs on every fixture of the suite: for every value of every derived
result it re-reads the value's source cells, recomputes the measure with a plain reference
implementation (fractions and :mod:`statistics`), and checks that the cells of the group's
rows left out of it are exactly those its reasons count.  Every workbook is built with
tests/build.py.
"""

import importlib
import math
import statistics
import time
import tracemalloc
from datetime import date, datetime
from datetime import time as clock
from fractions import Fraction

import pytest

from build import NS_MAIN, NS_R, Error, Formula, WorkbookBuilder, zip_parts
from test_catalogue import FIXTURES, MARKER

from xlsx_agent import (DeriveError, Locale, NameNotFound, QueryError, Region, RegionError,
                        Workbook)
from xlsx_agent import derive as derive_module
from xlsx_agent.derive import FUNCTIONS, TYPES_BY_FUNCTION, _round
from xlsx_agent.address import TableAddress
from xlsx_agent.query import _key
from xlsx_agent.read.types import serial_text, text_date, text_number

COLUMNS = ["Region", "Rep", "Amount", "Rate", "Day", "At", "Opens", "Paid", "Note"]


def sales(date1904: bool = False) -> WorkbookBuilder:
    """A sales table with a column of every measured type and every reason a cell is left
    out: a number stored as text, a bool and an error among the amounts, a blank, a formula
    without a cached value, a hidden row; amounts whose sums show exactness and rounding."""
    book = WorkbookBuilder(date1904=date1904)
    pct = book.style(num_fmt="0%")
    sheet = book.sheet("Sales")
    sheet.rows("A1", [COLUMNS])
    rows = [
        ["North", "Ann", 2.675, 0.1, date(2026, 1, 5), datetime(2026, 1, 5, 9, 30),
         clock(9, 0), True, "first"],
        ["North", "Bob", 1e16, 0.2, date(2026, 1, 6), datetime(2026, 1, 6, 12, 0),
         clock(10, 30), False, f"marked {MARKER}"],
        ["South", "Cy", "7.25", 0.25, date(2026, 2, 1), datetime(2026, 2, 1, 0, 0),
         clock(12, 0), True, None],
        ["North", "Ann", 1.0, 0.3, date(2026, 1, 9), datetime(2026, 1, 9, 8, 15),
         clock(8, 0), True, "hidden row"],
        ["South", "Dee", None, 0.5, date(2026, 3, 1), datetime(2026, 3, 1, 23, 59, 59),
         clock(17, 45), False, None],
        ["North", "Bob", -1e16, 0.05, date(2025, 12, 31), datetime(2025, 12, 31, 8, 0),
         clock(8, 0), True, None],
        ["East", "Eve", Formula("C2*2"), 0.75, date(2026, 1, 5), datetime(2026, 1, 5, 9, 30),
         clock(9, 0), False, None],
        ["East", "Eve", Error("#N/A"), 1, date(2026, 2, 10), datetime(2026, 2, 10, 10, 0),
         clock(9, 15), True, None],
        ["South", "Cy", True, 0.125, date(2026, 2, 2), datetime(2026, 2, 2, 14, 0),
         clock(11, 0), True, None],
        ["West", "Fay", 2.5, 0.375, date(2026, 4, 1), datetime(2026, 4, 1, 9, 0),
         clock(9, 0), False, None],
        ["West", "Fay", -2.5, 0.625, date(2026, 4, 2), datetime(2026, 4, 2, 9, 0),
         clock(9, 0), True, None],
        ["South", "Dee", 0.125, 0.875, date(2026, 2, 3), datetime(2026, 2, 3, 9, 0),
         clock(9, 30), False, None],
    ]
    for r, values in enumerate(rows, start=2):
        for c, value in enumerate(values):
            if value is None:
                continue
            sheet.cell(f"{chr(65 + c)}{r}", value, style=pct if c == 3 else 0)
    sheet.hide_row(5)
    sheet.table("Sales", f"A1:I{len(rows) + 1}", COLUMNS, write_header=False)
    return book


#: Every workbook the suite builds: the catalogue's list, its own, the query's and these.
EVERY = FIXTURES + [("test_catalogue", "shop"), ("test_catalogue", "warehouse"),
                    ("test_query", "goods"), ("test_derive", "sales")]


def _open(book: WorkbookBuilder, **options) -> Workbook:
    return Workbook.open(book.to_bytes(), **options)


@pytest.fixture(scope="module")
def wb() -> Workbook:
    return _open(sales())


# -- the reference ----------------------------------------------------------------------------


_NUMBERS = ("number", "percent")
_MOMENTS = ("date", "time", "datetime")


def _cells(workbook: Workbook, sources, table: "dict | None" = None, first: int = 1) -> list:
    """The cells of ``sources``: read again from the workbook, or looked up in ``table`` (the
    region's rows, read once, by sheet row) when given."""
    if table is None:
        return [v for source in sources for row in workbook.read(source) for v in row]
    out = []
    for area in (a for source in sources for a in _areas(workbook, source)):
        for number in range(area.first.row, area.last.row + 1):
            row = table[number]
            for column in range(area.first.column, area.last.column + 1):
                cell = row[column - first]
                assert (cell.row, cell.column) == (number, column)
                out.append(cell)
    return out


def _areas(workbook: Workbook, source) -> list:
    """The areas a source names: a range's own, or a key set's rows, one by one."""
    if isinstance(source, TableAddress):
        return [a.area for a in workbook.resolve(source).areas]
    return [source.area]


def _reason(value, t: str) -> str:
    if value.issue == "no_cached_value":
        return "no_cached_value"
    if value.t == "blank":
        return "blank"
    if value.t == "error":
        return "error"
    if value.issue == "number_as_text" and t in _NUMBERS:
        return "number_as_text"
    if value.issue == "date_as_text" and t in _MOMENTS:
        return "date_as_text"
    return "wrong_type"


def _moment(exact: Fraction, t: str, cells, date1904: bool) -> str:
    serial = float(exact) if t == "date" else float(exact / 86_400_000)
    elapsed = t == "time" and all(str(c.v).startswith("PT") for c in cells)
    return serial_text(serial, t, date1904, elapsed)


def _reference(fn: str, t: str, cells, date1904: bool):
    """The measure recomputed from its cells, plainly."""
    if fn == "count":
        return len(cells)
    if fn == "distinct":
        if t in _MOMENTS:
            return len({_key(c, t, date1904) for c in cells})
        return len({c.v for c in cells})
    if not cells:
        return None
    if fn in ("first", "last"):
        return cells[0 if fn == "first" else -1].v
    if t in _NUMBERS:
        exact = [Fraction(c.v) for c in cells]
        if fn == "sum":
            return float(sum(exact))
        if fn == "mean":
            return float(sum(exact) / len(exact))
        if fn == "median":
            return float(statistics.median(exact))
        pick = min if fn == "min" else max
        return pick(cells, key=lambda c: c.v).v
    keys = [Fraction(_key(c, t, date1904)) for c in cells]
    if fn == "mean":
        return _moment(sum(keys) / len(keys), t, cells, date1904)
    if fn == "median":
        return _moment(statistics.median(keys), t, cells, date1904)
    pick = min if fn == "min" else max
    return pick(cells, key=lambda c: _key(c, t, date1904)).v


def _measures(types: dict[str, str]) -> list[dict]:
    """Every measure each column's type takes, and the counts and shares of rows."""
    out = [{"fn": "count", "as": "rows"}, {"fn": "share", "as": "share of rows"}]
    for column, t in types.items():
        for fn in FUNCTIONS:
            if t in TYPES_BY_FUNCTION[fn]:
                out.append({"fn": fn, "column": column, "as": f"{fn}~{column}"})
    return out


def check_provenance(workbook: Workbook, region: Region, ds, measures: list[dict],
                     types: dict[str, str], *, once: bool = False) -> int:
    """Re-read every value's sources and recompute it; check what was left out.  The
    number of values checked.  ``once``: read the region once and look each source up in
    it (a large fixture), instead of reading each source from the workbook."""
    date1904 = workbook.date1904
    data = region.data
    first = region.area.first.column
    table = {row[0].row: row for row in workbook.read(region.range.on(region.sheet_ref))} \
        if data is not None else {}
    keys = ds.columns[:ds.key_columns]
    checked = 0
    sheet = region.sheet_ref.text
    for where in ds.provenance:
        assert all(s.whole.text == region.address.text if isinstance(s, TableAddress)
                   else s.sheet.text == sheet for p in where.values() for s in p.sources)
    lookup = (table, first) if once else (None, 1)

    def read(sources) -> list:
        return _cells(workbook, sources, *lookup)

    for row, where in zip(ds.unrounded, ds.provenance):
        if keys:
            group = read(where[keys[0]].sources)
            assert len(group) == where[keys[0]].count
            assert all(c.v == row[keys[0]] for c in group)
            rows = {c.row for c in group}
            checked += 1
        else:
            rows = set(range(data.first.row, data.last.row + 1)) if data is not None else set()
        for m in measures:
            name, fn, column = m["as"], m["fn"], m.get("column")
            value, prov = row[name], where[name]
            cells = read(prov.sources)
            assert len(cells) == prov.count
            used = {c.row for c in cells}
            assert used <= rows and len(used) == len(cells)
            checked += 1
            if column is None:
                assert used == rows
                if fn == "count":
                    assert value == len(rows)
                else:
                    assert value == (float(Fraction(len(rows), ds.used)) if rows else None)
                    assert ds.totals[name].count == ds.used
                continue
            t = types[column]
            assert all(c.t == t for c in cells), (name, [c.a for c in cells if c.t != t])
            # the group's rows left out of it, by reason, exactly
            offset = region.column_index(column)
            reasons: dict[str, int] = {}
            for number in rows - used:
                cell = table[number][offset]
                assert cell.t != t or t == "blank"
                reasons[_reason(cell, t)] = reasons.get(_reason(cell, t), 0) + 1
            assert reasons == prov.excluded, name
            if fn == "share":
                total = read(ds.totals[name].sources)
                whole = sum(Fraction(c.v) for c in total)
                expected = float(sum(Fraction(c.v) for c in cells) / whole) \
                    if cells and whole else None
                assert ds.totals[name].value == float(whole)
            else:
                expected = _reference(fn, t, cells, date1904)
            assert value == expected, (name, value, expected)
            assert type(value) is type(expected) or value is None
    return checked


def _regions(workbook: Workbook) -> list[Region]:
    found: list[Region] = list(workbook.tables)
    for name in workbook.names:
        if name.is_range and name.address is not None and not name.builtin:
            try:
                found.append(workbook.region(name.address))
            except (RegionError, NameNotFound):
                pass
    for sheet in workbook.sheets:
        if sheet.type == "worksheet":
            found.extend(r for r in workbook.find_regions(sheet.name) if isinstance(r, Region))
    return found


@pytest.mark.parametrize("module,name", EVERY, ids=[f"{m}.{n}" for m, n in EVERY])
def test_every_value_of_every_fixture_is_rederived_from_its_sources(module, name):
    book = getattr(importlib.import_module(module), name)()
    workbook = _open(book)
    checked = 0
    for region in _regions(workbook):
        stats = region.profile()
        types = {s.column: s.type for s in stats}
        measures = _measures(types)
        groups = [s.column for s in stats if s.values is not None and s.type != "blank"][:1]
        for group_by in ([], groups):
            for start in range(0, len(measures), 40):
                chunk = measures[start:start + 40]
                if not any(m["as"] == "rows" for m in chunk):
                    chunk = [measures[0]] + chunk
                ds = workbook.derive(region, group_by=group_by, measures=chunk)
                assert ds.matched == region.rows
                assert ds.used + sum(c.count for reasons in ds.left_out.values()
                                     for c in reasons.values()) == ds.matched
                assert sum(row["rows"] for row in ds.rows) == ds.used
                checked += check_provenance(workbook, region, ds, chunk, types,
                                            once=region.rows > 50)
    assert checked or not _regions(workbook)


@pytest.mark.parametrize("date1904", [False, True])
def test_sales_by_region_is_rederived_with_where_bins_and_pivot(date1904):
    workbook = _open(sales(date1904))
    region = workbook.table("t:Sales")
    types = {s.column: s.type for s in region.profile()}
    measures = _measures({c: types[c] for c in ("Amount", "Rate", "Day", "At", "Opens")})
    for options in ({"group_by": ["Region", "Rep"]},
                    {"group_by": ["Region"], "where": {"Paid": {"eq": True}}},
                    {"group_by": ["Region"], "where": {"Day": {"gte": "2026-01-06"}}},
                    {"group_by": ["Paid"], "where": [{"column": "Rate", "op": "lte",
                                                      "value": "0.5"}]}):
        ds = workbook.derive("t:Sales", measures=measures, **options)
        check_provenance(workbook, region, ds, measures, types)
    # bins: the bin column's cells of each row are in their bin
    ds = workbook.derive("t:Sales", group_by=["Region"], measures=measures,
                         bins={"column": "Rate", "edges": [0, 0.25, 0.5, 1]})
    check_provenance(workbook, region, ds, measures, types)
    for row, where in zip(ds.rows, ds.provenance):
        lo, hi = {"[0, 0.25)": (0, 0.25), "[0.25, 0.5)": (0.25, 0.5),
                  "[0.5, 1]": (0.5, 1.0000001)}[row["Rate"]]
        assert all(lo <= c.v < hi for c in _cells(workbook, where["Rate"].sources))


# -- measures ---------------------------------------------------------------------------------


def test_the_measures_of_a_group(wb):
    ds = wb.derive("t:Sales", group_by=["Region"], measures=[
        {"fn": "count", "as": "rows"},
        {"fn": "count", "column": "Amount", "as": "amounts"},
        {"fn": "sum", "column": "Amount", "as": "total"},
        {"fn": "mean", "column": "Amount", "as": "mean"},
        {"fn": "median", "column": "Amount", "as": "median"},
        {"fn": "min", "column": "Amount", "as": "low"},
        {"fn": "max", "column": "Amount", "as": "high"},
        {"fn": "distinct", "column": "Rep", "as": "reps"}])
    assert ds.columns == ("Region", "rows", "amounts", "total", "mean", "median", "low",
                          "high", "reps")
    assert ds.key_columns == 1
    assert ds.types == {"Region": "text", "rows": "number", "amounts": "number",
                        "total": "number", "mean": "number", "median": "number",
                        "low": "number", "high": "number", "reps": "number"}
    by = {row["Region"]: row for row in ds.rows}
    assert list(by) == ["North", "South", "East", "West"]            # as first met
    # North: 2.675 + 1e16 + 1.0 - 1e16, exactly 3.675 (a plain left-to-right sum gives 2)
    assert by["North"]["total"] == 3.675
    assert ((2.675 + 1e16) + 1.0) - 1e16 != 3.675
    assert by["North"]["mean"] == float(Fraction(3.675) / 4) or \
        by["North"]["mean"] == float((Fraction(2.675) + Fraction(1.0)) / 4)
    assert by["North"]["mean"] == float((Fraction(2.675) + 1) / 4)
    assert by["North"]["median"] == float((Fraction(1.0) + Fraction(2.675)) / 2)
    assert (by["North"]["low"], by["North"]["high"]) == (-1e16, 1e16)
    assert by["North"]["reps"] == 2
    # South: "7.25" as text, a blank, a bool: left out by reason; 0.125 the only number
    assert by["South"]["rows"] == 4 and by["South"]["amounts"] == 1
    assert by["South"]["total"] == by["South"]["mean"] == 0.125
    assert ds.provenance[1]["total"].excluded == {"blank": 1, "number_as_text": 1,
                                                  "wrong_type": 1}
    # East: no cached value and an error: no values, so no sum, but a count of 0
    assert by["East"]["amounts"] == 0
    assert by["East"]["total"] is None and by["East"]["median"] is None
    assert ds.provenance[2]["total"].excluded == {"no_cached_value": 1, "error": 1}
    assert ds.provenance[2]["total"].sources == ()
    # West: 2.5 and -2.5
    assert by["West"]["total"] == 0.0 and by["West"]["median"] == 0.0
    assert ds.excluded["total"]["number_as_text"].first == ("Sales!C4",)
    assert ds.excluded["total"]["error"].as_dict() == {"count": 1, "first": ["Sales!C9"]}
    assert set(ds.excluded["total"]) == {"blank", "no_cached_value", "error",
                                         "number_as_text", "wrong_type"}
    assert "rows" not in ds.excluded and "reps" not in ds.excluded


def test_sources_are_compact_ranges(wb):
    ds = wb.derive("t:Sales", group_by=["Region"],
                   measures=[{"fn": "max", "column": "Day", "as": "last"}])
    assert [s.text for s in ds.sources["North"]["Region"]] == ["Sales!A2:A3", "Sales!A5",
                                                               "Sales!A7"]
    assert [s.text for s in ds.sources["North"]["last"]] == ["Sales!E2:E3", "Sales!E5",
                                                             "Sales!E7"]
    assert ds.rows[0]["last"] == "2026-01-09"
    whole = wb.derive("t:Sales", measures=[{"fn": "count", "column": "Rate", "as": "n"}])
    assert whole.key_columns == 0 and whole.rows == [{"n": 12}]
    assert [s.text for s in whole.sources[None]["n"]] == ["Sales!D2:D13"]
    assert whole.provenance[0]["n"].as_dict() == {"sources": ["Sales!D2:D13"], "count": 12}


def test_dates_and_times_are_measured_as_queries_compare_them(wb):
    ds = wb.derive("t:Sales", group_by=["Region"], measures=[
        {"fn": "mean", "column": "Day", "as": "day"},
        {"fn": "median", "column": "At", "as": "at"},
        {"fn": "min", "column": "Opens", "as": "opens"},
        {"fn": "max", "column": "At", "as": "latest"},
        {"fn": "distinct", "column": "Opens", "as": "times"}])
    by = {row["Region"]: row for row in ds.rows}
    assert ds.types["day"] == "date" and ds.types["at"] == "datetime"
    assert ds.types["opens"] == "time"
    # West: 1 and 2 April: the mean is 1.5 April, shown as its day
    assert by["West"]["day"] == "2026-04-01"
    assert by["West"]["at"] == "2026-04-01T21:00:00"
    assert by["North"]["opens"] == "08:00:00" and by["North"]["times"] == 3
    assert by["South"]["latest"] == "2026-03-01T23:59:59"


def test_shares_are_fractions_of_the_stated_total(wb):
    ds = wb.derive("t:Sales", group_by=["Region"], measures=[
        {"fn": "share", "column": "Rate", "as": "of rates"},
        {"fn": "share", "as": "of rows"}])
    total = sum(Fraction(x) for x in (0.1, 0.2, 0.25, 0.3, 0.5, 0.05, 0.75, 1, 0.125,
                                      0.375, 0.625, 0.875))
    assert ds.totals["of rates"].value == float(total)
    assert [s.text for s in ds.totals["of rates"].sources] == ["Sales!D2:D13"]
    assert ds.totals["of rows"].value == 12 and ds.totals["of rows"].count == 12
    north = ds.rows[0]
    assert north["of rates"] == float((Fraction(0.1) + Fraction(0.2) + Fraction(0.3)
                                       + Fraction(0.05)) / total)
    assert north["of rows"] == 4 / 12
    assert ds.types["of rates"] == ds.types["of rows"] == "percent"
    assert math.isclose(sum(row["of rows"] for row in ds.rows), 1.0)


def test_no_rows_give_a_count_of_zero_and_no_other_value(wb):
    ds = wb.derive("t:Sales", where={"Region": {"eq": "Nowhere"}},
                   measures=[{"fn": "count", "as": "n"}, {"fn": "sum", "column": "Amount",
                                                            "as": "s"},
                             {"fn": "share", "as": "part"}])
    assert ds.rows == [{"n": 0, "s": None, "part": None}]
    assert ds.matched == ds.used == 0
    grouped = wb.derive("t:Sales", where={"Region": {"eq": "Nowhere"}}, group_by=["Region"])
    assert grouped.rows == [] and grouped.groups == 0
    assert grouped.columns == ("Region", "count")


def test_where_reuses_the_query_filters_and_counts_what_it_did_not_compare(wb):
    ds = wb.derive("t:Sales", where={"Amount": {"gte": 1}}, measures=[{"fn": "count"}])
    # 2.675, 1e16, 1.0, 2.5; text, bool, blank, error and no cached value never compared
    assert ds.rows == [{"count": 4}]
    assert ds.not_compared["Amount"].count == 3                 # text, bool, error
    assert ds.not_compared["Amount"].first == ("Sales!C4", "Sales!C9", "Sales!C10")
    with pytest.raises(QueryError) as caught:
        wb.derive("t:Sales", where={"Region": {"gte": 3}})
    assert caught.value.code == "wrong_type" and caught.value.column_type == "text"
    with pytest.raises(QueryError) as caught:
        wb.derive("t:Sales", where={"Amount": {"eq": "lots"}})
    assert caught.value.code == "wrong_type"
    with pytest.raises(QueryError) as caught:
        wb.derive("t:Sales", where={"Amount": {"near": 3}})
    assert caught.value.code == "bad_filter"


def test_group_keys_by_type_and_value_and_blank_as_a_group(wb):
    ds = wb.derive("t:Sales", group_by=["Note"], measures=[{"fn": "count", "as": "n"}])
    assert [(row["Note"], row["n"]) for row in ds.rows] == [
        ("first", 1), (f"marked {MARKER}", 1), (None, 9), ("hidden row", 1)]
    # a group of more than one column; a formula without a cached value has no group
    ds = wb.derive("t:Sales", group_by=["Amount"], measures=[{"fn": "count", "as": "n"}])
    assert ds.left_out == {"Amount": {"no_cached_value": ds.left_out["Amount"]
                                      ["no_cached_value"]}}
    assert ds.left_out["Amount"]["no_cached_value"].first == ("Sales!C8",)
    assert ds.used == 11 and ds.matched == 12
    keys = [row["Amount"] for row in ds.rows]
    assert True in keys and "7.25" in keys and "#N/A" in keys and None in keys
    multi = wb.derive("t:Sales", group_by=["Region", "Rep"])
    assert multi.key_columns == 2
    assert multi.rows[0] == {"Region": "North", "Rep": "Ann", "count": 2}
    assert [s.text for s in multi.sources[("North", "Ann")]["Rep"]] == ["Sales!B2", "Sales!B5"]


def test_hidden_rows_are_used_and_counted(wb):
    ds = wb.derive("t:Sales", measures=[{"fn": "count"}])
    assert ds.rows == [{"count": 12}]
    assert ds.hidden.count == 1 and ds.hidden.first[0].endswith("5")


def test_numbers_as_text_are_read_only_when_asked():
    workbook = _open(sales())
    measures = [{"fn": "sum", "column": "Amount", "as": "total"}]
    plain = workbook.derive("t:Sales", group_by=["Region"], measures=measures)
    assert plain.rows[1]["total"] == 0.125 and not plain.coerced
    ds = workbook.derive("t:Sales", group_by=["Region"], measures=measures,
                         coerce=["number_as_text"])
    assert ds.rows[1]["total"] == 7.375
    assert ds.coerced["total"].as_dict() == {"count": 1, "first": ["Sales!C4"]}
    assert ds.provenance[1]["total"].excluded == {"blank": 1, "wrong_type": 1}
    assert [s.text for s in ds.sources["South"]["total"]] == ["Sales!C4", "Sales!C13"]
    with pytest.raises(DeriveError) as caught:
        workbook.derive("t:Sales", coerce=["dates_as_text"])
    assert caught.value.code == "bad_coerce" and caught.value.valid_options == ["number_as_text",
                                                                     "date_as_text"]


def test_numbers_as_text_are_read_in_the_workbooks_locale():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["Item", "Qty"], ["a", 2], ["b", "1.234,5"], ["c", "12,5%"], ["d", 3]])
    sheet.table("Q", "A1:B5", ["Item", "Qty"], write_header=False)
    nl = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                language="nl-BE", list_separator=";")
    workbook = _open(book, locale=nl)
    ds = workbook.derive("t:Q", measures=[{"fn": "sum", "column": "Qty", "as": "s"}],
                         coerce=["number_as_text"])
    assert ds.rows == [{"s": float(5 + Fraction("1234.5") + Fraction("0.125"))}]
    assert text_number("1.234,5", nl) == 1234.5 and text_number("12,5%", nl) == 0.125
    assert text_number("x", nl) is None
    assert text_number("1,234.5") == 1234.5 and text_number("1,234.5", grouping=False) is None


def test_first_and_last_follow_the_row_order_for_any_type(wb):
    ends = [{"fn": fn, "column": column, "as": f"{fn} {column}"}
            for column in ("Amount", "Day", "Note", "Paid") for fn in ("first", "last")]
    ds = wb.derive("t:Sales", group_by=["Region"], measures=ends)
    north = ds.rows[0]
    assert ds.columns[0] == "Region" and north["Region"] == "North"
    # North is rows 2, 3, 5 and 7: first and last of each column's own type
    assert (north["first Amount"], north["last Amount"]) == (2.675, -1e16)
    assert (north["first Day"], north["last Day"]) == ("2026-01-05", "2025-12-31")
    assert (north["first Note"], north["last Note"]) == ("first", "hidden row")
    assert (north["first Paid"], north["last Paid"]) == (True, True)
    assert ds.types["first Day"] == "date" and ds.types["last Note"] == "text"
    # the sources are the cells compared, as min's: the first is the first of them
    note = ds.provenance[0]["first Note"]
    assert [s.text for s in note.sources] == ["Sales!I2:I3", "Sales!I5"]
    assert note.excluded == {"blank": 1}
    # South's amounts: text, a bool and a blank left out; with coerce, the text as read
    south = wb.derive("t:Sales", group_by=["Region"], measures=ends[:2],
                      coerce=["number_as_text"]).rows[1]
    assert (south["first Amount"], south["last Amount"]) == (7.25, 0.125)
    whole = wb.derive("t:Sales", measures=[{"fn": "last", "column": "Note", "as": "n"}])
    assert whole.rows == [{"n": "hidden row"}]


def test_min_and_max_of_coerced_text_are_given_as_read(wb):
    ds = wb.derive("t:Sales", group_by=["Region"], coerce=["number_as_text"],
                   measures=[{"fn": "max", "column": "Amount", "as": "top"}])
    assert ds.rows[1] == {"Region": "South", "top": 7.25}       # not the text "7.25"


def checks(day_order: bool = False) -> WorkbookBuilder:
    """A table of checks whose dates are stored as text: ISO 8601 text in Seen (one with a
    time of day, one blank), text in the en-US day order in Due (one with a two-digit
    year), and a date column, Done, with one date as ISO text."""
    book = WorkbookBuilder()
    sheet = book.sheet("Log")
    sheet.rows("A1", [["Check", "Rating", "Seen", "Due", "Done"],
                      ["c1", 3.5, "2026-03-04", "03/05/2026", date(2026, 3, 9)],
                      ["c2", 4, "2026-02-11", "02/20/2026", "2026-03-01"],
                      ["c3", 2.5, None, "03/01/26", date(2026, 3, 12)],
                      ["c4", 5, "2026-03-20T09:30", "04/02/2026", date(2026, 3, 10)]])
    sheet.table("Checks", "A1:E5", ["Check", "Rating", "Seen", "Due", "Done"],
                write_header=False)
    return book


def test_a_column_of_dates_stored_as_text_is_refused_saying_why():
    workbook = _open(checks())
    with pytest.raises(DeriveError) as caught:
        workbook.derive("t:Checks", measures=[{"fn": "min", "column": "Seen"}])
    error = caught.value
    assert (error.code, error.column_type, error.issue) == ("wrong_type", "text", "date_as_text")
    assert error.cells.as_dict() == {"count": 3, "first": ["Log!C2", "Log!C3", "Log!C5"]}
    assert "3 of its values are dates stored as text (date_as_text, first Log!C2)" in str(error)
    assert 'coerce ["date_as_text"] reads them as dates' in str(error)
    assert "first" in error.valid_options and "last" in error.valid_options
    assert "2026-03-04" not in str(error)                     # no cell's value
    # text in the locale's day order is not read as a date without local_dates
    with pytest.raises(DeriveError) as local:
        workbook.derive("t:Checks", measures=[{"fn": "max", "column": "Due"}])
    assert local.value.issue == "local_date_text" and local.value.cells.count == 4
    assert "local_dates" in str(local.value) and "coerce" not in str(local.value)
    # numbers stored as text, the same way
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [["Qty"], ["1"], ["2"], [3]])
    with pytest.raises(DeriveError) as numbers:
        _open(book).derive("t:S!A1", measures=[{"fn": "sum", "column": "Qty"}])
    assert numbers.value.issue == "number_as_text"
    assert 'coerce ["number_as_text"] reads them as numbers' in str(numbers.value)
    # and a refusal of another kind names no issue
    with pytest.raises(DeriveError) as plain:
        workbook.derive("t:Checks", measures=[{"fn": "sum", "column": "Check"}])
    assert plain.value.issue is None and "stored as text" not in str(plain.value)


def test_dates_stored_as_text_are_read_as_dates_when_asked():
    workbook = _open(checks())
    measures = [{"fn": fn, "column": "Seen", "as": fn} for fn in ("min", "max", "first",
                                                                  "last", "mean")]
    ds = workbook.derive("t:Checks", measures=measures, coerce=["date_as_text"])
    assert ds.types["min"] == "date"
    # a date column compares days: 2026-03-20T09:30 is its day
    assert ds.rows == [{"min": "2026-02-11", "max": "2026-03-20", "first": "2026-03-04",
                        "last": "2026-03-20", "mean": "2026-03-02"}]
    assert ds.coerced["min"].as_dict() == {"count": 3, "first": ["Log!C2", "Log!C3", "Log!C5"]}
    assert ds.provenance[0]["min"].excluded == {"blank": 1}
    # without coerce, first and last take the text as it is
    plain = workbook.derive("t:Checks", measures=measures[2:4])
    assert plain.rows == [{"first": "2026-03-04", "last": "2026-03-20T09:30"}]
    assert plain.types["first"] == "text"


def test_a_date_column_counts_its_dates_as_text_apart_and_reads_them_when_asked():
    workbook = _open(checks())
    measures = [{"fn": "min", "column": "Done", "as": "earliest"}]
    plain = workbook.derive("t:Checks", measures=measures)
    assert plain.rows == [{"earliest": "2026-03-09"}]
    assert plain.excluded["earliest"]["date_as_text"].as_dict() == {"count": 1,
                                                                    "first": ["Log!E3"]}
    read = workbook.derive("t:Checks", measures=measures, coerce=["date_as_text"])
    assert read.rows == [{"earliest": "2026-03-01"}] and read.coerced["earliest"].count == 1


def test_dates_in_the_locales_day_order_are_read_with_local_dates():
    workbook = _open(checks(), local_dates=True)
    measures = [{"fn": "min", "column": "Due", "as": "lo"},
                {"fn": "max", "column": "Due", "as": "hi"}]
    with pytest.raises(DeriveError) as caught:
        workbook.derive("t:Checks", measures=measures)
    assert caught.value.issue == "date_as_text" and caught.value.cells.count == 4
    ds = workbook.derive("t:Checks", measures=measures, coerce=["date_as_text"])
    # en-US is month first; a two-digit year's century is not read: left out, counted
    assert ds.rows == [{"lo": "2026-02-20", "hi": "2026-04-02"}]
    assert ds.excluded["lo"]["date_as_text"].as_dict() == {"count": 1, "first": ["Log!D4"]}
    nl = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                language="nl-BE", list_separator=";")
    dmy = _open(checks(), local_dates=True, locale=nl)
    ds = dmy.derive("t:Checks", measures=measures[:1], coerce=["date_as_text"])
    # dmy: 03/05/2026 is 3 May and 04/02/2026 4 February; 02/20/2026 is no day there
    assert ds.rows == [{"lo": "2026-02-04"}]
    assert set(ds.excluded["lo"]) == {"date_as_text", "wrong_type"}
    assert ds.excluded["lo"]["wrong_type"].first == ("Log!D3",)
    assert text_date(" 2026-03-20 09:30 ") == "2026-03-20T09:30"
    assert text_date("3.5.2026", nl, local_dates=True) == "2026-05-03"
    assert text_date("3.5.2026", nl) is None and text_date("3.5.26", nl, local_dates=True) is None
    assert text_date("2026-02-30") is None and text_date("soon") is None


# -- refusals ---------------------------------------------------------------------------------


@pytest.mark.parametrize("measure,column_type,takes", [
    ({"fn": "sum", "column": "Region"}, "text", ["count", "first", "last", "distinct"]),
    ({"fn": "min", "column": "Note"}, "text", ["count", "first", "last", "distinct"]),
    ({"fn": "mean", "column": "Paid"}, "bool", ["count", "first", "last", "distinct"]),
    ({"fn": "sum", "column": "Day"}, "date",
     ["count", "mean", "median", "min", "max", "first", "last", "distinct"]),
    ({"fn": "share", "column": "Opens"}, "time",
     ["count", "mean", "median", "min", "max", "first", "last", "distinct"]),
])
def test_a_measure_its_columns_type_does_not_take_is_refused(wb, monkeypatch, measure,
                                                             column_type, takes):
    passes = []
    original = derive_module._second_pass
    monkeypatch.setattr(derive_module, "_second_pass",
                        lambda plan: passes.append(1) or original(plan))
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", group_by=["Region"],
                  measures=[{"fn": "count"}, dict(measure, **{"as": "x"})])
    error = caught.value
    assert (error.code, error.field, error.column, error.column_type) == (
        "wrong_type", "measures", measure["column"], column_type)
    assert error.valid_options == takes
    assert measure["column"] in str(error) and column_type in str(error)
    assert not passes                                   # refused before anything is computed
    assert isinstance(error, QueryError)


def test_a_measure_on_a_column_without_values_is_refused():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["Item", "Empty"], ["a"], ["b"]])
    sheet.table("E", "A1:B3", ["Item", "Empty"], write_header=False)
    workbook = _open(book)
    with pytest.raises(DeriveError, match="holds no values") as caught:
        workbook.derive("t:E", measures=[{"fn": "sum", "column": "Empty"}])
    assert caught.value.column_type == "blank"
    ds = workbook.derive("t:E", measures=[{"fn": "count", "column": "Empty", "as": "n"}])
    assert ds.rows == [{"n": 0}] and ds.provenance[0]["n"].excluded == {"blank": 2}


@pytest.mark.parametrize("options,code,field", [
    ({"measures": [{"fn": "total", "column": "Amount"}]}, "bad_measure", "measures"),
    ({"measures": [{"fn": "sum"}]}, "bad_measure", "measures"),
    ({"measures": [{"fn": "sum", "column": "Amount", "label": "x"}]}, "bad_measure",
     "measures"),
    ({"measures": []}, "bad_measure", "measures"),
    ({"measures": {"fn": "count"}}, "bad_measure", "measures"),
    ({"measures": [{"fn": "count", "as": "n"}, {"fn": "sum", "column": "Amount", "as": "N"}]},
     "duplicate_name", "measures"),
    ({"group_by": ["Region"], "measures": [{"fn": "count", "as": "region"}]},
     "duplicate_name", "measures"),
    ({"group_by": ["Region", "region"]}, "bad_group", "group_by"),
    ({"group_by": [3]}, "bad_group", "group_by"),
    ({"bins": {"column": "Amount"}}, "bad_bins", "bins"),
    ({"bins": {"column": "Amount", "edges": [1]}}, "bad_bins", "bins"),
    ({"bins": {"column": "Amount", "edges": [5, 1]}}, "bad_bins", "bins"),
    ({"bins": {"column": "Amount", "edges": [0, 1, 1]}}, "bad_bins", "bins"),
    ({"bins": {"column": "Region", "edges": [0, 1]}}, "wrong_type", "bins"),
    ({"bins": {"column": "Day", "edges": [0, "soon"]}}, "wrong_type", "bins"),
    ({"group_by": ["Region"], "bins": {"column": "Region", "edges": [0, 1]}}, "bad_bins",
     "bins"),
    ({"pivot": ["Region"]}, "bad_pivot", "pivot"),
    ({"group_by": ["Region"], "pivot": "Region"}, "bad_pivot", "pivot"),
    ({"sort": ["nothing"]}, "bad_sort", "sort"),
    ({"sort": [("count", "down")]}, "bad_sort", "sort"),
    ({"top": 0}, "bad_top", "top"),
    ({"top": True}, "bad_top", "top"),
    ({"round": 1.5}, "bad_round", "round"),
    ({"round": 16}, "bad_round", "round"),
    ({"round": {"count": 2}}, "bad_round", "round"),
    ({"id": "not an id"}, "bad_source", "id"),
])
def test_bad_arguments_are_refused(wb, options, code, field):
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", **options)
    assert (caught.value.code, caught.value.field) == (code, field)


def test_a_sort_by_an_input_column_names_the_output_columns_and_the_one_meant(wb):
    measures = [{"fn": "sum", "column": "Amount", "as": "total"}]
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", group_by=["Region"], measures=measures, sort=[("Amount", "desc")])
    assert caught.value.valid_options == ["Region", "total"]
    assert str(caught.value) == ("sort names an output column: one of Region, total; "
                                 "'Amount' is an input column: did you mean 'total', its "
                                 "measure?")
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", group_by=["Region"], measures=measures, sort=["totl"])
    assert str(caught.value).endswith("one of Region, total; did you mean 'total'?")
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", group_by=["Region"], measures=measures, sort=["Opens"])
    assert str(caught.value) == "sort names an output column: one of Region, total"


def test_a_sort_refusal_under_a_pivot_points_to_its_columns(wb):
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", pivot="Region", measures=[{"fn": "sum", "column": "Amount"}],
                  sort=["Amount"])
    assert "or a pivot's (see pivot_columns)" in str(caught.value)
    assert "input column" not in str(caught.value)


def test_sources_other_than_a_table_are_refused(wb):
    for source in ("d:by_region", "Sales!A1:C4", "t:Sales[Amount]", "t:Sales/#2", 42):
        with pytest.raises(DeriveError) as caught:
            wb.derive(source)
        assert caught.value.code == "bad_source" and caught.value.field == "from"
    with pytest.raises(NameNotFound):
        wb.derive("t:Nothing")
    with pytest.raises(NameNotFound):
        wb.derive("t:Sales", measures=[{"fn": "sum", "column": "Amont"}])


def test_no_cell_text_in_any_refusal(wb):
    """Errors name columns, functions, types and codes, never a value, a key or a filter's."""
    attempts = [
        {"where": {"Note": {"gte": MARKER}}},
        {"where": {"Amount": {"eq": MARKER}}},
        {"measures": [{"fn": "sum", "column": "Note", "as": MARKER}]},
        {"bins": {"column": "Day", "edges": [MARKER, "2026-01-01"]}},
        {"sort": [MARKER]},
        {"round": {MARKER: 2}},
        {"group_by": ["Note"], "pivot": "Region", "top": 0},
        {"group_by": ["Region"], "pivot": "Note", "sort": ["nothing"]},
        {"group_by": ["Region"], "pivot": "Note", "round": {"nothing": 1}},
    ]
    for options in attempts:
        with pytest.raises(QueryError) as caught:
            wb.derive("t:Sales", **options)
        assert MARKER not in str(caught.value)
        assert MARKER not in repr(caught.value.valid_options)


# -- limits ------------------------------------------------------------------------------------


@pytest.mark.parametrize("limit,value,options,code", [
    ("MAX_GROUPS", 3, {"group_by": ["Region"]}, "too_many_groups"),
    ("MAX_GROUPS", 5, {"group_by": ["Region"], "bins": {"column": "Rate",
                                                       "edges": [0, 0.5, 1]}},
     "too_many_groups"),
    ("MAX_PIVOT_VALUES", 3, {"group_by": ["Paid"], "pivot": "Region"}, "too_many_columns"),
    ("MAX_ROWS", 3, {"group_by": ["Region"]}, "too_many_rows"),
    ("MAX_CELLS", 7, {"group_by": ["Region"]}, "too_many_cells"),
    ("MAX_RUNS", 10, {"group_by": ["Region"]}, "too_many_sources"),
    ("MAX_KEPT", 5, {"group_by": ["Region"],
                     "measures": [{"fn": "median", "column": "Rate"}]}, "too_many_values"),
    ("MAX_KEPT", 5, {"measures": [{"fn": "distinct", "column": "Rep"}]}, "too_many_values"),
])
def test_the_limits_refuse_clearly(wb, monkeypatch, limit, value, options, code):
    monkeypatch.setattr(derive_module, limit, value)
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", **options)
    assert caught.value.code == code


def test_top_keeps_a_result_within_the_row_limit(wb, monkeypatch):
    monkeypatch.setattr(derive_module, "MAX_ROWS", 3)
    ds = wb.derive("t:Sales", group_by=["Region"], sort=[("count", "desc")], top=2)
    assert ds.groups == 4 and len(ds.rows) == 2


# -- rounding ----------------------------------------------------------------------------------


@pytest.mark.parametrize("x,digits,expected", [
    (2.5, 0, 3.0), (-2.5, 0, -3.0), (0.5, 0, 1.0), (1.5, 0, 2.0), (-0.5, 0, -1.0),
    (0.125, 2, 0.13), (-0.125, 2, -0.13), (2.675, 2, 2.68), (1.005, 2, 1.01),
    (0.1 + 0.2, 1, 0.3), (1234.5, -1, 1230.0), (1235.0, -1, 1240.0), (-1250.0, -2, -1300.0),
    (1e300, 2, 1e300), (5e-324, 3, 0.0), (-0.0001, 2, 0.0), (7.0, 15, 7.0),
])
def test_rounding_is_half_away_from_zero_from_the_shown_decimal(x, digits, expected):
    got = _round(x, digits)
    assert got == expected and str(got) == str(expected)


def test_rounding_applies_to_the_reported_values_only(wb):
    measures = [{"fn": "sum", "column": "Amount", "as": "total"},
                {"fn": "mean", "column": "Rate", "as": "rate"},
                {"fn": "share", "as": "part"},
                {"fn": "count", "as": "n"},
                {"fn": "min", "column": "Day", "as": "first"}]
    ds = wb.derive("t:Sales", group_by=["Region"], measures=measures, round=2)
    assert ds.rounding == {"rule": "half_away_from_zero",
                           "digits": {"total": 2, "rate": 2, "part": 2}}
    north, raw = ds.rows[0], ds.unrounded[0]
    assert raw["total"] == 3.675 and north["total"] == 3.68
    assert raw["part"] == 4 / 12 and north["part"] == 0.33
    assert north["n"] == raw["n"] == 4 and north["first"] == raw["first"] == "2025-12-31"
    assert ds.rows[3]["total"] == 0.0
    per_column = wb.derive("t:Sales", group_by=["Region"], measures=measures,
                           round={"Total": 1})
    assert per_column.rounding["digits"] == {"total": 1}
    assert per_column.rows[0]["total"] == 3.7 and per_column.rows[0]["rate"] == raw["rate"]
    data = ds.as_dict()
    assert data["rows"][0][1] == 3.68 and data["unrounded"][0][1] == 3.675


# -- bins ---------------------------------------------------------------------------------------


def test_bins_count_a_distribution(wb):
    ds = wb.derive("t:Sales", bins={"column": "Amount", "edges": [-5, 0, 2.5, 10],
                                    "as": "band"},
                   measures=[{"fn": "count", "as": "n"},
                             {"fn": "sum", "column": "Amount", "as": "total"}])
    assert ds.columns == ("band", "n", "total") and ds.key_columns == 1
    assert ds.types["band"] == "text"
    assert [(row["band"], row["n"]) for row in ds.rows] == [
        ("[-5, 0)", 1), ("[0, 2.5)", 2), ("[2.5, 10]", 2)]
    assert ds.rows[2]["total"] == 5.175
    # each lower edge in, each upper edge out but the last's; outside or not a number: left
    # out, by reason
    assert ds.left_out["Amount"]["outside_bins"].count == 2              # 1e16 and -1e16
    assert set(ds.left_out["Amount"]) == {"outside_bins", "number_as_text", "blank",
                                          "no_cached_value", "error", "wrong_type"}
    assert ds.used == 5 and ds.matched == 12
    assert [s.text for s in ds.sources["[2.5, 10]"]["band"]] == ["Sales!C2", "Sales!C11"]


def test_bins_of_dates_list_every_bin_for_every_group(wb):
    ds = wb.derive("t:Sales", group_by=["Region"],
                   bins={"column": "Day", "edges": ["2026-01-01", "2026-02-01", 46113]})
    assert ds.columns == ("Region", "Day", "count")
    labels = ["[2026-01-01, 2026-02-01)", "[2026-02-01, 2026-04-01]"]
    assert [(row["Region"], row["Day"], row["count"]) for row in ds.rows] == [
        ("North", labels[0], 3), ("North", labels[1], 0),
        ("South", labels[0], 0), ("South", labels[1], 4),
        ("East", labels[0], 1), ("East", labels[1], 1),
        ("West", labels[0], 0), ("West", labels[1], 1)]
    assert ds.left_out["Day"]["outside_bins"].first == ("Sales!E7", "Sales!E12")
    assert ds.provenance[1]["count"].sources == ()
    assert ds.groups == 8


def test_bins_of_numbers_as_text_with_coerce():
    workbook = _open(sales())
    ds = workbook.derive("t:Sales", bins={"column": "Amount", "edges": [0, 10]},
                         coerce=["number_as_text"])
    assert ds.rows == [{"Amount": "[0, 10]", "count": 5}]
    assert ds.coerced["Amount"].first == ("Sales!C4",)


# -- pivot ---------------------------------------------------------------------------------------


def test_pivot_spreads_measures_over_a_columns_values(wb):
    ds = wb.derive("t:Sales", group_by=["Region"], pivot="Paid",
                   measures=[{"fn": "count", "as": "n"}])
    assert ds.columns == ("Region", "true", "false")
    assert ds.pivot_columns == {"true": {"measure": "n", "v": True, "t": "bool"},
                                "false": {"measure": "n", "v": False, "t": "bool"}}
    assert ds.rows == [{"Region": "North", "true": 3, "false": 1},
                       {"Region": "South", "true": 2, "false": 2},
                       {"Region": "East", "true": 1, "false": 1},
                       {"Region": "West", "true": 1, "false": 1}]
    assert [s.text for s in ds.sources["North"]["true"]] == ["Sales!A2", "Sales!A5", "Sales!A7"]
    many = wb.derive("t:Sales", pivot="Region",
                     measures=[{"fn": "count", "as": "n"},
                               {"fn": "sum", "column": "Amount", "as": "total"}])
    assert many.columns == ("North n", "North total", "South n", "South total", "East n",
                            "East total", "West n", "West total")
    assert many.rows[0]["East total"] is None and many.rows[0]["East n"] == 2
    assert many.provenance[0]["East total"].excluded == {"no_cached_value": 1, "error": 1}
    rounded = wb.derive("t:Sales", pivot="Region", round={"total": 0},
                        measures=[{"fn": "count", "as": "n"},
                                  {"fn": "sum", "column": "Amount", "as": "total"}])
    assert rounded.rounding["digits"] == {"total": 0}
    assert rounded.rows[0]["North total"] == 4.0 and rounded.rows[0]["South total"] == 0.0
    assert rounded.unrounded[0]["North total"] == 3.675


def test_pivot_cells_without_rows_and_clashing_names():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["Store", "Code", "Qty"], ["N", 1, 2], ["N", "1", 3], ["S", None, 4],
                      ["S", "(blank)", 5], ["N", "Store", 6]])
    sheet.table("P", "A1:C6", ["Store", "Code", "Qty"], write_header=False)
    workbook = _open(book)
    ds = workbook.derive("t:P", group_by=["Store"], pivot="Code",
                         measures=[{"fn": "sum", "column": "Qty", "as": "q"}])
    assert ds.columns == ("Store", "1", "1 (text)", "(blank)", "(blank) (text)",
                          "Store (text)")
    assert ds.rows == [{"Store": "N", "1": 2.0, "1 (text)": 3.0, "(blank)": None,
                        "(blank) (text)": None, "Store (text)": 6.0},
                       {"Store": "S", "1": None, "1 (text)": None, "(blank)": 4.0,
                        "(blank) (text)": 5.0, "Store (text)": None}]
    assert ds.pivot_columns["(blank)"] == {"measure": "q", "v": None, "t": "blank"}
    with pytest.raises(DeriveError) as caught:
        workbook.derive("t:P", pivot="Code", measures=[{"fn": "count", "as": "n"},
                                                       {"fn": "sum", "column": "Qty",
                                                        "as": "N"}])
    assert caught.value.code == "duplicate_name"


# -- listed pivot values (round 27, X6) ------------------------------------------------------------


def test_pivot_values_give_one_column_per_listed_value_in_order(wb):
    measures = [{"fn": "count", "as": "n"}, {"fn": "sum", "column": "Amount", "as": "total"},
                {"fn": "distinct", "column": "Rep", "as": "reps"},
                {"fn": "mean", "column": "Rate", "as": "rate"},
                {"fn": "max", "column": "Day", "as": "last day"}]
    ds = wb.derive("t:Sales", group_by=["Paid"], pivot="Region",
                   pivot_values=["West", "Central", "North"], measures=measures)
    assert ds.columns[:6] == ("Paid", "West n", "West total", "West reps", "West rate",
                              "West last day")
    assert [c.split(" ")[0] for c in ds.columns[1:]] == ["West"] * 5 + ["Central"] * 5 \
        + ["North"] * 5
    # a value without rows: count and distinct 0, every other measure blank, never 0
    for row in ds.rows:
        assert (row["Central n"], row["Central reps"]) == (0, 0)
        assert row["Central total"] is row["Central rate"] is row["Central last day"] is None
        assert ds.provenance[ds.rows.index(row)]["Central total"].count == 0
    assert ds.pivot_columns["Central n"] == {"measure": "n", "v": "Central", "t": "text"}
    assert [(r["Paid"], r["West n"], r["North n"]) for r in ds.rows] == [(True, 1, 3),
                                                                         (False, 1, 1)]
    # the values not listed: named, their rows counted, left out of every column
    assert ds.pivot_unlisted == [{"name": "South", "v": "South", "t": "text", "rows": 4},
                                 {"name": "East", "v": "East", "t": "text", "rows": 2}]
    assert ds.left_out["Region"]["not_listed"].count == 6
    assert ds.left_out["Region"]["not_listed"].first[:2] == ("Sales!A4", "Sales!A6")
    assert (ds.matched, ds.used) == (12, 6)
    assert ds.as_dict()["pivot_unlisted"][1] == {"name": "East", "v": "East", "t": "text",
                                                 "rows": 2}
    # every value listed: nothing left out, and the columns in the order given
    every = wb.derive("t:Sales", pivot="Region",
                      pivot_values=["West", "East", "South", "North"])
    assert every.columns == ("West", "East", "South", "North")
    assert every.rows == [{"West": 2, "East": 2, "South": 4, "North": 4}]
    assert every.pivot_unlisted == [] and every.left_out == {} and "pivot_unlisted" not in \
        every.as_dict()


def test_a_group_whose_rows_hold_only_unlisted_values_is_still_a_row(wb):
    """The columns are the list, the rows the groups: a group of unlisted values only is a
    row of empty cells, its rows counted as left out."""
    ds = wb.derive("t:Sales", group_by=["Region"], pivot="Rep", pivot_values=["Ann", "Cy"],
                   measures=[{"fn": "count", "as": "n"},
                             {"fn": "distinct", "column": "Day", "as": "days"},
                             {"fn": "sum", "column": "Rate", "as": "rate"}])
    assert [(r["Region"], r["Ann n"], r["Cy n"]) for r in ds.rows] == [
        ("North", 2, 0), ("South", 0, 2), ("East", 0, 0), ("West", 0, 0)]
    assert ds.rows[2] == {"Region": "East", "Ann n": 0, "Ann days": 0, "Ann rate": None,
                          "Cy n": 0, "Cy days": 0, "Cy rate": None}
    assert ds.groups == 4 and (ds.matched, ds.used) == (12, 4)
    assert ds.provenance[2]["Region"].count == 0          # its key's sources: the rows used
    assert ds.left_out["Rep"]["not_listed"].count == 8
    # sorted and cut as any row
    top = wb.derive("t:Sales", group_by=["Region"], pivot="Rep", pivot_values=["Ann", "Cy"],
                    sort=[("Cy", "desc"), "Region"], top=3)
    assert [r["Region"] for r in top.rows] == ["South", "East", "North"]
    # bins keep theirs: a row outside every bin has no bin, and its group no row
    binned = wb.derive("t:Sales", bins={"column": "Rate", "edges": [0, 0.2]}, group_by=["Region"],
                       pivot="Paid", pivot_values=["true"])
    assert [r["Region"] for r in binned.rows] == ["North", "South"]   # not East, West
    assert {u["name"]: u["rows"] for u in ds.pivot_unlisted} == {"Bob": 2, "Dee": 2,
                                                                  "Eve": 2, "Fay": 2}
    # without a group, one row, even when no row holds a listed value
    none = wb.derive("t:Sales", pivot="Rep", pivot_values=["Gus"],
                     measures=[{"fn": "count", "as": "n"},
                               {"fn": "sum", "column": "Amount", "as": "total"}])
    assert none.rows == [{"Gus n": 0, "Gus total": None}] and none.used == 0


def test_pivot_values_match_by_type_and_value_as_groups_do():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["Size", "Code", "Qty"], [5, "5", 1], [5.0, "05", 2], ["5", "5", 4],
                      [7, "x", 8], [None, "5", 16], [True, "x", 32]])
    sheet.table("P", "A1:C7", ["Size", "Code", "Qty"], write_header=False)
    workbook = _open(book)
    measures = [{"fn": "sum", "column": "Qty", "as": "q"}]
    # a number column: "5", 5 and 5.0 are the number 5; the text "5" is another value
    for listed in (["5", "6"], [5, "6"], [5.0, 6]):
        ds = workbook.derive("t:P", pivot="Size", pivot_values=listed, measures=measures)
        assert ds.columns == ("5", "6") and ds.rows == [{"5": 3.0, "6": None}], listed
        assert ds.pivot_columns["5"] == {"measure": "q", "v": 5, "t": "number"}
    assert ds.pivot_unlisted == [{"name": "5", "v": "5", "t": "text", "rows": 1},
                                 {"name": "7", "v": 7, "t": "number", "rows": 1},
                                 {"name": "(blank)", "v": None, "t": "blank", "rows": 1},
                                 {"name": "true", "v": True, "t": "bool", "rows": 1}]
    # a text column: "5" is that text, exactly ("05" is another)
    ds = workbook.derive("t:P", pivot="Code", pivot_values=["5", "05"], measures=measures)
    assert ds.rows == [{"5": 21.0, "05": 2.0}]
    for listed, code in ((["5", "x"], "wrong_type"), (["5", "5.0"], "bad_pivot"),
                         ([True], "wrong_type")):
        with pytest.raises(DeriveError) as caught:
            workbook.derive("t:P", pivot="Size", pivot_values=listed)
        assert (caught.value.code, caught.value.field) == (code, "pivot_values"), listed
    with pytest.raises(DeriveError) as caught:
        workbook.derive("t:P", pivot="Code", pivot_values=[5])
    assert (caught.value.code, caught.value.column_type) == ("wrong_type", "text")
    assert str(caught.value) == ("Code is a text column: each of pivot_values is a value of "
                                 "its type, as filters are written")


def test_pivot_values_of_dates_and_bools(wb):
    ds = wb.derive("t:Sales", pivot="Day", pivot_values=["2026-01-05", "2026-12-25"])
    assert ds.rows == [{"2026-01-05": 2, "2026-12-25": 0}]
    assert ds.pivot_columns["2026-12-25"] == {"measure": "count", "v": "2026-12-25",
                                              "t": "date"}
    paid = wb.derive("t:Sales", group_by=["Region"], pivot="Paid",
                     pivot_values=["false", "true"])
    assert paid.columns == ("Region", "false", "true")
    assert paid.rows[0] == {"Region": "North", "false": 1, "true": 3}


@pytest.mark.parametrize("options", [
    {"pivot_values": ["North"]},
    {"pivot": "Region", "pivot_values": "North"},
    {"pivot": "Region", "pivot_values": []},
    {"pivot": "Region", "pivot_values": [{"v": "North"}]},
    {"pivot": "Region", "pivot_values": [None]},
])
def test_bad_pivot_values_are_refused(wb, options):
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", **options)
    assert (caught.value.code, caught.value.field) == ("bad_pivot", "pivot_values")


def test_pivot_values_are_bounded_and_never_quoted(wb, monkeypatch):
    monkeypatch.setattr(derive_module, "MAX_PIVOT_VALUES", 2)
    with pytest.raises(DeriveError) as caught:
        wb.derive("t:Sales", pivot="Region", pivot_values=["North", "South", "East"])
    assert caught.value.code == "bad_pivot" and "1 to 2 values" in str(caught.value)
    # listed, the column's other values are not limited; the names given are, at most 2
    ds = wb.derive("t:Sales", pivot="Rep", pivot_values=["Ann"])
    assert len(ds.pivot_unlisted) == 2 and ds.left_out["Rep"]["not_listed"].count == 10
    for listed in ([MARKER, MARKER], [MARKER, 1.5]):
        with pytest.raises(QueryError) as caught:
            wb.derive("t:Sales", pivot="Amount" if listed[1] == 1.5 else "Note",
                      pivot_values=listed)
        assert MARKER not in str(caught.value)


def test_pivot_values_compose_with_sort_top_and_round(wb):
    ds = wb.derive("t:Sales", group_by=["Region"], pivot="Paid", pivot_values=["true", "false"],
                   measures=[{"fn": "sum", "column": "Rate", "as": "rate"}],
                   sort=[("true", "desc")], top=2, round={"rate": 1})
    assert ds.columns == ("Region", "true", "false")
    assert ds.rows == [{"Region": "East", "true": 1.0, "false": 0.8},
                       {"Region": "West", "true": 0.6, "false": 0.4}]
    assert ds.unrounded[0] == {"Region": "East", "true": 1, "false": 0.75}
    assert ds.groups == 4


def test_pivot_values_match_a_coerced_columns_text():
    """A date column with a date stored as text: listed as a date, it matches with coerce,
    as group_by would group it; without coerce the text is a value of its own."""
    workbook = _open(checks())
    plain = workbook.derive("t:Checks", pivot="Done", pivot_values=["2026-03-01", "2026-03-09"])
    assert plain.rows == [{"2026-03-01": 0, "2026-03-09": 1}]
    assert plain.pivot_unlisted == [
        {"name": "2026-03-01", "v": "2026-03-01", "t": "text", "rows": 1},
        {"name": "2026-03-12", "v": "2026-03-12", "t": "date", "rows": 1},
        {"name": "2026-03-10", "v": "2026-03-10", "t": "date", "rows": 1}]
    read = workbook.derive("t:Checks", pivot="Done", pivot_values=["2026-03-01", "2026-03-09"],
                           coerce=["date_as_text"])
    assert read.rows == [{"2026-03-01": 1, "2026-03-09": 1}]
    assert read.coerced["Done"].as_dict() == {"count": 1, "first": ["Log!E3"]}
    assert [u["name"] for u in read.pivot_unlisted] == ["2026-03-12", "2026-03-10"]
    # a column of dates all stored as text pivots as dates when coerced
    seen = workbook.derive("t:Checks", pivot="Seen", coerce=["date_as_text"],
                           pivot_values=["2026-03-20", "2026-03-04"])
    assert seen.rows == [{"2026-03-20": 1, "2026-03-04": 1}]
    assert seen.pivot_columns["2026-03-20"]["t"] == "date"


def test_pivot_values_from_a_dataset(wb):
    by_rep = wb.derive("t:Sales", group_by=["Region", "Rep"],
                       measures=[{"fn": "count", "as": "n"}])
    ds = wb.derive(by_rep, pivot="Region", pivot_values=["South", "North", "Central"],
                   measures=[{"fn": "sum", "column": "n", "as": "rows"}])
    assert ds.rows == [{"South": 4.0, "North": 4.0, "Central": None}]
    # the sources carried through to the original cells (a count's: its rows' first column)
    assert [str(s) for s in ds.provenance[0]["South"].sources] == [
        "Sales!A4", "Sales!A6", "Sales!A10", "Sales!A13"]
    assert [u["name"] for u in ds.pivot_unlisted] == ["East", "West"]


# -- sort and top --------------------------------------------------------------------------------


def test_sort_by_measures_and_keys(wb):
    measures = [{"fn": "sum", "column": "Amount", "as": "total"}, {"fn": "count", "as": "n"}]
    ds = wb.derive("t:Sales", group_by=["Region"], measures=measures,
                   sort=[("total", "desc")])
    # a group without a value comes last either way
    assert [row["Region"] for row in ds.rows] == ["North", "South", "West", "East"]
    ds = wb.derive("t:Sales", group_by=["Region"], measures=measures, sort=["total"])
    assert [row["Region"] for row in ds.rows] == ["West", "South", "North", "East"]
    ds = wb.derive("t:Sales", group_by=["Region"], measures=measures,
                   sort=[("n", "desc"), "region"])
    assert [row["Region"] for row in ds.rows] == ["North", "South", "East", "West"]
    ds = wb.derive("t:Sales", group_by=["Region"], measures=measures,
                   sort=[{"column": "Region", "desc": True}], top=2)
    assert [row["Region"] for row in ds.rows] == ["West", "South"] and ds.groups == 4
    assert len(ds.provenance) == 2 and ds.provenance[0]["Region"].count == 2


def test_sort_a_mixed_key_column_by_type_then_value(wb):
    ds = wb.derive("t:Sales", group_by=["Amount"], sort=["Amount"])
    keys = [row["Amount"] for row in ds.rows]
    assert keys == [-1e16, -2.5, 0.125, 1.0, 2.5, 2.675, 1e16, "7.25", True, "#N/A", None]


def test_sort_bins_by_their_order_and_dates_by_their_day(wb):
    ds = wb.derive("t:Sales", bins={"column": "Rate", "edges": [0, 0.25, 0.5, 1]},
                   sort=[("Rate", "desc")])
    assert [row["Rate"] for row in ds.rows] == ["[0.5, 1]", "[0.25, 0.5)", "[0, 0.25)"]
    ds = wb.derive("t:Sales", group_by=["Region"],
                   measures=[{"fn": "min", "column": "Day", "as": "first"}], sort=["first"])
    assert [row["first"] for row in ds.rows] == ["2025-12-31", "2026-01-05", "2026-02-01",
                                                 "2026-04-01"]


# -- the result as data ----------------------------------------------------------------------


def test_the_result_as_plain_data(wb):
    ds = wb.derive("t:Sales", where={"Region": {"in": ["West"]}}, group_by=["Rep"],
                   measures=[{"fn": "sum", "column": "Amount", "as": "total"},
                             {"fn": "share", "column": "Rate", "as": "part"}],
                   round=1, id="west")
    data = ds.as_dict()
    assert data == {
        "from": "t:Sales", "id": "west", "columns": ["Rep", "total", "part"],
        "key_columns": 1, "types": {"Rep": "text", "total": "number", "part": "percent"},
        "rows": [["Fay", 0.0, 1.0]],
        "rounding": {"rule": "half_away_from_zero", "digits": {"total": 1, "part": 1}},
        "provenance": [[{"sources": ["Sales!B11:B12"], "count": 2},
                        {"sources": ["Sales!C11:C12"], "count": 2},
                        {"sources": ["Sales!D11:D12"], "count": 2}]],
        "matched": 2, "used": 2, "groups": 1,
        "totals": {"part": {"value": 1.0, "sources": ["Sales!D11:D12"], "count": 2}},
        "cache": {"absent": 1}}


# -- bounded ----------------------------------------------------------------------------------


def _big(rows: int, groups: int = 4, interleaved: bool = False) -> bytes:
    """A table of ``rows`` rows: a category (consecutive runs of rows, or interleaved), an
    amount and a date-free number column."""
    book = WorkbookBuilder()
    book.sheet("Big")
    parts = book.parts()
    names = [f"cat-{k}" for k in range(groups)]
    strings = "".join(f"<si><t>{n}</t></si>" for n in ["Category", "Amount"] + names)
    parts["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{NS_MAIN}" '
        f'count="{groups + 2}" uniqueCount="{groups + 2}">{strings}</sst>').encode()
    per = max(1, rows // groups)

    def group(r: int) -> int:
        return (r % groups) if interleaved else min(groups - 1, (r - 2) // per)

    body = ['<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>']
    body += [f'<row r="{r}"><c r="A{r}" t="s"><v>{2 + group(r)}</v></c>'
             f'<c r="B{r}"><v>{r * 0.25}</v></c></row>' for r in range(2, rows + 2)]
    parts["xl/worksheets/sheet1.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><sheetData>{"".join(body)}'
        f"</sheetData></worksheet>").encode()
    if "xl/sharedStrings.xml" not in book.parts():
        parts = _with_strings(parts)
    return zip_parts(parts)


def _with_strings(parts: dict[str, bytes]) -> dict[str, bytes]:
    """``parts`` with the shared strings part related and typed."""
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    rels = rels.replace("</Relationships>", f'<Relationship Id="rIdS" Type="{NS_R}/sharedStrings"'
                        ' Target="sharedStrings.xml"/></Relationships>')
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
    types = parts["[Content_Types].xml"].decode()
    types = types.replace("</Types>", '<Override PartName="/xl/sharedStrings.xml" ContentType='
                          '"application/vnd.openxmlformats-officedocument.spreadsheetml.'
                          'sharedStrings+xml"/></Types>')
    parts["[Content_Types].xml"] = types.encode()
    return parts


_MEASURES = [{"fn": "count", "as": "n"}, {"fn": "sum", "column": "Amount", "as": "total"},
             {"fn": "mean", "column": "Amount", "as": "mean"},
             {"fn": "min", "column": "Amount", "as": "low"},
             {"fn": "max", "column": "Amount", "as": "high"},
             {"fn": "share", "column": "Amount", "as": "part"}]


def _derive_peak(data: bytes, rows: int, measures=_MEASURES):
    workbook = Workbook.open(data)
    workbook.declare_region("Big", f"Big!A1:B{rows + 1}")
    workbook.scan("Big")
    tracemalloc.start()
    try:
        started = time.perf_counter()
        ds = workbook.derive("t:Big", group_by=["Category"], measures=measures)
        elapsed = time.perf_counter() - started
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    return ds, peak, elapsed


def test_a_large_table_is_derived_in_memory_bounded_by_its_groups():
    small = _derive_peak(_big(3_000), 3_000)
    large = _derive_peak(_big(30_000), 30_000)
    ds = large[0]
    assert ds.used == 30_000 and [row["n"] for row in ds.rows] == [7_500] * 4
    assert ds.rows[0]["total"] == math.fsum(r * 0.25 for r in range(2, 7_502))
    assert [s.text for s in ds.sources["cat-0"]["total"]] == ["Big!B2:B7501"]
    # ten times the rows, nearly the same peak: what is held is a row and each group's
    # state, never the table
    assert large[1] < small[1] * 1.5 + 64_000, (small[1], large[1])
    assert large[1] < 1_500_000


def test_interleaved_groups_hold_their_runs():
    """Rows of a group that are not consecutive are one run each: the sources grow with
    the rows (16 bytes a run), not with anything else."""
    rows = 20_000
    ds, peak, _ = _derive_peak(_big(rows, interleaved=True), rows,
                               [{"fn": "sum", "column": "Amount", "as": "total"}])
    assert sum(where["total"].count for where in ds.provenance) == rows
    assert sum(len(where["total"].sources) for where in ds.provenance) == rows
    assert peak < 2_000_000, peak


@pytest.mark.slow
def test_a_200000_row_table_is_derived_in_bounded_memory():
    ds, peak, elapsed = _derive_peak(_big(200_000, groups=6), 200_000)
    assert ds.used == 200_000 and len(ds.rows) == 6
    assert peak < 1_500_000, peak
    print(f"\n200,000 rows x 2 columns derived (6 groups, 6 measures): peak "
          f"{peak / 1e6:.2f} MB of Python allocations, {elapsed:.1f} s (under tracemalloc)")
