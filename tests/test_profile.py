"""R5, "Column profile": the profile of every fixture column matches a hand-written
expectation; a large column is profiled in bounded memory.

Every workbook is built with tests/build.py; the one rule the builder will not write (an
``x14:dataValidation`` in the worksheet's ``extLst``) and the large sheet are written by
hand and put into a built package with ``zip_parts``.
"""

import math
import time
import tracemalloc
from datetime import date, datetime, time as clock

import pytest

from build import NS_MAIN, NS_R, Error, Formula, IsoDate, WorkbookBuilder, zip_parts
from xlsx_agent import ColumnStats, Count, Workbook
from xlsx_agent import profile as profile_module
from xlsx_agent.read.types import TEXT_LIMIT

MARKER = "MARKER-7Q"  # cell content that must never reach a diagnostic
LONG = "x" * TEXT_LIMIT

NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"


def _open(book: WorkbookBuilder) -> Workbook:
    return Workbook.open(book.to_bytes())


def _by_column(stats: list[ColumnStats]) -> dict[str, dict]:
    return {s.column: s.as_dict() for s in stats}


def _count(n: int, *first: str) -> dict:
    return {"count": n, "first": list(first)}


def _values(*pairs) -> list[dict]:
    return [{"v": v, "t": t, "count": n} for v, t, n in pairs]


# -- fixture: a table of every kind of column ---------------------------------------------------


def inventory() -> WorkbookBuilder:
    """A table with a totals row: text keys, text with a validation list from another
    sheet, numbers with a number stored as text and an error, dates with a date stored as
    text, formulas (one without a cached value, two volatile), text with an inline
    validation list, a hidden column with text at the limit, and percentages."""
    book = WorkbookBuilder()
    stock = book.sheet("Stock")
    share = book.style(num_fmt="0%")
    stock.rows("A2", [
        ["SKU-0001", "Bolt", 0.25, date(2026, 1, 5), Formula("C2*4", cached=1), "Active", "first"],
        ["SKU-0002", "Nut", "12.50", date(2026, 1, 6), Formula("C3*4"), "Active", MARKER],
        ["SKU-0003", "Bolt", Error("#N/A"), "2026-02-01", Formula("TODAY()-C4", cached=3),
         "Retired", None],
        ["SKU-0004", MARKER, 3.75, date(2026, 2, 2), Formula("RAND()", cached=0.5), "Active",
         LONG],
        ["SKU-0005", "Washer", 4, date(2026, 2, 4), 2, "Paused", None]])
    for row, value in zip(range(2, 7), (0.1, 0.25, 0.5, 0.25, 1)):
        stock.cell(f"H{row}", value, style=share)
    stock.cell("A7", "Total")
    stock.cell("C7", Formula("SUBTOTAL(109,[Price])", cached=8))
    stock.table("Products", "A1:H7", ["SKU", "Item", "Price", "Added", "Qty", "Status", "Note",
                                      "Share"],
                totals_labels={"SKU": "Total"}, totals_functions={"Price": "sum"})
    stock.hide_column("G")
    stock.validation_list("F2:F6", ["Active", "Retired"])
    stock.validation_list("B2:B100", source="Lists!$A$1:$A$3")
    stock.validation("C2:C6", "decimal", "0", "1000", operator="between")  # not a list
    lists = book.sheet("Lists")
    lists.rows("A1", [["Bolt"], ["Nut"], ["Washer"]])
    return book


INVENTORY = {
    "SKU": {
        "address": "t:Products[SKU]", "column": "SKU", "range": "Stock!A2:A6",
        "rows": 5, "cells": 5, "filled": 5, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "text", "types": {"text": _count(5, "Stock!A2", "Stock!A3", "Stock!A4")},
        "distinct": 5,
        "values": _values(*[(f"SKU-000{i}", "text", 1) for i in range(1, 6)]),
        "text_length": {"median": 8, "max": 8, "median_words": 1, "max_words": 1,
                        "longest": "Stock!A2"}},
    "Item": {
        "address": "t:Products[Item]", "column": "Item", "range": "Stock!B2:B6",
        "rows": 5, "cells": 5, "filled": 5, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "text", "types": {"text": _count(5, "Stock!B2", "Stock!B3", "Stock!B4")},
        "distinct": 4,
        "values": _values(("Bolt", "text", 2), ("Nut", "text", 1), (MARKER, "text", 1),
                          ("Washer", "text", 1)),
        # lengths 4, 3, 4, 9, 6
        "text_length": {"median": 4, "max": 9, "median_words": 1, "max_words": 1,
                        "longest": "Stock!B5"},
        "validation": [{"values": ["Bolt", "Nut", "Washer"], "cells": 5,
                        "source": "Lists!A1:A3"}]},
    "Price": {
        "address": "t:Products[Price]", "column": "Price", "range": "Stock!C2:C6",
        "rows": 5, "cells": 5, "filled": 5, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "number",
        "types": {"number": _count(3, "Stock!C2", "Stock!C5", "Stock!C6"),
                  "text": _count(1, "Stock!C3"), "error": _count(1, "Stock!C4")},
        "min": 0.25, "max": 4, "mean": 8 / 3, "measured": 3,
        "distinct": 5,
        "values": _values((0.25, "number", 1), ("12.50", "text", 1), ("#N/A", "error", 1),
                          (3.75, "number", 1), (4, "number", 1)),
        "text_length": {"median": 5, "max": 5, "median_words": 1, "max_words": 1,
                        "longest": "Stock!C3"},
        "flags": {"mixed_types": _count(2, "Stock!C3", "Stock!C4"),
                  "number_as_text": _count(1, "Stock!C3")}},
    "Added": {
        "address": "t:Products[Added]", "column": "Added", "range": "Stock!D2:D6",
        "rows": 5, "cells": 5, "filled": 5, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "date",
        "types": {"date": _count(4, "Stock!D2", "Stock!D3", "Stock!D5"),
                  "text": _count(1, "Stock!D4")},
        # days 0, 1, 28 and 30 after 5 January: the mean is 14.75 days after, a date
        "min": "2026-01-05", "max": "2026-02-04", "mean": "2026-01-19", "measured": 4,
        "distinct": 5,
        "values": _values(("2026-01-05", "date", 1), ("2026-01-06", "date", 1),
                          ("2026-02-01", "text", 1), ("2026-02-02", "date", 1),
                          ("2026-02-04", "date", 1)),
        "text_length": {"median": 10, "max": 10, "median_words": 1, "max_words": 1,
                        "longest": "Stock!D4"},
        "date_text": {"min": "2026-02-01", "max": "2026-02-01", "measured": 1},
        "flags": {"date_as_text": _count(1, "Stock!D4"),
                  "mixed_types": _count(1, "Stock!D4")}},
    "Qty": {
        "address": "t:Products[Qty]", "column": "Qty", "range": "Stock!E2:E6",
        "rows": 5, "cells": 5, "filled": 4, "fill": 0.8, "formulas": 4, "formula_share": 0.8,
        "type": "number", "types": {"number": _count(4, "Stock!E2", "Stock!E4", "Stock!E5")},
        "min": 0.5, "max": 3, "mean": 1.625, "measured": 4,
        "distinct": 4,
        "values": _values((1, "number", 1), (3, "number", 1), (0.5, "number", 1),
                          (2, "number", 1)),
        "flags": {"no_cached_value": _count(1, "Stock!E3"),
                  "volatile": _count(2, "Stock!E4", "Stock!E5")}},
    "Status": {
        "address": "t:Products[Status]", "column": "Status", "range": "Stock!F2:F6",
        "rows": 5, "cells": 5, "filled": 5, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "text", "types": {"text": _count(5, "Stock!F2", "Stock!F3", "Stock!F4")},
        "distinct": 3,
        "values": _values(("Active", "text", 3), ("Retired", "text", 1), ("Paused", "text", 1)),
        "text_length": {"median": 6, "max": 7, "median_words": 1, "max_words": 1,
                        "longest": "Stock!F4"},
        "validation": [{"values": ["Active", "Retired"], "cells": 5}]},
    "Note": {
        "address": "t:Products[Note]", "column": "Note", "range": "Stock!G2:G6",
        "rows": 5, "cells": 3, "filled": 3, "fill": 0.6, "formulas": 0, "formula_share": 0.0,
        "type": "text", "types": {"text": _count(3, "Stock!G2", "Stock!G3", "Stock!G5")},
        "distinct": 3,
        "values": _values(("first", "text", 1), (MARKER, "text", 1), (LONG, "text", 1)),
        "text_length": {"median": 9, "max": TEXT_LIMIT, "median_words": 1, "max_words": 1,
                        "longest": "Stock!G5"},
        "flags": {"at_limit": _count(1, "Stock!G5"),
                  "hidden": _count(5, "Stock!G2", "Stock!G3", "Stock!G4")}},
    "Share": {
        "address": "t:Products[Share]", "column": "Share", "range": "Stock!H2:H6",
        "rows": 5, "cells": 5, "filled": 5, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "percent",
        "types": {"percent": _count(5, "Stock!H2", "Stock!H3", "Stock!H4")},
        "min": 0.1, "max": 1, "mean": math.fsum([0.1, 0.25, 0.5, 0.25, 1]) / 5, "measured": 5,
        "distinct": 4,
        "values": _values((0.25, "percent", 2), (0.1, "percent", 1), (0.5, "percent", 1),
                          (1, "percent", 1))},
}


def test_every_column_of_a_table_matches_its_expectation():
    wb = _open(inventory())
    assert _by_column(wb.profile("t:Products")) == INVENTORY


def test_a_column_alone_matches_its_profile_in_the_table():
    wb = _open(inventory())
    products = wb.table("t:Products")
    for name, expected in INVENTORY.items():
        assert products.column(name.upper()).stats().as_dict() == expected
        assert wb.column(f"t:Products[{name}]").stats().as_dict() == expected
    assert repr(products.column("price")) == "Column('t:Products[Price]')"
    assert str(products.column("price").range) == "Stock!C2:C6"
    assert _by_column(products.profile(["Share", "sku"])) == {
        "Share": INVENTORY["Share"], "SKU": INVENTORY["SKU"]}
    assert _by_column(wb.profile(products, ["Qty"])) == {"Qty": INVENTORY["Qty"]}


def test_the_stats_are_typed_values():
    stats = _open(inventory()).column("t:Products[Price]").stats()
    assert stats.type == "number" and stats.fill == 1.0 and stats.measured == 3
    assert stats.flags["number_as_text"] == Count(1, ("Stock!C3",))
    assert stats.mean == pytest.approx(2.6666666666666665)


# -- fixture: dates and times in the 1904 system, a detected region -----------------------------


def log_1904() -> WorkbookBuilder:
    """A detected region in a 1904 workbook: dates, times of day, dates with times,
    elapsed times, ISO dates (``t="d"``), booleans, and a tie between two types."""
    book = WorkbookBuilder(date1904=True)
    sheet = book.sheet("Log")
    elapsed = book.style(num_fmt="[h]:mm:ss")
    sheet.rows("A1", [["Day", "At", "When", "Took", "Stamp", "Flag", "Tie"],
                      [date(2026, 1, 1), clock(8, 0), datetime(2026, 1, 1, 6, 0), None,
                       IsoDate(date(2026, 3, 1)), True, 1],
                      [date(2026, 1, 3), clock(10, 0), datetime(2026, 1, 2, 18, 0), None,
                       IsoDate(date(2026, 2, 1)), False, "a"],
                      [date(2026, 1, 2), clock(12, 0), datetime(2026, 1, 1, 12, 0), None,
                       IsoDate(date(2026, 4, 1)), True, None]])
    for row, value in zip((2, 3, 4), (1.5, 0.5, 1.0)):
        sheet.cell(f"D{row}", value, style=elapsed)
    return book


def _plain(address, column, letter, *, type, types, filled=3, **rest):
    out = {"address": address, "column": column, "range": f"Log!{letter}2:{letter}4",
           "rows": 3, "cells": filled, "filled": filled, "fill": filled / 3, "formulas": 0,
           "formula_share": 0.0, "type": type, "types": types}
    out.update(rest)
    return out


LOG = {
    "Day": _plain("t:Log!A1[Day]", "Day", "A", type="date",
                  types={"date": _count(3, "Log!A2", "Log!A3", "Log!A4")},
                  min="2026-01-01", max="2026-01-03", mean="2026-01-02", measured=3, distinct=3,
                  values=_values(("2026-01-01", "date", 1), ("2026-01-03", "date", 1),
                                 ("2026-01-02", "date", 1))),
    "At": _plain("t:Log!A1[At]", "At", "B", type="time",
                 types={"time": _count(3, "Log!B2", "Log!B3", "Log!B4")},
                 min="08:00:00", max="12:00:00", mean="10:00:00", measured=3, distinct=3,
                 values=_values(("08:00:00", "time", 1), ("10:00:00", "time", 1),
                                ("12:00:00", "time", 1))),
    # 6, 42 and 12 hours after midnight on 1 January: the mean is 20 hours after
    "When": _plain("t:Log!A1[When]", "When", "C", type="datetime",
                   types={"datetime": _count(3, "Log!C2", "Log!C3", "Log!C4")},
                   min="2026-01-01T06:00:00", max="2026-01-02T18:00:00",
                   mean="2026-01-01T20:00:00", measured=3, distinct=3,
                   values=_values(("2026-01-01T06:00:00", "datetime", 1),
                                  ("2026-01-02T18:00:00", "datetime", 1),
                                  ("2026-01-01T12:00:00", "datetime", 1))),
    "Took": _plain("t:Log!A1[Took]", "Took", "D", type="time",
                   types={"time": _count(3, "Log!D2", "Log!D3", "Log!D4")},
                   min="PT12H0M0S", max="PT36H0M0S", mean="PT24H0M0S", measured=3, distinct=3,
                   values=_values(("PT36H0M0S", "time", 1), ("PT12H0M0S", "time", 1),
                                  ("PT24H0M0S", "time", 1))),
    # ISO text has no serial: ordered as text, no mean
    "Stamp": _plain("t:Log!A1[Stamp]", "Stamp", "E", type="date",
                    types={"date": _count(3, "Log!E2", "Log!E3", "Log!E4")},
                    min="2026-02-01", max="2026-04-01", mean=None, measured=3, distinct=3,
                    values=_values(("2026-03-01", "date", 1), ("2026-02-01", "date", 1),
                                   ("2026-04-01", "date", 1))),
    "Flag": _plain("t:Log!A1[Flag]", "Flag", "F", type="bool",
                   types={"bool": _count(3, "Log!F2", "Log!F3", "Log!F4")}, distinct=2,
                   values=_values((True, "bool", 2), (False, "bool", 1))),
    # one number, one text: a tie goes to the type listed first
    "Tie": _plain("t:Log!A1[Tie]", "Tie", "G", type="number", filled=2,
                  types={"number": _count(1, "Log!G2"), "text": _count(1, "Log!G3")},
                  min=1, max=1, mean=1.0, measured=1, distinct=2,
                  values=_values((1, "number", 1), ("a", "text", 1)),
                  text_length={"median": 1, "max": 1, "median_words": 1, "max_words": 1,
                               "longest": "Log!G3"},
                  flags={"mixed_types": _count(1, "Log!G3")}),
}


def test_dates_and_times_are_measured_as_dates_in_the_1904_system():
    wb = _open(log_1904())
    assert wb.date1904
    assert _by_column(wb.profile("t:Log!A1")) == LOG
    assert wb.column("t:Log!A1[Day]").stats().as_dict() == LOG["Day"]


def test_a_mean_date_keeps_the_type_of_the_column():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["Day"], [date(2026, 1, 1)], [date(2026, 1, 2)]])
    stats = _open(book).column("t:S!A1[Day]").stats()
    # 1.5 days after: the date shows the day, as a date format does
    assert (stats.min, stats.max, stats.mean) == ("2026-01-01", "2026-01-02", "2026-01-01")


# -- fixture: names, validation sources, a hidden sheet -----------------------------------------


def sources() -> bytes:
    """A hidden sheet with a sheet-scoped name over a headerless range; validation lists
    from a defined name, an x14 extension rule on another sheet, a formula, an unknown name
    and whole columns; and a declared region with no data rows."""
    book = WorkbookBuilder()
    stock = book.sheet("Stock", state="hidden")
    stock.rows("A2", [["Red", 1], ["Blue", 2], ["Red", 3]])
    stock.validation_list("A2:A3", source="Choices")
    stock.validation_list("A4", source="Picks")  # local to Stock: found from Stock
    stock.validation_list("B2:B4", source='INDIRECT("Lists!C1:C2")')
    stock.validation_list("B3", source="Nowhere")
    stock.validation_list("B4", source="Lists!$D:$D")
    lists = book.sheet("Lists")
    lists.rows("A1", [["Red", "Small", None, 10], ["Blue", "Large", None, 20],
                      [None, None, None, 30]])
    book.define_name("Picks", "Stock!$A$2:$B$4", sheet="Stock")
    book.define_name("Choices", "Lists!$A$1:$A$2")
    book.define_name("Empty", "Lists!$C$1")
    parts = book.parts()
    sheet = parts["xl/worksheets/sheet1.xml"]
    ext = (f'<extLst><ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" xmlns:x14="{NS_X14}">'
           f'<x14:dataValidations xmlns:xm="{NS_XM}" count="1">'
           f'<x14:dataValidation type="list"><x14:formula1><xm:f>Lists!$B$1:$B$2</xm:f>'
           f'</x14:formula1><xm:sqref>A3:A9</xm:sqref></x14:dataValidation>'
           f'</x14:dataValidations></ext></extLst>')
    parts["xl/worksheets/sheet1.xml"] = sheet.replace(b"</worksheet>",
                                                      ext.encode() + b"</worksheet>")
    return zip_parts(parts)


SOURCES = {
    "Column A": {
        "address": "t:Stock!Picks[Column A]", "column": "Column A", "range": "Stock!A2:A4",
        "rows": 3, "cells": 3, "filled": 3, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "text", "types": {"text": _count(3, "Stock!A2", "Stock!A3", "Stock!A4")},
        "distinct": 2, "values": _values(("Red", "text", 2), ("Blue", "text", 1)),
        "text_length": {"median": 3, "max": 4, "median_words": 1, "max_words": 1,
                        "longest": "Stock!A3"},
        "validation": [{"values": ["Red", "Blue"], "cells": 2, "source": "Lists!A1:A2"},
                       {"values": ["Red", 1, "Blue", 2, "Red", 3], "cells": 1,
                        "source": "Stock!A2:B4"},
                       {"values": ["Small", "Large"], "cells": 2, "source": "Lists!B1:B2"}],
        "flags": {"hidden": _count(3, "Stock!A2", "Stock!A3", "Stock!A4")}},
    "Column B": {
        "address": "t:Stock!Picks[Column B]", "column": "Column B", "range": "Stock!B2:B4",
        "rows": 3, "cells": 3, "filled": 3, "fill": 1.0, "formulas": 0, "formula_share": 0.0,
        "type": "number", "types": {"number": _count(3, "Stock!B2", "Stock!B3", "Stock!B4")},
        "min": 1, "max": 3, "mean": 2.0, "measured": 3,
        "distinct": 3, "values": _values((1, "number", 1), (2, "number", 1), (3, "number", 1)),
        "validation": [{"values": None, "cells": 3, "reason": "formula"},
                       {"values": None, "cells": 1, "reason": "unknown_name"},
                       {"values": [10, 20, 30], "cells": 1, "source": "Lists!D:D"}],
        "flags": {"hidden": _count(3, "Stock!B2", "Stock!B3", "Stock!B4")}},
}


def test_a_sheet_scoped_name_and_its_validation_sources():
    wb = Workbook.open(sources())
    assert _by_column(wb.profile("t:Stock!Picks")) == SOURCES
    stats = wb.column("t:Stock!Picks[Column A]").stats()
    # the rule's text is kept for a program, never shown
    assert stats.validation[0].formula == "Choices"
    assert "Choices" not in repr(stats.validation[0])


def test_validation_rules_are_read_with_the_scan():
    wb = Workbook.open(sources())
    rules = wb.scan("Stock").validations
    assert [(r.type, [a.ref for a in r.areas]) for r in rules] == [
        ("list", ["A2:A3"]), ("list", ["A4"]), ("list", ["B2:B4"]), ("list", ["B3"]),
        ("list", ["B4"]), ("list", ["A3:A9"])]
    built = _open(inventory()).scan("Stock").validations
    assert [r.type for r in built] == ["list", "list", "decimal"]


def test_a_long_allowed_list_is_cut(monkeypatch):
    monkeypatch.setattr(profile_module, "ALLOWED_CAP", 2)
    stats = Workbook.open(sources()).column("t:Stock!Picks[Column B]").stats()
    assert stats.validation[2].as_dict() == {"values": [10, 20], "cells": 1,
                                             "source": "Lists!D:D", "cut": True}


def test_a_region_with_no_data_rows():
    wb = Workbook.open(sources())
    wb.declare_region("Head", "Lists!A1:B1")
    assert wb.column("t:Head[Red]").stats().as_dict() == {
        "address": "t:Head[Red]", "column": "Red", "range": None, "rows": 0, "cells": 0,
        "filled": 0, "fill": 0.0, "formulas": 0, "formula_share": 0.0, "type": "blank",
        "types": {}, "distinct": 0, "values": []}


# -- distinct values, bounded --------------------------------------------------------------------


def _column_of(values) -> Workbook:
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["V"], *[[v] for v in values]])
    return _open(book)


def test_values_are_listed_only_when_few():
    few = _column_of([f"v{i % 12}" for i in range(30)]).column("t:S!A1[V]").stats()
    assert few.distinct == 12 and len(few.values) == 12
    assert few.values[0].as_dict() == {"v": "v0", "t": "text", "count": 3}
    many = _column_of([f"v{i % 13}" for i in range(30)]).column("t:S!A1[V]").stats()
    assert many.distinct == 13 and many.values is None
    # a number and its text are two values
    assert _column_of([1, "1", 1]).column("t:S!A1[V]").stats().distinct == 2


def test_distinct_counting_stops_at_its_cap(monkeypatch):
    monkeypatch.setattr(profile_module, "DISTINCT_CAP", 20)
    assert _column_of(range(20)).column("t:S!A1[V]").stats().distinct == 20
    over = _column_of(range(21)).column("t:S!A1[V]").stats()
    assert over.distinct is None and over.values is None
    assert over.as_dict()["distinct"] is None


def test_long_text_is_counted_by_digest():
    head = "y" * 100
    stats = _column_of([head + "1", head + "2", head + "1"]).column("t:S!A1[V]").stats()
    assert stats.distinct == 2
    assert [d.as_dict() for d in stats.values] == _values((head + "1", "text", 2),
                                                          (head + "2", "text", 1))


def test_text_lengths_are_in_characters_and_words_with_the_longest_cell():
    notes = ["one", "two words", "  three  short\twords ", "a b c d e f", "longest of them all",
             "longest of them all", 7]
    stats = _column_of(notes).column("t:S!A1[V]").stats()
    # words 1, 2, 3, 6, 4, 4; characters 3, 9, 21, 11, 19, 19
    assert (stats.text_words_median, stats.text_words_max) == (3.5, 6)
    assert (stats.text_length_median, stats.text_length_max) == (15, 21)
    assert stats.text_longest == "S!A4"                   # the first of the longest
    assert stats.as_dict()["text_length"] == {"median": 15, "max": 21, "median_words": 3.5,
                                              "max_words": 6, "longest": "S!A4"}
    assert _column_of([1, 2]).column("t:S!A1[V]").stats().text_longest is None


def test_a_column_of_dates_stored_as_text_gives_their_range():
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [["Seen"], ["2026-03-04"], ["03/05/2026"], ["2026-02-11 08:00"],
                                ["03/01/26"], ["soon"], [date(2026, 1, 2)]])
    plain = Workbook.open(book.to_bytes()).column("t:S!A1[Seen]").stats()
    assert plain.type == "text" and plain.measured == 0
    assert plain.flags["date_as_text"].as_dict() == {"count": 2, "first": ["S!A2", "S!A4"]}
    assert plain.as_dict()["date_text"] == {"min": "2026-02-11T08:00", "max": "2026-03-04",
                                            "measured": 2}
    # with local_dates, en-US text dates too; a two-digit year is flagged, not read
    local = Workbook.open(book.to_bytes(), local_dates=True).column("t:S!A1[Seen]").stats()
    assert local.flags["date_as_text"].count == 4
    assert (local.date_text_min, local.date_text_max, local.date_text_measured) == (
        "2026-02-11T08:00", "2026-03-05", 3)
    assert "date_text" not in _column_of(["x"]).column("t:S!A1[V]").stats().as_dict()


def test_the_text_length_median_of_an_even_count_is_the_mean_of_the_middle_two():
    stats = _column_of(["ab", "abcd", "abc", "abcdefgh"]).column("t:S!A1[V]").stats()
    assert (stats.text_length_median, stats.text_length_max) == (3.5, 8)


def test_an_exact_mean():
    # summed as math.fsum sums, streamed: 1e100, 0.1 and -1e100 add up to 0.1, where a
    # running float sum loses the 0.1
    stats = _column_of([1e100, 0.1, -1e100]).column("t:S!A1[V]").stats()
    assert stats.mean == 0.1 / 3
    assert ((1e100 + 0.1) - 1e100) / 3 == 0.0


# -- no content in diagnostics -------------------------------------------------------------------


def test_errors_name_columns_never_content():
    wb = _open(inventory())
    with pytest.raises(KeyError) as missing:
        wb.column(f"t:Products[{MARKER}]")
    assert "Products" in str(missing.value)
    with pytest.raises(ValueError) as row:
        wb.column("t:Products/SKU-0001/Price")
    assert "SKU-0001" not in str(row.value)
    for stats in wb.profile("t:Products"):
        for code, count in stats.flags.items():
            assert MARKER not in code and all(MARKER not in a for a in count.first)


# -- bounded work on a large column --------------------------------------------------------------


def _numbers(rows: int) -> bytes:
    """A headerless column of ``rows`` distinct numbers, a defined name over it, and a
    column of a few texts beside it."""
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.cell("B1", "north")
    sheet.cell("B2", "south")
    book.define_name("Nums", f"S!$A$1:$B${rows}")
    parts = book.parts()
    body = "".join(f'<row r="{r}"><c r="A{r}"><v>{r * 1.25}</v></c>'
                   f'<c r="B{r}" t="s"><v>{r % 2}</v></c></row>' for r in range(1, rows + 1))
    parts["xl/worksheets/sheet1.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><sheetData>{body}</sheetData>'
        f"</worksheet>").encode()
    return zip_parts(parts)


def _profile_peak(rows: int) -> tuple[list[ColumnStats], int, float]:
    wb = Workbook.open(_numbers(rows))
    wb.scan("S")
    tracemalloc.start()
    try:
        started = time.perf_counter()
        stats = wb.profile("t:Nums")
        elapsed = time.perf_counter() - started
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    return stats, peak, elapsed


def test_a_large_column_is_profiled_in_bounded_memory(monkeypatch):
    monkeypatch.setattr(profile_module, "DISTINCT_CAP", 1_000)
    small = _profile_peak(3_000)
    large = _profile_peak(30_000)
    numbers, texts = large[0]
    assert (numbers.rows, numbers.filled, numbers.distinct) == (30_000, 30_000, None)
    assert (numbers.min, numbers.max, numbers.mean) == (1.25, 37_500.0,
                                                        math.fsum(r * 1.25 for r in
                                                                  range(1, 30_001)) / 30_000)
    assert texts.distinct == 2 and texts.text_length_median == 5
    # ten times the rows, nearly the same peak: past the distinct cap, what is held is a
    # row, the counts and a bounded set, never the column
    assert large[1] < small[1] * 1.5 + 64_000, (small[1], large[1])
    assert large[1] < 1_500_000


@pytest.mark.slow
def test_a_200000_row_column_is_profiled_in_bounded_memory():
    stats, peak, elapsed = _profile_peak(200_000)
    assert stats[0].rows == 200_000 and stats[0].distinct is None
    # the distinct set at its cap of 10,000 keys, and a row at a time
    assert peak < 4_000_000, peak
    print(f"\n200,000 rows x 2 columns profiled: peak {peak / 1e6:.2f} MB of Python "
          f"allocations, {elapsed:.1f} s (under tracemalloc)")
