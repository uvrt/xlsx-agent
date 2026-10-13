"""R3, "Cell reader": every row of the messy-workbooks table that concerns cells has a
fixture here; a 200,000-row sheet reads in bounded memory; reading a range is lazy.

Every workbook is built with tests/build.py; the few sheets the builder cannot write
(rows out of order, a million sparse rows, entities) are written by hand below and put
into a built package with ``zip_parts``.
"""

import itertools
import subprocess
import sys
import time
import tracemalloc
from datetime import date, datetime, time as clock

import pytest

from build import (NS_MAIN, NS_R, Error, Formula, Inline, IsoDate, Rich, Run,
                   WorkbookBuilder, zip_parts)
from xlsx_agent import Locale, NameNotFound, OpenError, Problem, Workbook
from xlsx_agent.read.stream import SheetReader, is_volatile
from xlsx_agent.read.strings import decode_xstring
from xlsx_agent.read.types import TEXT_LIMIT, looks_like_date, looks_like_number, serial_text

MARKER = "MARKER-7Q"  # cell content that must never reach a diagnostic


def _open(book: WorkbookBuilder, **options) -> Workbook:
    return Workbook.open(book.to_bytes(), **options)


def _with_sheet(book: WorkbookBuilder, body: str, *, before: str = "", after: str = "",
                number: int = 1) -> bytes:
    """``book``'s package with sheet ``number`` replaced by one written by hand."""
    parts = book.parts()
    parts[f"xl/worksheets/sheet{number}.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}">{before}<sheetData>{body}</sheetData>'
        f"{after}</worksheet>").encode()
    return zip_parts(parts)


def _grid(wb: Workbook, address: str) -> list[list]:
    return [[value.v for value in row] for row in wb.read(address)]


def _types(wb: Workbook, address: str) -> list[list]:
    return [[value.t for value in row] for row in wb.read(address)]


# -- values and types ----------------------------------------------------------------------------


def _values_book() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Values")
    percent = book.style(num_fmt="0.0%")
    sheet.rows("A1", [[12, 12.5, -0.001, 1e20],
                      ["plain", Inline("inline text"), Rich([Run("new", bold=True), Run(" item")]),
                       Inline(Rich(["in", Run("line", italic=True)]))],
                      [True, False, Error("#DIV/0!"), Error("#N/A")],
                      [date(2026, 1, 5), datetime(2026, 1, 5, 10, 30), clock(10, 30, 15), None],
                      [IsoDate(date(2026, 1, 5)), IsoDate(datetime(2026, 1, 5, 10, 30)),
                       IsoDate(clock(10, 30)), None]])
    sheet.cell("A6", 0.25, style=percent)
    sheet.cell("B6", None, style=percent)  # a cell with only a style
    return book


def test_every_type_reads_as_values_and_types_says():
    wb = _open(_values_book())
    rows = [[value.as_dict() for value in row] for row in wb.read("Values!A1:D6")]
    assert rows[0] == [{"a": "Values!A1", "v": 12, "t": "number"},
                       {"a": "Values!B1", "v": 12.5, "t": "number"},
                       {"a": "Values!C1", "v": -0.001, "t": "number"},
                       {"a": "Values!D1", "v": 1e20, "t": "number"}]
    assert type(rows[0][0]["v"]) is int and type(rows[0][1]["v"]) is float
    assert [v["v"] for v in rows[1]] == ["plain", "inline text", "new item", "inline"]
    assert {v["t"] for v in rows[1]} == {"text"}
    assert [(v["v"], v["t"]) for v in rows[2]] == [(True, "bool"), (False, "bool"),
                                                   ("#DIV/0!", "error"), ("#N/A", "error")]
    assert [(v["v"], v["t"]) for v in rows[3][:3]] == [
        ("2026-01-05", "date"), ("2026-01-05T10:30:00", "datetime"), ("10:30:15", "time")]
    assert rows[3][0]["serial"] == 46027.0
    assert [(v["v"], v["t"]) for v in rows[4][:3]] == [
        ("2026-01-05", "date"), ("2026-01-05T10:30:00", "datetime"), ("10:30:00", "time")]
    assert rows[5][0] == {"a": "Values!A6", "v": 0.25, "t": "percent"}
    assert rows[5][1] == {"a": "Values!B6", "v": None, "t": "blank"}
    assert rows[3][3] == {"a": "Values!D4", "v": None, "t": "blank"}


def test_a_style_only_cell_is_blank_and_outside_the_used_range():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.cell("A1", 1)
    sheet.cell("D9", None, style=book.style(num_fmt="0.0%"))
    wb = _open(book)
    assert wb.used_range("S").ref == "A1"
    assert [value.t for value in wb.cells("S")] == ["number", "blank"]


def test_rich_text_is_flattened_and_phonetic_runs_dropped():
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [[Rich([Run("東京", bold=True), Run("都")], phonetic=[(0, 2, "トウキョウ")]),
                                 Inline(Rich(["大阪"], phonetic=[(0, 2, "オオサカ")]))]])
    wb = _open(book)
    assert _grid(wb, "S!A1:B1") == [["東京都", "大阪"]]


def test_escaped_characters_in_strings_are_decoded():
    assert decode_xstring("a_x000D_b") == "a\rb"
    assert decode_xstring("_x005F_x000D_") == "_x000D_"
    assert decode_xstring("no escapes") == "no escapes"
    book = WorkbookBuilder()
    book.sheet("S").rows("A1", [["line_x000D_break", Inline("tab_x0009_here")]])
    wb = _open(book)
    assert _grid(wb, "S!A1:B1") == [["line\rbreak", "tab\there"]]


def test_a_shared_string_index_that_is_missing():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", "only")
    data = _with_sheet(book, '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>7</v></c>'
                             '<c r="C1" t="s"><v>x</v></c></row>')
    wb = Workbook.open(data)
    row = next(wb.read("S!A1:C1"))
    assert [(v.v, v.t, v.issue) for v in row] == [("only", "text", None),
                                                 (None, "blank", "missing_string"),
                                                 (None, "blank", "missing_string")]


def test_odd_stored_values_are_reported_never_guessed():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", 1)
    data = _with_sheet(book, '<row r="1"><c r="A1"><v>twelve</v></c><c r="B1" t="b"><v>2</v></c>'
                             '<c r="C1" t="d"><v>soon</v></c><c r="D1"><v>NaN</v></c>'
                             '<c r="E1"><v>１２</v></c><c r="F1" t="zz"><v>3</v></c></row>')
    row = next(Workbook.open(data).read("S!A1:F1"))
    assert [(v.t, v.issue) for v in row] == [("blank", "bad_value")] * 5 + [("number", None)]


# -- dates ---------------------------------------------------------------------------------------


@pytest.mark.parametrize("serial, kind, expected", [
    (1, "date", "1900-01-01"),
    (59, "date", "1900-02-28"),
    (60, "date", "1900-02-29"),       # the day the 1900 system counts (ECMA-376 §18.17.4.1)
    (61, "date", "1900-03-01"),
    (2958465, "date", "9999-12-31"),
    (60.5, "datetime", "1900-02-29T12:00:00"),
    (0.5, "time", "12:00:00"),
    (1.75, "time", "18:00:00"),
    (0.0000115740740740741, "time", "00:00:01"),
    (0.5000005787, "time", "12:00:00.050"),
    (1.99999999999, "datetime", "1900-01-02T00:00:00"),  # rounded to the millisecond
    (46027.4375, "datetime", "2026-01-05T10:30:00"),
])
def test_the_1900_date_system(serial, kind, expected):
    assert serial_text(serial, kind, False) == expected


@pytest.mark.parametrize("serial, kind", [(0, "date"), (0.5, "datetime"), (-1, "date"),
                                          (-0.5, "time"), (2958466, "date"),
                                          (float("inf"), "date")])
def test_outside_the_1900_date_system(serial, kind):
    assert serial_text(serial, kind, False) is None


def test_the_1904_date_system():
    assert serial_text(0, "date", True) == "1904-01-01"
    assert serial_text(1, "date", True) == "1904-01-02"
    assert serial_text(2957003, "date", True) == "9999-12-31"
    assert serial_text(2957004, "date", True) is None
    assert serial_text(0.25, "datetime", True) == "1904-01-01T06:00:00"
    book = WorkbookBuilder(date1904=True)
    book.sheet("S").rows("A1", [[date(2026, 1, 5), datetime(1904, 1, 1, 6), 1462]])
    wb = _open(book)
    assert wb.date1904
    row = next(wb.read("S!A1:C1"))
    assert [(v.v, v.t) for v in row[:2]] == [("2026-01-05", "date"), ("1904-01-01T06:00:00", "datetime")]
    assert row[0].serial == 44565.0  # the same day is 1,462 days fewer than in the 1900 system
    assert not _open(_values_book()).date1904


def test_elapsed_time_is_a_duration():
    assert serial_text(1.5, "time", False, elapsed=True) == "PT36H0M0S"
    assert serial_text(0.0006944444444, "time", False, elapsed=True) == "PT0H1M0S"
    assert serial_text(1.5000005787, "time", False, elapsed=True) == "PT36H0M0.05S"


def test_dates_on_the_sheet_and_out_of_range_serials():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    day = book.style(num_fmt="yyyy-mm-dd")
    sheet.rows("A1", [[1, 60, 61, 0, -3, 3e6]], style=day)
    sheet.cell("A2", 1.5, style=book.style(num_fmt="[h]:mm:ss"))
    sheet.cell("B2", 0.5, style=book.style(num_fmt_id=20))
    wb = _open(book)
    row = next(wb.read("S!A1:F1"))
    assert [(v.v, v.t, v.issue) for v in row] == [
        ("1900-01-01", "date", None), ("1900-02-29", "date", None), ("1900-03-01", "date", None),
        (0, "number", "date_out_of_range"), (-3, "number", "date_out_of_range"),
        (3000000, "number", "date_out_of_range")]
    assert [(v.v, v.t) for v in next(wb.read("S!A2:B2"))] == [("PT36H0M0S", "time"),
                                                              ("12:00:00", "time")]


# -- formulas ------------------------------------------------------------------------------------


def _formulas_book() -> WorkbookBuilder:
    book = WorkbookBuilder()
    sheet = book.sheet("Calc")
    sheet.rows("A1", [[1, 2, 3]])
    sheet.cell("D1", Formula("SUM(A1:C1)", cached=6))
    sheet.cell("D2", Formula('A1&"x"', cached="1x"))
    sheet.cell("D3", Formula("A1>0", cached=True))
    sheet.cell("D4", Formula("1/0", cached=Error("#DIV/0!")))
    sheet.cell("D5", Formula("SUM(Z1:Z9)"))              # no cached value
    sheet.cell("D6", Formula('IF(A1>9,"big","")', cached=""))
    sheet.cell("D7", Formula("TODAY()", cached=46027))
    sheet.cell("E1", Formula("A1*2", cached=2, kind="shared", ref="E1:E3", si=0))
    sheet.cell("E2", Formula("", cached=4, kind="shared", si=0))
    sheet.cell("E3", Formula("", kind="shared", si=0))     # a member with no cached value
    sheet.cell("F1", Formula("A1:A3*2", cached=2, kind="array", ref="F1:F3"))
    sheet.cell("F2", 4)
    sheet.cell("F3", 6)
    sheet.cell("G1", Formula("NOW()", cached=46027.5, kind="shared", ref="G1:G2", si=1))
    sheet.cell("G2", Formula("", cached=46027.5, kind="shared", si=1))
    return book


def test_cached_formula_values_and_no_cached_value():
    wb = _open(_formulas_book())
    column = [row[0] for row in wb.read("Calc!D1:D7")]
    assert [(v.v, v.t, v.formula, v.issue) for v in column] == [
        (6, "number", True, None), ("1x", "text", True, None), (True, "bool", True, None),
        ("#DIV/0!", "error", True, None), (None, "blank", True, "no_cached_value"),
        ("", "text", True, None), (46027, "number", True, None)]
    assert column[4].as_dict() == {"a": "Calc!D5", "v": None, "t": "blank", "formula": True,
                                   "issue": "no_cached_value"}
    assert column[0].formula_text == "SUM(A1:C1)"
    assert not wb.cell("Calc!A1").formula


def test_shared_and_array_formulas():
    wb = _open(_formulas_book())
    e1, e2, e3 = (row[0] for row in wb.read("Calc!E1:E3"))
    assert (e1.v, e1.formula_text, e1.shared_from) == (2, "A1*2", None)
    assert (e2.v, e2.formula_text, e2.shared_from) == (4, None, "Calc!E1")
    assert (e3.t, e3.issue, e3.shared_from) == ("blank", "no_cached_value", "Calc!E1")
    f1, f2, f3 = (row[0] for row in wb.read("Calc!F1:F3"))
    assert (f1.v, f1.array, f1.formula) == (2, "F1:F3", True)
    assert (f2.v, f2.array, f2.formula) == (4, None, False)  # inside the block: a value only
    assert f1.as_dict()["array"] == "F1:F3"


def test_a_dynamic_array_spill_and_an_external_reference_keep_their_cached_values():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", 1)
    data = _with_sheet(book, '<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f>'
                             '<v>1</v></c><c r="B1"><f>[1]Prices!B2</f><v>9.5</v></c></row>'
                             '<row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>')
    wb = Workbook.open(data)
    spill = [row[0] for row in wb.read("S!A1:A3")]
    assert [(v.v, v.array) for v in spill] == [(1, "A1:A3"), (2, None), (3, None)]
    linked = wb.cell("S!B1")
    assert (linked.v, linked.t, linked.formula, linked.issue) == (9.5, "number", True, None)


def test_volatile_functions_are_marked():
    wb = _open(_formulas_book())
    assert wb.cell("Calc!D7").volatile
    assert not wb.cell("Calc!D1").volatile
    g1, g2 = (row[0] for row in wb.read("Calc!G1:G2"))
    assert g1.volatile and g2.volatile  # a shared group's member, from its first cell
    assert is_volatile("_xlfn.NOW()") and is_volatile("1+ offset(A1,1,1)") and is_volatile("INDIRECT ( \"A1\" )")
    assert not is_volatile('IF(A1="NOW()",1,2)')
    assert not is_volatile("'TODAY() sheet'!A1+1")
    assert not is_volatile("RANDARRAY(3)") and not is_volatile("MYNOW()") and not is_volatile("NOW")


def test_calculation_flags_say_cached_values_may_be_stale():
    for calc, flags in [({"fullCalcOnLoad": "1"}, ["full_calc_on_load"]),
                        ({"calcMode": "manual"}, ["manual_calculation"]),
                        ({"fullCalcOnLoad": "1", "calcMode": "manual"},
                         ["full_calc_on_load", "manual_calculation"]),
                        ({"calcMode": "auto"}, []), (None, [])]:
        book = WorkbookBuilder(calc=calc)
        book.sheet("S").cell("A1", 1)
        wb = _open(book)
        assert wb.calculation == flags
        assert [p.code for p in wb.problems] == flags


def test_problems_by_code_and_first_address():
    wb = _open(_formulas_book())
    # the first address in the order of the file: row 3's E3 comes before row 5's D5
    assert wb.problems == [Problem("no_cached_value", "Calc!E3", 2),
                           Problem("volatile", "Calc!G1", 3)]


# -- the messy-workbooks table -------------------------------------------------------------------


def test_hidden_sheets_rows_and_columns_mark_their_values_hidden():
    book = WorkbookBuilder()
    front = book.sheet("Front")
    front.rows("A1", [[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    front.hide_row(2)
    front.hide_column("C")
    front.hide_row(5)  # a hidden row with no cells, inside a read range
    book.sheet("Back", state="hidden").cell("A1", 1)
    book.sheet("Deep", state="veryHidden").cell("A1", 1)
    wb = _open(book)
    hidden = [[value.hidden for value in row] for row in wb.read("Front!A1:D5")]
    assert hidden == [[False, False, True, False], [True, True, True, True],
                      [False, False, True, False], [False, False, True, False],
                      [True, True, True, True]]
    assert wb.cell("Back!A1").hidden and wb.cell("Deep!A1").hidden
    assert wb.cell("Back!B7").hidden  # a blank on a hidden sheet too
    assert wb.cell("Front!B2").as_dict() == {"a": "Front!B2", "v": 5, "t": "number", "hidden": True}
    assert wb.scan("Front").hidden_rows == 2


def test_a_protected_sheet_and_workbook_are_read():
    book = WorkbookBuilder(protected=True)
    sheet = book.sheet("Locked")
    sheet.protected = True
    sheet.rows("A1", [["SKU", 4]])
    wb = _open(book)
    assert wb.protected and wb.sheets[0].protected
    assert _grid(wb, "Locked!A1:B1") == [["SKU", 4]]


def test_numbers_and_dates_stored_as_text_with_the_callers_locale():
    texts = ["12", "1,234.50", "-3e5", "12%", " 7 ", "abc", "2026-01-05", "05/01/2026",
             "1.234,50", "31/02/2026", "12 apples", "1,23,4"]
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [texts])
    sheet.cell("A2", Formula('"12"', cached="12"))  # a formula's text: not stored as text
    data = book.to_bytes()

    def issues(wb):
        return dict(zip(texts, (v.issue for v in next(wb.read(f"S!A1:L1")))))

    default = issues(Workbook.open(data))
    assert default == {"12": "number_as_text", "1,234.50": "number_as_text", "-3e5": "number_as_text",
                       "12%": "number_as_text", " 7 ": "number_as_text", "abc": None,
                       "2026-01-05": "date_as_text", "05/01/2026": None, "1.234,50": None,
                       "31/02/2026": None, "12 apples": None, "1,23,4": None}
    comma = issues(Workbook.open(data, locale=Locale(decimal=",", thousands=".", day_order="dmy")))
    assert comma["1.234,50"] == "number_as_text" and comma["1,234.50"] is None
    assert comma["05/01/2026"] == "date_as_text" and comma["31/02/2026"] is None
    assert comma["2026-01-05"] == "date_as_text"
    assert Workbook.open(data).cell("S!A2").issue is None
    assert all(v.t == "text" for v in next(Workbook.open(data).read("S!A1:L1")))


def test_locale_settings_are_checked():
    assert Locale().decimal == "."
    assert Locale(language="ja-jp").language == "ja-JP"
    for bad in ({"decimal": ""}, {"decimal": "1"}, {"decimal": ",", "thousands": ","},
                {"day_order": "dym"}, {"language": "en-US"}):
        with pytest.raises(ValueError):
            Locale(**bad)
    assert looks_like_number("1 234,5", Locale(decimal=",", thousands=" "))
    assert not looks_like_number("", Locale())
    assert looks_like_date("1/31/26", Locale(day_order="mdy"))
    assert not looks_like_date("2026-02-30", Locale())


def test_cells_at_the_length_limit_are_flagged():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["x" * TEXT_LIMIT, Inline("y" * TEXT_LIMIT), "z" * (TEXT_LIMIT - 1)]])
    wb = _open(book)
    row = next(wb.read("S!A1:C1"))
    assert [v.at_limit for v in row] == [True, True, False]
    assert len(row[0].v) == TEXT_LIMIT
    assert Problem("at_limit", "S!A1", 2) in wb.problems


def test_rows_and_cells_without_r_are_placed_as_excel_places_them():
    book = WorkbookBuilder()
    sheet = book.sheet("Sparse")
    sheet.omit_refs = True
    sheet.rows("A1", [[1, 2, 3], [4, None, 6]])
    sheet.cell("C1000", "far")
    wb = _open(book)
    assert [v.a for v in wb.cells("Sparse")] == ["Sparse!A1", "Sparse!B1", "Sparse!C1",
                                                  "Sparse!A2", "Sparse!C2", "Sparse!C1000"]
    data = _with_sheet(book, '<row><c><v>1</v></c><c r="C1"><v>3</v></c><c><v>4</v></c></row>'
                             '<row><c><v>5</v></c></row><row r="7"><c r="B7"><v>7</v></c><c><v>8</v></c></row>'
                             '<row><c t="s"><v>0</v></c></row>')
    wb = Workbook.open(data)
    assert [(v.ref, v.v) for v in wb.cells("Sparse")] == [("A1", 1), ("C1", 3), ("D1", 4), ("A2", 5),
                                                          ("B7", 7), ("C7", 8), ("A8", "far")]


def test_rows_and_cells_out_of_order_read_in_place():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", 1)
    data = _with_sheet(book, '<row r="3"><c r="B3"><v>32</v></c><c r="A3"><v>31</v></c></row>'
                             '<row r="1"><c r="A1"><v>11</v></c></row>'
                             '<row r="2"><c r="C2"><v>23</v></c></row>')
    wb = Workbook.open(data)
    assert not wb.scan("S").ordered
    assert _grid(wb, "S!A1:C3") == [[11, None, None], [None, None, 23], [31, 32, None]]
    assert wb.used_range("S").ref == "A1:C3"
    assert wb.scan(wb.sheets[0]).ordered is False
    assert _open(_values_book()).scan("Values").ordered


def test_a_wrong_or_missing_dimension_is_ignored():
    book = WorkbookBuilder()
    sheet = book.sheet("Wrong")
    sheet.dimension = "A1:B2"
    sheet.rows("B2", [[1, 2, 3], [4, 5, 6]])
    sheet.cell("C1000", 7)
    missing = book.sheet("Missing")
    missing.dimension = False
    missing.rows("C3", [[1]])
    empty = book.sheet("Empty")
    empty.dimension = "A1:Z99"
    wb = _open(book)
    assert wb.scan("Wrong").dimension == "A1:B2"
    assert wb.used_range("Wrong").ref == "B2:D1000"
    assert wb.used_range("Missing").ref == "C3"
    assert wb.scan("Missing").dimension is None
    assert wb.used_range("Empty") is None
    assert list(wb.read("Empty!A:C")) == []
    assert _grid(wb, "Empty!A1:B1") == [[None, None]]
    assert _grid(wb, "Wrong!A:B")[:2] == [[1], [4]]
    assert len(_grid(wb, "Wrong!A:B")) == 999  # rows 2 to 1000, column B (clipped from A:B)


def test_whole_columns_and_rows_are_clipped_to_the_used_range():
    book = WorkbookBuilder()
    book.sheet("S").rows("B2", [[1, 2], [3, 4], [5, 6]])
    wb = _open(book)
    assert [[v.a for v in row] for row in wb.read("S!A:B")] == [["S!B2"], ["S!B3"], ["S!B4"]]
    assert _grid(wb, "S!3:3") == [[3, 4]]
    assert _grid(wb, "S!2:9") == [[1, 2], [3, 4], [5, 6]]
    assert list(wb.read("S!E:F")) == []
    assert list(wb.read("S!7:9")) == []
    assert _grid(wb, "S!A1:A2") == [[None], [None]]  # a rectangle is read as given


def _million(book: WorkbookBuilder) -> bytes:
    """A sheet of one million sparse rows: a value every 50,000th row, and in the last."""
    rows = "".join(f'<row r="{r}"><c r="B{r}"><v>{r}</v></c></row>'
                   for r in [1, *range(50_000, 1_048_576, 50_000), 1_048_576])
    return _with_sheet(book, rows, before='<dimension ref="A1"/>')


def test_a_sheet_of_a_million_sparse_rows():
    book = WorkbookBuilder()
    book.sheet("Tall").cell("A1", 1)
    wb = Workbook.open(_million(book))
    assert wb.used_range("Tall").ref == "B1:B1048576"
    assert wb.scan("Tall").cells == 22
    rows = wb.read("Tall!B:B")
    first = [row[0].v for row in itertools.islice(rows, 3)]
    assert first == [1, None, None]
    assert wb.cell("Tall!B1048576").v == 1_048_576
    assert _grid(wb, "Tall!1048576:1048576") == [[1_048_576]]


def test_error_values_keep_their_code():
    book = WorkbookBuilder()
    codes = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]
    book.sheet("S").rows("A1", [[Error(code) for code in codes]])
    row = next(_open(book).read("S!A1:G1"))
    assert [(v.v, v.t) for v in row] == [(code, "error") for code in codes]


def test_merges_are_listed():
    book = WorkbookBuilder()
    sheet = book.sheet("S")
    sheet.rows("A1", [["Region", None, "Q1"], ["North", "x", 3]])
    sheet.merge("A1:B1")
    sheet.merge("C1:C2")
    wb = _open(book)
    assert [area.ref for area in wb.merges("S")] == ["A1:B1", "C1:C2"]
    assert wb.merges(wb.sheets[0]) == wb.merges("S")
    assert _open(_values_book()).merges("Values") == []


def test_strict_namespaces_are_read_like_transitional():
    book = _values_book()
    parts = {}
    for name, data in book.parts().items():
        data = data.replace(NS_MAIN.encode(), b"http://purl.oclc.org/ooxml/spreadsheetml/main")
        data = data.replace(NS_R.encode(), b"http://purl.oclc.org/ooxml/officeDocument/relationships")
        parts[name] = data
    strict = Workbook.open(zip_parts(parts))
    transitional = _open(_values_book())
    assert [[v.as_dict() for v in row] for row in strict.read("Values!A1:D6")] == \
        [[v.as_dict() for v in row] for row in transitional.read("Values!A1:D6")]


def test_entities_in_a_sheet_are_never_expanded():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", 1)
    entities = '<!ENTITY a0 "lol">' + "".join(
        f'<!ENTITY a{n} "{("&a%d;" % (n - 1)) * 10}">' for n in range(1, 10))
    parts = book.parts()
    parts["xl/worksheets/sheet1.xml"] = (
        f'<?xml version="1.0"?><!DOCTYPE worksheet [{entities}]>'
        f'<worksheet xmlns="{NS_MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>&a9;</t>'
        f"</is></c></row></sheetData></worksheet>").encode()
    with pytest.raises(OpenError) as refused:
        wb = Workbook.open(zip_parts(parts))
        list(wb.cells("S"))
    assert refused.value.code == "damaged"
    assert "lol" not in str(refused.value)


def test_a_damaged_sheet_is_refused_by_name_at_read():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", MARKER)
    parts = book.parts()
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"][:-30] + MARKER.encode()
    try:
        wb = Workbook.open(zip_parts(parts))
    except OpenError as error:  # refused at open by the protection scan: fine too
        assert MARKER not in str(error)
        return
    for attempt in (lambda: wb.scan("S"), lambda: list(wb.cells("S")), lambda: list(wb.read("S!A1:B2"))):
        with pytest.raises(OpenError) as refused:
            attempt()
        assert refused.value.code == "damaged" and MARKER not in str(refused.value)
        assert "xl/worksheets/sheet1.xml" in str(refused.value)


# -- addresses -----------------------------------------------------------------------------------


def test_values_are_addressed_as_addresses_write_them():
    book = WorkbookBuilder()
    book.sheet("Q3 sales", sheet_id=3).cell("B2", 9)
    book.sheet("Bob's").cell("A1", 1)
    wb = _open(book)
    assert wb.cell("'Q3 sales'!B2").a == "'Q3 sales'!B2"
    assert wb.cell("#3!B2").a == "'Q3 sales'!B2"
    assert wb.cell("'Bob''s'!A1").as_dict() == {"a": "'Bob''s'!A1", "v": 1, "t": "number"}
    assert wb.cell("'Q3 sales'!Z99").as_dict() == {"a": "'Q3 sales'!Z99", "v": None, "t": "blank"}
    with pytest.raises(NameNotFound):
        wb.cell("Nowhere!A1")
    with pytest.raises(ValueError, match="later milestone"):
        list(wb.read("d:by_category"))
    with pytest.raises(NameNotFound):  # a t: address is resolved (R4): there is no such table
        list(wb.read("t:Products"))
    with pytest.raises(ValueError, match="one cell"):
        wb.cell("'Q3 sales'!A1:B2")


def test_no_cell_content_in_problems():
    book = WorkbookBuilder(calc={"fullCalcOnLoad": "1"})
    sheet = book.sheet("S")
    sheet.rows("A1", [[MARKER, "12", Formula(f'"{MARKER}"&NOW()'), MARKER * 5000]])
    sheet.cell("A2", Formula(f'IF(A1="{MARKER}",1,0)'))
    wb = _open(book)
    assert wb.problems
    assert MARKER not in repr(wb.problems)
    assert MARKER not in repr(wb.scan("S"))
    # a value's formula text is content: it is kept out of the value's repr
    assert MARKER not in repr(wb.cell("S!A2"))


# -- laziness and bounded memory -----------------------------------------------------------------


def _big(rows: int) -> bytes:
    """A made-up stock list of ``rows`` rows: a number, a date and a shared string each."""
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    for i in range(100):
        sheet.cell(f"C{i + 1}", f"item-{i:03d}")
    day = book.style(num_fmt="yyyy-mm-dd")
    body = "".join(f'<row r="{r}"><c r="A{r}"><v>{r * 1.5}</v></c><c r="B{r}" s="{day}">'
                   f'<v>{40000 + r % 3000}</v></c><c r="C{r}" t="s"><v>{r % 100}</v></c></row>'
                   for r in range(1, rows + 1))
    return _with_sheet(book, body)


def _count_values(monkeypatch) -> list[int]:
    calls = [0]
    original = SheetReader.value

    def counted(self, *args):
        calls[0] += 1
        return original(self, *args)

    monkeypatch.setattr(SheetReader, "value", counted)
    return calls


def test_reading_a_range_or_a_whole_column_is_lazy(monkeypatch):
    wb = Workbook.open(_big(5000))
    wb.scan("Stock")  # once per sheet, kept
    calls = _count_values(monkeypatch)
    rows = wb.read("Stock!A10:B12")
    assert calls[0] == 0  # nothing is read before it is asked for
    assert [[v.v for v in row] for row in rows] == [
        [r * 1.5, serial_text(40000 + r, "date", False)] for r in (10, 11, 12)]
    assert calls[0] == 6  # the cells of A10:B12, and no other
    calls[0] = 0
    column = wb.read("Stock!C:C")
    first = [row[0].v for row in itertools.islice(column, 3)]
    assert first == ["item-001", "item-002", "item-003"]
    assert calls[0] == 3
    assert wb.used_range("Stock").ref == "A1:C5000"


def _peak(data: bytes, address: str) -> tuple[int, int, float]:
    """(rows read, the peak of Python allocations while reading, seconds)."""
    wb = Workbook.open(data)
    tracemalloc.start()
    try:
        started = time.perf_counter()
        count = sum(1 for _ in wb.read(address))
        elapsed = time.perf_counter() - started
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    return count, peak, elapsed


def test_memory_does_not_grow_with_the_rows():
    small = _peak(_big(1_500), "Stock!A:C")
    large = _peak(_big(15_000), "Stock!A:C")
    assert (small[0], large[0]) == (1_500, 15_000)
    # ten times the rows, nearly the same peak: what is held is a row, the shared strings
    # and a bounded cache, never the sheet
    assert large[1] < small[1] * 1.5 + 64_000
    assert large[1] < 1_000_000


_RSS_SCRIPT = """
import resource, sys, time
from xlsx_agent import Workbook
def rss():
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak if sys.platform == "darwin" else peak * 1024
wb = Workbook.open(sys.argv[1])
wb.package.read("xl/worksheets/sheet1.xml")
before = rss()
started = time.perf_counter()
wb.scan("Stock")
scanned = time.perf_counter() - started
started = time.perf_counter()
count = sum(1 for _ in wb.read("Stock!A:C"))
read = time.perf_counter() - started
print(count, rss() - before, round(scanned, 1), round(read, 1))
"""


@pytest.mark.slow
def test_a_200000_row_sheet_reads_in_bounded_memory(tmp_path):
    data = _big(200_000)
    # Python's allocations, which tracemalloc sees
    count, peak, elapsed = _peak(data, "Stock!A:C")
    assert count == 200_000
    assert peak < 1_000_000, peak  # the sheet part alone is about 25 MB
    print(f"\n200,000 rows: peak {peak / 1e6:.2f} MB of Python allocations, {elapsed:.1f} s "
          "(scan and read, under tracemalloc)")
    # and the parser's, which it does not: the growth of the process's peak while streaming
    # (the tree of the whole sheet would be some hundreds of MB)
    if sys.platform == "win32":
        return
    path = tmp_path / "stock.xlsx"
    path.write_bytes(data)
    out = subprocess.run([sys.executable, "-c", _RSS_SCRIPT, str(path)], check=True,
                         capture_output=True, text=True).stdout.split()
    rows, growth, scanned, read = int(out[0]), int(out[1]), out[2], out[3]
    assert rows == 200_000
    assert growth < 40_000_000, growth
    print(f"200,000 rows: peak resident memory grew {growth / 1e6:.1f} MB while streaming; "
          f"scan {scanned} s, read {read} s")
