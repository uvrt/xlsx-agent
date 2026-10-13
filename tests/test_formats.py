"""Number formats: the built-in ids of ECMA-376 Part 1, §18.8.30, and format codes
classified by their tokens (§18.8.31)."""

import re

import pytest

from build import WorkbookBuilder, zip_parts
from xlsx_agent import Locale, Workbook
from xlsx_agent.read.styles import (BUILTIN_FORMATS, LANGUAGE_FORMATS, builtin_format,
                                    classify_format)

# What each id of §18.8.30's all-languages table is, read from its code by hand.
ALL_LANGUAGES = {
    0: "general", 1: "number", 2: "number", 3: "number", 4: "number", 9: "percent",
    10: "percent", 11: "number", 12: "number", 13: "number", 14: "date", 15: "date",
    16: "date", 17: "date", 18: "time", 19: "time", 20: "time", 21: "time", 22: "datetime",
    37: "number", 38: "number", 39: "number", 40: "number", 45: "time", 46: "time",
    47: "time", 48: "number", 49: "text",
}


def test_every_builtin_id_of_the_all_languages_table():
    assert set(BUILTIN_FORMATS) == set(ALL_LANGUAGES)
    for number, kind in ALL_LANGUAGES.items():
        found = builtin_format(number)
        assert found.kind == kind, number
        assert found.code == BUILTIN_FORMATS[number] and found.builtin
    assert builtin_format(46).elapsed and not builtin_format(21).elapsed
    # whatever language the caller names, these are the same
    for language in LANGUAGE_FORMATS:
        assert builtin_format(14, language).kind == "date"


@pytest.mark.parametrize("number", [5, 6, 7, 8, 23, 24, 25, 26, 41, 42, 43, 44, 63, 82, 163])
def test_an_id_the_standard_does_not_list_is_unknown(number):
    assert builtin_format(number).kind == "unknown"
    assert builtin_format(number).code is None


def test_language_dependent_ids_without_a_language():
    kinds = {number: builtin_format(number).kind for number in [*range(27, 37), *range(50, 59)]}
    assert {n for n, k in kinds.items() if k == "date"} == {27, 28, 29, 30, 31, 36, 50, 51, 54, 57, 58}
    assert {n for n, k in kinds.items() if k == "time"} == {32, 33}
    # a time in one language and a date in another: never guessed
    assert {n for n, k in kinds.items() if k == "ambiguous"} == {34, 35, 52, 53, 55, 56}
    assert builtin_format(30).code is None  # the same kind, but a different code per language
    # Thai ids: only th-TH lists them
    thai = {n: builtin_format(n).kind for n in LANGUAGE_FORMATS["th-TH"]}
    assert thai == {59: "number", 60: "number", 61: "number", 62: "number", 67: "percent",
                    68: "percent", 69: "number", 70: "number", 71: "date", 72: "date",
                    73: "date", 74: "date", 75: "time", 76: "time", 77: "datetime",
                    78: "time", 79: "time", 80: "time", 81: "date"}
    assert builtin_format(79).elapsed


def test_language_dependent_ids_in_a_language():
    assert builtin_format(34, "ja-JP").kind == "date"
    assert builtin_format(34, "zh-TW").kind == "time"
    assert builtin_format(52, "zh-TW").kind == "time"
    assert builtin_format(52, "zh-CN").kind == "date"
    assert builtin_format(55, "ko-KR").kind == "date"
    assert builtin_format(27, "ja-JP").code == "[$-411]ge.m.d"
    assert builtin_format(27, "ja-JP").lcid == 0x411
    # an id another language lists is unknown in this one
    assert builtin_format(59, "ja-JP").kind == "unknown"
    assert builtin_format(27, "th-TH").kind == "unknown"


@pytest.mark.parametrize("code, kind", [
    ("General", "general"),
    ("general", "general"),
    ("0", "number"),
    ("#,##0.00", "number"),
    ("0.00E+00", "number"),
    ("##0.0e-0", "number"),
    ("[Red]0.00;[Blue]-0.00", "number"),
    ('"Total: "0', "number"),
    ('0" days"', "number"),
    ("0\\d", "number"),
    ("0_);(0)", "number"),
    ("#,##0.00 *-", "number"),
    ("[<100]0;0.00", "number"),
    ("0.0%", "percent"),
    ('0"%"', "number"),
    ("0\\%", "number"),
    ("0%;[Red]-0%", "percent"),
    ("@", "text"),
    ('"Code: "@', "text"),
    ("yyyy-mm-dd", "date"),
    ("YYYY-MM-DD", "date"),
    ("d-mmm", "date"),
    ("m/d", "date"),
    ("mmmm", "date"),
    ("dd/mm/yyyy hh:mm", "datetime"),
    ("h:mm", "time"),
    ("hh:mm:ss.000", "time"),
    ("mm:ss", "time"),
    ("m:ss", "time"),
    ("h AM/PM", "time"),
    ("h:mm a/p", "time"),
    ("[h]:mm", "time"),
    ("[mm]:ss", "time"),
    ("[$-409]mmmm d, yyyy", "date"),
    ("[$-F800]dddd, mmmm dd, yyyy", "date"),
    ("[$-F400]h:mm:ss AM/PM", "time"),
    ("[$€-407] #,##0.00", "number"),
    ("ggge", "date"),
    ("ว/ด/ปปปป", "date"),
    ("ช:นน", "time"),
    ("d/m/bb", "date"),
    ('"yyyy"0', "number"),
    ("", "number"),
])
def test_format_codes_are_classified_by_their_tokens(code, kind):
    assert classify_format(code)[0] == kind


def test_elapsed_time_and_locale_tags():
    assert classify_format("[h]:mm:ss") == ("time", True, None)
    assert classify_format("h:mm:ss") == ("time", False, None)
    assert classify_format("[$-409]h:mm")[2] == 0x409
    assert classify_format("[$-F800]dddd")[2] == 0xF800
    assert classify_format("[Red]0")[2] is None


def test_a_cells_format_comes_from_its_style():
    book = WorkbookBuilder()
    sheet = book.sheet("Prices")
    styles = {
        "A1": book.style(num_fmt_id=9),        # built-in 0%
        "A2": book.style(num_fmt="0.0%"),      # the workbook's own
        "A3": book.style(num_fmt_id=14),       # built-in date
        "A4": book.style(num_fmt_id=5),        # not in the standard's list, no code
        "A5": book.style(num_fmt_id=34),       # a time or a date, by language
        "A6": book.style(num_fmt="#,##0.00"),
    }
    for ref, style in styles.items():
        sheet.cell(ref, 46027 if ref in ("A3", "A5") else 0.25, style=style)
    data = book.to_bytes()
    wb = Workbook.open(data)
    got = {value.ref: (value.t, value.v, value.issue) for value in wb.cells("Prices")}
    assert got == {"A1": ("percent", 0.25, None), "A2": ("percent", 0.25, None),
                   "A3": ("date", "2026-01-05", None),
                   "A4": ("number", 0.25, "unknown_format"),
                   "A5": ("number", 46027, "ambiguous_format"),
                   "A6": ("number", 0.25, None)}
    assert wb.number_format(styles["A2"]).code == "0.0%"
    assert wb.number_format(999).kind == "general"  # past the end of cellXfs
    formats = {value.ref: value.format for value in wb.cells("Prices", formats=True)}
    assert formats == {"A1": "0%", "A2": "0.0%", "A3": "mm-dd-yy", "A4": None, "A5": None,
                       "A6": "#,##0.00"}
    # a language named by the caller reads the language-dependent id
    ja = Workbook.open(data, locale=Locale(language="ja-jp"))
    assert ja.cell("Prices!A5").as_dict() == {"a": "Prices!A5", "v": "2026-01-05", "t": "date",
                                              "serial": 46027.0}
    assert Workbook.open(data, locale=Locale(language="zh-TW")).cell("Prices!A5").t == "time"


def test_a_workbook_without_styles_is_general_throughout():
    book = WorkbookBuilder()
    book.sheet("S").cell("A1", 3)
    parts = book.parts()
    del parts["xl/styles.xml"]
    parts["xl/_rels/workbook.xml.rels"] = re.sub(
        rb'<Relationship [^>]*/styles"[^>]*/>', b"", parts["xl/_rels/workbook.xml.rels"])
    wb = Workbook.open(zip_parts(parts))
    assert wb.cell("S!A1").as_dict() == {"a": "S!A1", "v": 3, "t": "number"}
    assert wb.number_format(0).kind == "general"
