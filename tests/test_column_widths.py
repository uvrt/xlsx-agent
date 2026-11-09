"""Column widths read back (production R76, :mod:`xlsx_agent.read.widths`).

The expected widths are Excel's own (Excel for Mac 16.106, 2026-11-09): made-up workbooks of
these sheet formats and columns under each default font, every column's ``column width``
read through Excel's object model.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError

#: The sheet formats measured, by sheet name.
FORMATS = {
    "nobase": '<sheetFormatPr defaultRowHeight="15"/>',
    "base10": '<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>',
    "def12": '<sheetFormatPr defaultColWidth="12" defaultRowHeight="15"/>',
    "noformat": "",
    "base10def9": '<sheetFormatPr baseColWidth="10" defaultColWidth="9" defaultRowHeight="15"/>',
    "base4": '<sheetFormatPr baseColWidth="4" defaultRowHeight="15"/>',
}
#: B its own width, C hidden with one, D hidden without, E a width Windows writes, F zero.
COLS = ('<cols><col min="2" max="2" width="12.83203125" customWidth="1"/>'
        '<col min="3" max="3" width="20" hidden="1" customWidth="1"/>'
        '<col min="4" max="4" hidden="1"/>'
        '<col min="5" max="5" width="9.140625" customWidth="1"/>'
        '<col min="6" max="6" width="0" customWidth="1"/></cols>')

#: What Excel showed for column A (the default), B and E, by font and sheet format.
MEASURED = {
    ("Aptos Narrow", 11): {"nobase": 8.0, "base10": 10.0, "def12": 11.17, "noformat": 10.0,
                           "base10def9": 8.17, "base4": 4.0, "B": 12.0, "E": 8.33},
    ("Aptos Narrow", 12): {"nobase": 8.0, "base10": 10.0, "def12": 11.17, "noformat": 10.0,
                           "base10def9": 8.17, "base4": 4.0, "B": 12.0, "E": 8.33},
    ("Calibri", 11): {"nobase": 8.0, "base10": 10.0, "def12": 11.17, "noformat": 10.0,
                      "base10def9": 8.17, "base4": 4.0, "B": 12.0, "E": 8.33},
    ("Calibri", 16): {"nobase": 8.0, "base10": 10.0, "def12": 11.38, "noformat": 10.0,
                      "base10def9": 8.38, "base4": 4.0, "B": 12.25, "E": 8.5},
}


def book(font: str = "Aptos Narrow", size: float = 11, sheets=FORMATS, cols: str = COLS) -> bytes:
    data = Workbook.new(list(sheets)).to_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            part = source.read(info.filename).decode()
            if info.filename == "xl/styles.xml":
                part = re.sub(r'<sz val="[0-9.]+"/>', f'<sz val="{size}"/>', part, count=1)
                part = re.sub(r'<name val="[^"]+"/>', f'<name val="{font}"/>', part, count=1)
            number = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", info.filename)
            if number:
                key = list(sheets)[int(number.group(1)) - 1]
                part = re.sub(r"<sheetFormatPr[^>]*/>", sheets[key] + cols, part)
            target.writestr(info, part.encode())
    return out.getvalue()


@pytest.mark.parametrize("font", list(MEASURED))
@pytest.mark.parametrize("sheet", list(FORMATS))
def test_widths_are_what_excel_shows(font, sheet):
    wb = Workbook.open(book(*font))
    read = wb.column_widths(f"{sheet}!A:G")
    expected = MEASURED[font]
    a, b, c, d, e, f, g = read.columns
    assert read.default_width == expected[sheet] == a.width == g.width
    assert a.default and a.stored is None and not a.hidden
    assert (b.width, b.stored, b.default) == (expected["B"], 12.83203125, False)
    assert (e.width, e.stored) == (expected["E"], 9.140625)
    # Hidden: Excel shows 0 while hidden and keeps the width it shows with again.
    assert c.hidden and c.stored == 20 and c.width == pytest.approx(
        19.17 if font[1] != 16 else 19.38)
    assert d.hidden and d.default and d.width == expected[sheet]
    assert f.hidden and f.width == 0 and f.stored == 0


def test_a_sheet_alone_reads_up_to_its_last_used_or_sized_column():
    wb = Workbook.new(["Stock"])
    wb.set_cells({"Stock!B2": "count"})
    wb.set_column_widths("Stock!E:E", 15)
    read = wb.column_widths("Stock")
    assert [c.column for c in read.columns] == ["A", "B", "C", "D", "E"]
    assert read["E"].width == 15 and read["A"].default
    assert wb.column_widths("Stock!").columns == read.columns


def test_one_column_a_range_and_a_quoted_sheet_name():
    wb = Workbook.new(["Q3 sales ", "Q3 sales"])
    wb.set_column_widths("'Q3 sales '!C:C", 20)
    assert wb.column_widths("'Q3 sales '!C")["C"].width == 20
    assert [c.column for c in wb.column_widths("'Q3 sales '!B2:D9").columns] == ["B", "C", "D"]
    assert wb.column_widths("'Q3 sales'!C")["C"].default        # the twin is its own sheet


def test_widths_read_after_an_edit_and_after_a_save_agree():
    wb = Workbook.new(["Stock"])
    wb.set_column_widths("Stock!B:C", 14)
    wb.hide_columns("Stock", ["C"])
    edited = wb.column_widths("Stock!A:D")
    again = Workbook.open(wb.to_bytes()).column_widths("Stock!A:D")
    assert edited == again
    assert [(c.width, c.hidden) for c in again.columns] == [(8.0, False), (14.0, False),
                                                            (14.0, True), (8.0, False)]


def test_more_columns_than_the_limit_are_counted_and_the_next_address_given():
    """Tool design rule 4: a size budget, and nothing left out silently."""
    wb = Workbook.new(["Stock"])
    read = wb.column_widths("Stock!A:Z", limit=10)
    assert len(read.columns) == 10 and read.more == 16 and read.next == "Stock!K:Z"
    assert "16 more column(s)" in read.notes[0]
    rest = wb.column_widths(read.next, limit=26)
    assert rest.more == 0 and [c.column for c in rest.columns][0] == "K"
    assert len(wb.column_widths("Stock!A:XFD", limit=16384).columns) == 16384


@pytest.mark.parametrize("target, limit, code", [
    ("Stock!A1:", 256, "bad_target"), ("Nowhere!A:C", 256, "bad_target"), ("", 256, "bad_target"),
    ("Stock!A:C", 0, "bad_value"), ("Stock!A:C", 16385, "bad_value"), (None, 256, "bad_target"),
])
def test_refusals_say_what_is_taken(target, limit, code):
    """Tool design rule 2: the refusal names the forms taken (and the sheets there are)."""
    wb = Workbook.new(["Stock"])
    with pytest.raises(EditError) as refused:
        wb.column_widths(target, limit=limit)
    assert refused.value.code == code
    if target == "Nowhere!A:C":
        assert "Stock" in str(refused.value)


def test_an_unmeasured_default_font_gives_stored_widths_and_says_why():
    wb = Workbook.open(book("Made Up Sans", 11, {"base10": FORMATS["base10"]}))
    read = wb.column_widths("base10!A:B")
    assert read.digit is None and read.default_width is None
    assert [(c.width, c.stored) for c in read.columns] == [(None, None), (None, 12.83203125)]
    assert "Made Up Sans" in read.notes[0]


def test_as_dict_says_which_width_is_stored():
    wb = Workbook.open(book())
    shown = wb.column_widths("nobase!A:C").as_dict()
    assert shown["columns"] == [
        {"column": "A", "width": 8.0, "source": "default"},
        {"column": "B", "width": 12.0, "source": "column", "stored": 12.83203125},
        {"column": "C", "width": 19.17, "source": "column", "stored": 20.0, "hidden": True}]
