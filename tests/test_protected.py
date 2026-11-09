"""Protected sheets (round 46, R71): a protected sheet's unlocked cells are written as Excel
lets them be typed into, its locked cells refused with the way out, its protection never
touched; the other calls follow the protection's options as Excel does.

Made-up workbooks only: a staff rota protected as a template's owner would, with a password
hash, salt and spin count, and some cells unlocked for input -- by their own format, by
their column's, by their row's.  What Excel allows was measured in Excel for Mac 16.106
(2026-11-09; :mod:`xlsx_agent.edit.protection`)."""

from __future__ import annotations

import datetime as dt
import io
import re
import zipfile

import pytest

from ooxml_edit.sheets import Cell

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError

from build import WorkbookBuilder, zip_parts

#: A protection as Excel writes one with a password (made-up hash and salt).
GUARD = ('<sheetProtection algorithmName="SHA-512" hashValue="bWFkZS11cCBoYXNoIGZvciB0ZXN0'
         'cyBvbmx5IG1hZGUtdXAgaGFzaCBmb3IgdGVzdHMgb25seSBtYWRlLXVwIGhhc2ggZm9yIHRlc3RzIG9u'
         'bHk=" saltValue="c2FsdCBmb3IgdGVzdHM=" spinCount="100000" sheet="1" objects="1" '
         'scenarios="1"{extra}/>')

STYLES = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><protection locked="0"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>
</cellXfs>
<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Input" xfId="1"/></cellStyles>
</styleSheet>'''

LOCKED, UNLOCKED, UNLOCKED_NO_APPLY, STYLE_ONLY = 0, 1, 2, 3


def _book(extra: str = "", note: bool = False) -> bytes:
    """``Rota``: A1:D1 headers (locked); row 2: A2 locked, B2 unlocked, C2 unlocked without
    ``applyProtection``, D2 only its cell style unlocked (locked, as Excel reads it); column
    E unlocked by its format, E4 a locked cell in it; row 6 unlocked by its format; B3:C4
    unlocked.  ``Plain`` is not protected."""
    book = WorkbookBuilder()
    rota = book.sheet("Rota")
    for column, text in zip("ABCD", ("Name", "Mon", "Tue", "Wed")):
        rota.cell(f"{column}1", text)
    rota.cell("A2", "Ann", style=LOCKED)
    rota.cell("B2", 8, style=UNLOCKED)
    rota.cell("C2", 7, style=UNLOCKED_NO_APPLY)
    rota.cell("D2", 6, style=STYLE_ONLY)
    for ref in ("B3", "C3", "B4", "C4"):
        rota.cell(ref, None, style=UNLOCKED)
    rota.cell("E4", "fixed", style=LOCKED)
    rota.cell("A7", "end", style=LOCKED)
    if note:
        rota.comment("A7", "kept")
    book.sheet("Plain").cell("A1", "x")
    parts = book.parts()
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    sheet = sheet.replace("<sheetData>", '<cols><col min="5" max="5" width="9" style="1"/>'
                          "</cols><sheetData>")
    sheet = sheet.replace('<row r="7">', '<row r="6" s="1" customFormat="1"/><row r="7">')
    sheet = sheet.replace("</sheetData>", "</sheetData>" + GUARD.format(extra=extra))
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    parts["xl/styles.xml"] = STYLES
    return zip_parts(parts)


def _sheet_xml(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode()


def _guard(data: bytes) -> str:
    return re.search(r"<sheetProtection[^>]*/>", _sheet_xml(data)).group(0)


def _refused(wb: Workbook, call, *args, **options) -> EditError:
    before, version = wb.to_bytes(), wb.version
    with pytest.raises(EditError) as raised:
        call(*args, **options)
    assert wb.to_bytes() == before and wb.version == version
    return raised.value


# -- cells --------------------------------------------------------------------------------------


def test_unlocked_cells_are_written_and_the_protection_kept_byte_for_byte():
    data = _book()
    wb = Workbook.open(data)
    written = wb.set_cells({"Rota!B2": 9, "Rota!C2": {"formula": "=B2-1"},
                            "Rota!E20": "extra", "Rota!F6": "row six", "Rota!C4": None,
                            "Rota!B3": "x"})
    assert written.changed == 5          # C4 was blank already
    after = wb.to_bytes()
    assert _guard(after) == _guard(data)              # hash, salt, spin count and options
    assert wb.sheet("Rota").protected
    assert (wb.cell("Rota!B2").v, wb.cell("Rota!E20").v, wb.cell("Rota!F6").v) == \
        (9, "extra", "row six")


def test_a_locked_cell_is_refused_naming_the_unlocked_cells_and_the_way_out():
    wb = Workbook.open(_book())
    # A2 locked by its format, D2 locked though its cell style is not (Excel reads the
    # cell's format only), E4 locked in an unlocked column, G9 a new cell nothing unlocks.
    error = _refused(wb, wb.set_cells, {"Rota!B2": 1, "Rota!A2": "Bo", "Rota!D2": 1,
                                        "Rota!E4": 1, "Rota!G9": 1})
    assert error.code == "protected"
    assert error.addresses == ("Rota!A2", "Rota!D2", "Rota!E4", "Rota!G9")
    text = str(error)
    for shown in ("Rota!E:E", "Rota!6:6", "Rota!B2:C4"):
        assert shown in text, shown
    assert "sheetProtection" in text and "wb.unprotect_sheet" in text


def test_no_policy_or_call_lifts_a_locked_cell():
    wb = Workbook.open(_book(), policy={"cells": ["Rota!"]})
    assert _refused(wb, wb.set_cells, {"Rota!A2": "Bo"}).code == "protected"
    assert wb.set_cells({"Rota!B2": 1}).changed == 1
    with pytest.raises(EditError) as raised:
        wb.editor.write(wb.sheet("Rota"), Cell(2, 1), "Bo")
    assert raised.value.code == "protected"
    assert wb.editor.write(wb.sheet("Rota"), Cell(2, 2), 2)


def test_an_unlocked_cell_takes_a_date_s_and_a_line_break_s_format_as_in_excel():
    # Measured: on a protected sheet that allows no formatting, Excel gave an unlocked
    # General cell the short date, m/d/yyyy h:mm, h:mm and a wrapping format for what was
    # typed there.
    wb = Workbook.open(_book())
    written = wb.set_cells({"Rota!B3": dt.date(2026, 1, 5), "Rota!C3": dt.datetime(2026, 1, 5, 9),
                            "Rota!B4": dt.time(9, 30), "Rota!C4": {"markdown": "early\nlate"}})
    assert written.short_date == ("Rota!B3",)
    assert written.time_format == ("Rota!C3", "Rota!B4")
    assert wb.cell("Rota!B4").v == "09:30:00" and wb.cell("Rota!C4").v == "early\nlate"
    assert wb.sheet("Rota").protected


# -- blocks of rows -----------------------------------------------------------------------------


def test_a_block_is_written_into_unlocked_cells_and_refused_over_locked_ones():
    data = _book()
    wb = Workbook.open(data)
    wb.write_table("Rota", "B3", None, [["a", "b"], ["c", "d"]])
    assert wb.cell("Rota!C4").v == "d"
    appended = wb.append_rows("Rota!E:E", [["one"], ["two"]])
    assert appended.range == "Rota!E5:E6"
    error = _refused(wb, wb.write_table, "Rota", "A8", None, [["x"]])
    assert error.code == "protected" and error.addresses == ("Rota!A8",)
    assert _guard(wb.to_bytes()) == _guard(data)


def test_a_table_is_not_made_on_a_protected_sheet():
    wb = Workbook.open(_book())
    error = _refused(wb, wb.write_table, "Rota", "B3", ["Mon", "Tue"], [["a", "b"]],
                     as_table="Week")
    assert error.code == "protected" and "without as_table" in str(error)


def test_a_table_does_not_grow_on_a_protected_sheet():
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A1", [["Item", "Qty"], ["Lamp", 3]])
    sheet.table("Stock", "A1:B2", ["Item", "Qty"], write_header=False)
    sheet.protected = True
    wb = Workbook.open(book.to_bytes())
    error = _refused(wb, wb.append_rows, "t:Stock", [["Fan", 4]])
    assert error.code == "protected" and "Stock!A3:B3" in str(error)


# -- notes --------------------------------------------------------------------------------------


def test_notes_follow_the_protection_s_objects_option():
    wb = Workbook.open(_book(note=True))
    error = _refused(wb, wb.set_comment, "Rota!B2", "late")
    assert error.code == "protected" and "objects" in str(error)
    # objects="0": Excel added notes to locked and unlocked cells alike.
    data = _book(note=True)
    data = _zip_replace(data, "xl/worksheets/sheet1.xml", b'objects="1" ', b"")
    wb = Workbook.open(data)
    assert wb.set_comment("Rota!A2", "locked cell").changed
    assert wb.set_comment("Rota!A7", "edited").changed
    assert _guard(wb.to_bytes()) == _guard(data)


def _zip_replace(data: bytes, part: str, old: bytes, new: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    assert old in parts[part]
    parts[part] = parts[part].replace(old, new)
    return zip_parts(parts)


# -- rows and columns ---------------------------------------------------------------------------

ALLOWING = ' insertRows="0" insertColumns="0" deleteRows="0" deleteColumns="0"'


def test_rows_and_columns_follow_the_protection_s_options():
    wb = Workbook.open(_book())
    for call, at in ((wb.insert_rows, 3), (wb.insert_columns, "B"), (wb.delete_rows, 6),
                     (wb.delete_columns, "E")):
        error = _refused(wb, call, "Rota", at)
        assert error.code == "protected" and "unprotect" in str(error), call
    assert "insertRows" in str(_refused(wb, wb.insert_rows, "Rota", 3))
    data = _book(ALLOWING)
    wb = Workbook.open(data)
    wb.insert_rows("Rota", 3)
    wb.insert_columns("Rota", "B")
    assert wb.cell("Rota!C2").v == 8 and wb.cell("Rota!A8").v == "end"     # moved
    assert _guard(wb.to_bytes()) == _guard(data)


def test_a_deletion_takes_lines_whose_every_cell_is_unlocked():
    data = _book(ALLOWING)
    wb = Workbook.open(data)
    # Row 6: no cells, its format unlocked -- every cell of it unlocked: deleted, as Excel did.
    wb.delete_rows("Rota", 6)
    assert wb.cell("Rota!A6").v == "end"
    # Row 2 holds locked cells; column E holds E4, locked under its unlocked format; column
    # F's positions have no unlocked format.
    for call, at, cell in ((wb.delete_rows, 2, "Rota!A2"), (wb.delete_columns, "E", "Rota!E4"),
                           (wb.delete_columns, "F", "Rota!F1"), (wb.delete_rows, 9, "Rota!A9")):
        error = _refused(wb, call, "Rota", at)
        assert error.code == "protected" and f"{cell} is locked" in str(error), (at, str(error))
    assert _refused(wb, wb.move_rows, "Rota", 6, before=2).code == "protected"


def test_the_unlocked_cells_a_refusal_lists_are_capped_and_say_so():
    from xlsx_agent.edit.protection import MAX_HINTS

    book = WorkbookBuilder()
    sheet = book.sheet("Grid")
    for n in range(1, 30, 2):          # unlocked cells on every other row: no block forms
        sheet.cell(f"B{n}", n)
    parts = book.parts()
    parts["xl/styles.xml"] = STYLES
    text = parts["xl/worksheets/sheet1.xml"].decode().replace("<c r=\"B", "<c s=\"1\" r=\"B")
    parts["xl/worksheets/sheet1.xml"] = text.replace(
        "</sheetData>", "</sheetData>" + GUARD.format(extra="")).encode()
    wb = Workbook.open(zip_parts(parts))
    message = str(_refused(wb, wb.set_cells, {"Grid!A1": 1}))
    listed = re.findall(r"Grid!B\d+", message)
    assert len(listed) == MAX_HINTS and "and more" in message
