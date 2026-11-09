"""Formatting and layout from code (E7, production round 44, R53): ``wb.format``,
``set_column_widths``, ``set_row_heights``, ``freeze_panes``, ``set_gridlines``,
``protect_sheet`` and ``unprotect_sheet``.

On made-up stock lists whose stylesheet names real fonts and cell styles
(:func:`stock_book`).  What Excel writes and does was measured (Excel for Mac 16.106,
2026-11-08; the module docstrings of :mod:`xlsx_agent.edit.format` and
:mod:`xlsx_agent.edit.layout`); the expectations below are those measurements.  The tool
design rules each test holds to are named in its docstring.
"""

from __future__ import annotations

import base64
import io
import zipfile

import pytest
from lxml import etree
from ooxml_common.locales import Locale

from build import WorkbookBuilder
from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.edit.layout import (Face, digit_width, fit_padding, fitted_width, password_hash,
                                    stored_width, text_pixels)

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _font(name, size, bold=False, italic=False, color=None):
    parts = ["<b/>" if bold else "", "<i/>" if italic else "", f'<sz val="{size}"/>',
             f'<color rgb="{color}"/>' if color else "", f'<name val="{name}"/>']
    return "<font>" + "".join(parts) + "</font>"


#: The fonts of the fit fixture's cell formats 1 to 11 (format 0 is the default font).
FONTS = [("Calibri", 11, True), ("Calibri", 14), ("Arial", 10), ("Times New Roman", 12),
         ("Courier New", 10), ("Aptos Narrow", 11), ("Aptos", 11), ("Cambria", 11),
         ("Arial", 14, True), ("Calibri", 8), ("Calibri", 20)]
#: Cell format 12: a number format, an italic font, thin borders, centred, unlocked.
PRE = 12


def styles_xml(default=("Calibri", 11)) -> bytes:
    """A stylesheet as the probe's: ``Heading 1`` (built-in, leaving out the number format,
    fill, alignment and protection), ``Stock heading`` (everything), ``Stock note`` (font
    and fill only), ``Stock locked`` (alignment and protection only)."""
    fonts = [_font(*default)] + [_font(*f) for f in FONTS] + [
        _font("Calibri", 15, True, color="FF1F3864"), _font("Calibri", 11, italic=True),
        _font("Calibri", 14, True)]
    n = len(fonts)
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    xfs += [f'<xf numFmtId="0" fontId="{i}" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
            for i in range(1, len(FONTS) + 1)]
    xfs.append(f'<xf numFmtId="164" fontId="{n - 2}" fillId="0" borderId="2" xfId="0" '
               'applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1" '
               'applyProtection="1"><alignment horizontal="center"/><protection locked="0"/></xf>')
    thin = '<{0} style="thin"><color auto="1"/></{0}>'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
<fonts count="{n}">{"".join(fonts)}</fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color rgb="FF4472C4"/></bottom><diagonal/></border><border>{thin.format("left")}{thin.format("right")}{thin.format("top")}{thin.format("bottom")}<diagonal/></border></borders>
<cellStyleXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="{n - 3}" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="{n - 1}" fillId="2" borderId="0"/><xf numFmtId="0" fontId="{n - 1}" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0"><alignment horizontal="right"/><protection locked="1" hidden="1"/></xf></cellStyleXfs>
<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>
<cellStyles count="5"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Kop 1" xfId="1" builtinId="16"/><cellStyle name="Stock heading" xfId="2"/><cellStyle name="Stock note" xfId="3"/><cellStyle name="Stock locked" xfId="4"/></cellStyles>
</styleSheet>'''.encode()


def stock_book(default=("Calibri", 11), protected=False) -> WorkbookBuilder:
    book = WorkbookBuilder()
    book.styles_xml = styles_xml(default)
    stock = book.sheet("Stock")
    stock.protected = protected
    stock.rows("A1", [["SKU", "Item", "Quantity", "Price"],
                      ["P-1", "Desk lamp", 3, 12.5],
                      ["P-2", "Desk fan with a long name", 12, 1234.5]])
    stock.cell("B5", 12.5, style=PRE)
    book.sheet("Orders").rows("A1", [["Order", "Total"], ["O-1", 40]])
    return book


def _open(default=("Calibri", 11), **options) -> tuple[bytes, Workbook]:
    data = stock_book(default, **options).to_bytes()
    return data, Workbook.open(data)


def _part(wb: Workbook, name: str) -> etree._Element:
    with zipfile.ZipFile(io.BytesIO(wb.to_bytes())) as archive:
        return etree.fromstring(archive.read(name))


def _styles(wb):
    return _part(wb, "xl/styles.xml")


def _sheet(wb, number=1):
    return _part(wb, f"xl/worksheets/sheet{number}.xml")


def _xf(wb, ref, number=1):
    for node in _sheet(wb, number).iter(X + "c"):
        if node.get("r") == ref:
            index = int(node.get("s", "0"))
            return _styles(wb).find(X + "cellXfs")[index]
    raise AssertionError(ref)


def _child(styles, container, index):
    return styles.find(X + container)[int(index)]


def _counts(wb):
    root = _styles(wb)
    return {name: len(root.find(X + name)) for name in ("fonts", "fills", "borders", "cellXfs")}


# -- cell formats -------------------------------------------------------------------------------


def test_only_what_is_asked_changes_and_equal_formats_are_reused():
    """Rule 3 (a no-op is no error) and 6 (what Excel writes)."""
    data, wb = _open()
    done = wb.format("Stock!A1:D1", font={"bold": True, "color": "#1f3864"}, fill="#DDEBF7",
                     border={"bottom": "thin"})
    assert done.changed == 4 and done.cells == {"Stock": ["A1:D1"]}
    before = _counts(wb)
    again = wb.format("Stock!A1:D1", font={"bold": True, "color": "#1F3864"}, fill="#DDEBF7",
                      border={"bottom": "thin"})
    assert again.changed == 0 and again.formats_added == 0 and _counts(wb) == before
    styles = _styles(wb)
    xf = _xf(wb, "A1")
    assert {_xf(wb, ref).get("fontId") for ref in ("B1", "C1", "D1")} == {xf.get("fontId")}
    font = _child(styles, "fonts", xf.get("fontId"))
    assert [etree.QName(c).localname for c in font] == ["b", "sz", "color", "name"]
    assert font.find(X + "color").get("rgb") == "FF1F3864"
    fill = _child(styles, "fills", xf.get("fillId"))
    assert etree.tostring(fill[0], method="c14n").decode() == (
        '<patternFill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'patternType="solid"><fgColor rgb="FFDDEBF7"></fgColor><bgColor indexed="64">'
        '</bgColor></patternFill>')
    bottom = _child(styles, "borders", xf.get("borderId")).find(X + "bottom")
    assert bottom.get("style") == "thin" and bottom[0].get("indexed") == "64"
    assert (xf.get("applyFont"), xf.get("applyFill"), xf.get("applyBorder")) == ("1", "1", "1")
    # bold on a formatted cell keeps its number format, borders, alignment and protection
    wb.format("Stock!B5", font={"bold": True})
    kept = _xf(wb, "B5")
    assert kept.get("numFmtId") == "164" and kept.get("borderId") == "2"
    assert kept.find(X + "alignment").get("horizontal") == "center"
    assert kept.find(X + "protection").get("locked") == "0"
    italic_bold = _child(_styles(wb), "fonts", kept.get("fontId"))
    assert [etree.QName(c).localname for c in italic_bold] == ["b", "i", "sz", "name"]
    assert wb.calculation_status == "cached_unverified"     # formatting changes no value
    with zipfile.ZipFile(io.BytesIO(data)) as a, zipfile.ZipFile(io.BytesIO(wb.to_bytes())) as b:
        changed = {n for n in a.namelist() if a.read(n) != b.read(n)}
    assert changed == {"xl/styles.xml", "xl/worksheets/sheet1.xml"}


def test_a_number_format_code_is_a_built_in_id_or_added_once():
    _, wb = _open()
    wb.format("Stock!D2", number_format="0.00")
    wb.format("Stock!D3", number_format="#,##0.00")
    wb.format("Stock!C2:C3", number_format="0.0")
    wb.format("Orders!B2", number_format="0.0")
    assert _xf(wb, "D2").get("numFmtId") == "2" and _xf(wb, "D3").get("numFmtId") == "4"
    custom = _xf(wb, "C2").get("numFmtId")
    assert custom == "165" and _xf(wb, "B2", 2).get("numFmtId") == custom
    codes = {n.get("formatCode") for n in _styles(wb).find(X + "numFmts")}
    assert codes == {"0.000", "0.0"}
    wb.format("Stock!C2", number_format=10)
    assert _xf(wb, "C2").get("numFmtId") == "10"


def test_alignment_and_protection_as_excel_writes_them():
    _, wb = _open()
    wb.format("Stock!A2", alignment={"horizontal": "centre", "vertical": "middle"})
    wb.format("Stock!B2", alignment={"indent": 2})
    wb.format("Stock!C2", alignment={"wrap": True, "rotation": 45})
    wb.format("Stock!D2", protection={"locked": False})
    wb.format("Stock!A3", protection={"hidden": True})
    assert dict(_xf(wb, "A2").find(X + "alignment").attrib) == {"horizontal": "center",
                                                                "vertical": "center"}
    assert dict(_xf(wb, "B2").find(X + "alignment").attrib) == {"horizontal": "left",
                                                                "indent": "2"}
    assert dict(_xf(wb, "C2").find(X + "alignment").attrib) == {"textRotation": "45",
                                                                "wrapText": "1"}
    assert dict(_xf(wb, "D2").find(X + "protection").attrib) == {"locked": "0"}
    assert dict(_xf(wb, "A3").find(X + "protection").attrib) == {"hidden": "1"}
    assert _xf(wb, "D2").get("applyProtection") == "1"
    # taken back: the format is the plain one again
    wb.format("Stock!D2", protection={"locked": True})
    assert _xf(wb, "D2").get("numFmtId") == "0" and _xf(wb, "D2").find(X + "protection") is None


# Excel applied each style by name to a cell of format PRE (measured): the parts the style
# includes are the style's, the others the cell's; apply* is 1 where a part differs from
# the style's.
@pytest.mark.parametrize("name, own, flags, alignment, protection", [
    ("Heading 1", {"numFmtId": "164", "fillId": "0"}, {"NumberFormat", "Alignment",
                                                        "Protection"},
     {"horizontal": "center"}, {"locked": "0"}),
    ("kop 1", {"numFmtId": "164", "fillId": "0"}, {"NumberFormat", "Alignment",
                                                    "Protection"},
     {"horizontal": "center"}, {"locked": "0"}),
    ("Stock heading", {"numFmtId": "0", "borderId": "0"}, set(), None, None),
    ("Stock note", {"numFmtId": "164", "borderId": "2"}, {"NumberFormat", "Border",
                                                          "Alignment", "Protection"},
     {"horizontal": "center"}, {"locked": "0"}),
    ("Stock locked", {"numFmtId": "164", "borderId": "2", "fillId": "0"},
     {"NumberFormat", "Font", "Border"}, {"horizontal": "right"}, {"hidden": "1"}),
])
def test_a_named_cell_style_is_applied_as_excel_applies_it(name, own, flags, alignment,
                                                            protection):
    _, wb = _open()
    italic = _xf(wb, "B5").get("fontId")
    wb.format("Stock!B5", style=name)
    xf = _xf(wb, "B5")
    for key, value in own.items():
        assert xf.get(key) == value, key
    if name == "Stock locked":
        assert xf.get("fontId") == italic
    assert {a[5:] for a in xf.attrib if a.startswith("apply")} == flags
    assert {a: xf.get(a) for a in xf.attrib if a.startswith("apply")} == {
        "apply" + f: "1" for f in flags}
    found = xf.find(X + "alignment")
    assert (None if found is None else dict(found.attrib)) == alignment
    found = xf.find(X + "protection")
    assert (None if found is None else dict(found.attrib)) == protection


def test_normal_gives_the_plain_format_back():
    _, wb = _open()
    wb.format("Stock!B5", style="Normal")
    assert _xf(wb, "B5") is not None
    assert all(node.get("s") is None for node in _sheet(wb).iter(X + "c")
               if node.get("r") == "B5")


def test_like_gives_another_sheet_s_cell_format_and_options_apply_on_top():
    _, wb = _open()
    wb.format("Stock!A1", font={"bold": True}, fill="#FFFF00")
    wb.format("Orders!A1:B1", like="Stock!A1")
    assert _xf(wb, "A1", 2).get("fontId") == _xf(wb, "A1").get("fontId")
    wb.format("Orders!B2", like="Stock!A1", number_format="0.0")
    assert _xf(wb, "B2", 2).get("fillId") == _xf(wb, "A1").get("fillId")
    assert _xf(wb, "B2", 2).get("numFmtId") == "165"


def test_an_outline_border_goes_round_the_range():
    _, wb = _open()
    wb.format("Stock!A1:C3", border={"outline": "medium"})
    styles = _styles(wb)

    def sides(ref):
        border = _child(styles, "borders", _xf(wb, ref).get("borderId"))
        return {etree.QName(s).localname for s in border if s.get("style")}

    assert sides("A1") == {"left", "top"} and sides("B2") == set()
    assert sides("C3") == {"right", "bottom"} and sides("B1") == {"top"}


def test_whole_columns_and_rows_are_formatted_as_excel_formats_them():
    """Measured: a column's style with the default width, a row's s with customFormat,
    and each cell held; no cell made where there is none."""
    _, wb = _open()
    done = wb.format(["Stock!D:D", "Stock!2:2"], font={"bold": True})
    sheet = _sheet(wb)
    col = sheet.find(X + "cols")[0]
    assert (col.get("min"), col.get("max")) == ("4", "4") and col.get("style")
    # no sheetFormatPr: Excel for Mac's base 10 (measured), 10.83203125 with Calibri 11
    assert col.get("width") == str(stored_width(10, 6)) == "10.83203125"
    row = [r for r in sheet.iter(X + "row") if r.get("r") == "2"][0]
    assert row.get("customFormat") == "1" and row.get("s")
    held = {c.get("r") for c in sheet.iter(X + "c") if c.get("s")}
    assert {"D1", "D2", "D3", "A2", "B2", "C2"} <= held and "D5" not in held
    assert done.columns == ("Stock!D:D",) and done.rows == ("Stock!2:2",)


def test_refusals_list_every_problem_and_change_nothing():
    """Rule 2: each refusal names the field and what would work."""
    data, wb = _open()
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1", font={"bold": "yes", "colour": "red", "weight": 7},
                  border={"bottom": "wavy"}, alignment={"horizontal": "middle"})
    found = dict(error.value.problems)
    assert set(error.value.problems) >= {("bad_format", "font.bold"),
                                         ("bad_format", "font.colour"),
                                         ("bad_format", "font.weight"),
                                         ("bad_format", "border.bottom"),
                                         ("bad_format", "alignment.horizontal")}
    assert "#RRGGBB" in str(error.value) and found
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1", style="Nope")
    assert error.value.code == "no_cell_style" and "Stock heading" in str(error.value)
    with pytest.raises(EditError) as error:
        wb.format("Nowhere!A1", fill="#FFFFFF")
    assert error.value.code == "bad_target" and "Stock" in str(error.value)
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1")
    assert error.value.code == "bad_format"
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1", like="Stock!B5", style="Normal")
    assert error.value.code == "bad_format"
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1:Z50000", fill="#FFFFFF")
    assert error.value.code == "too_many_cells" and "whole" in str(error.value)
    assert wb.to_bytes() == data and wb.version == 0


def test_a_protected_sheet_and_the_write_policy_refuse_formatting():
    _, wb = _open(protected=True)
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1", font={"bold": True})
    assert error.value.code == "protected"
    data = stock_book().to_bytes()
    wb = Workbook.open(data, policy={"cells": ["Stock!A1:D3"]})
    wb.format("Stock!A1:D1", font={"bold": True})
    with pytest.raises(EditError) as error:
        wb.format("Stock!A1:E1", font={"bold": True})
    assert error.value.code == "not_writable"
    with pytest.raises(EditError) as error:
        wb.freeze_panes("Stock", rows=1)
    assert error.value.code == "not_writable"


def test_a_protected_sheet_allows_what_excel_allows():
    """Measured: on a sheet protected with Excel's defaults, freezing panes and hiding
    gridlines work and formatting, widths and heights do not; with formatCells,
    formatColumns and formatRows 0 they all do."""
    _, wb = _open(protected=True)
    assert wb.freeze_panes("Stock", at="B2").changed == 1
    assert wb.set_gridlines("Stock", False).changed == 1
    for call in (lambda: wb.set_column_widths("Stock!B:B", 20),
                 lambda: wb.set_row_heights("Stock!2:2", 30),
                 lambda: wb.format("Stock!A1", font={"bold": True})):
        with pytest.raises(EditError) as error:
            call()
        assert error.value.code == "protected"
    book = stock_book()
    data = book.to_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(
        b"</sheetData>", b'</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" '
        b'formatCells="0" formatColumns="0" formatRows="0"/>')
    from build import zip_parts

    wb = Workbook.open(zip_parts(parts))
    assert wb.sheet("Stock").protected
    assert wb.set_column_widths("Stock!B:B", 20).changed == 1
    assert wb.set_row_heights("Stock!2:2", 30).changed == 1
    assert wb.format("Stock!A1", font={"bold": True}).changed == 1
    assert wb.format("Stock!D:D", font={"bold": True}).columns == ("Stock!D:D",)


def test_formatting_is_one_step_undone_to_the_bytes_opened_and_atomic_in_a_batch():
    data, wb = _open()
    wb.format("Stock!A1:D1", font={"bold": True})
    wb.set_column_widths("Stock!A:D", "fit")
    wb.freeze_panes("Stock", rows=1)
    assert wb.version == 3
    while wb.undo():
        pass
    assert wb.to_bytes() == data
    with pytest.raises(EditError):
        with wb.batch():
            wb.format("Stock!A1", fill="#FFFF00")
            wb.format("Stock!A1", fill="yellow")
    assert wb.to_bytes() == data


# -- widths and heights -------------------------------------------------------------------------


def test_a_width_is_stored_as_excel_stores_one_typed_in():
    """Measured: 12, 20, 8.43 and 1 characters in a Calibri 11 workbook."""
    assert digit_width(Face("Calibri", 11)) == 6 and digit_width(Face("Calibri", 16)) == 8
    for given, stored in ((12, 12.83203125), (20, 20.83203125), (8.43, 9.33203125),
                          (1, 1.83203125)):
        assert stored_width(given, 6) == stored
    _, wb = _open()
    done = wb.set_column_widths("Stock!B:C", 12)
    cols = _sheet(wb).find(X + "cols")
    assert [(c.get("min"), c.get("max"), c.get("width"), c.get("customWidth")) for c in cols] \
        == [("2", "3", "12.83203125", "1")]
    assert done.widths == {"Stock!B": 12.83203125, "Stock!C": 12.83203125}
    assert wb.set_column_widths("Stock!B:C", 12).changed == 0          # rule 3
    wb.set_column_widths("Stock!C:C", None)
    cols = _sheet(wb).find(X + "cols")
    assert [(c.get("min"), c.get("max")) for c in cols] == [("2", "2")]
    with pytest.raises(EditError):
        wb.set_column_widths("Stock!B:B", 300)


def test_row_heights_are_points_with_custom_height():
    _, wb = _open()
    wb.set_row_heights("Stock!1:2", 30)
    rows = {r.get("r"): r for r in _sheet(wb).iter(X + "row")}
    assert rows["1"].get("ht") == "30" and rows["2"].get("customHeight") == "1"
    wb.set_row_heights("Stock!2:2", None)
    rows = {r.get("r"): r for r in _sheet(wb).iter(X + "row")}
    assert rows["2"].get("ht") is None and rows["1"].get("ht") == "30"


#: Excel's AutoFit of one text per column (measured; Excel for Mac 16.106, 2026-11-08).
FIT_TEXTS = ["0", "00", "0000000000", "a", "Stock", "Stock level by warehouse",
             "WWWWWWWWWW", "iiiiiiiiii", "Total, net of returns (EUR)", "x" * 40, 12345,
             3.14159, "m" * 20, "1111111111", "Aisle 7 / shelf B"]
FIT_MEASURED = {
    ("Calibri", 11): [2.1640625, 3.1640625, 11.1640625, 2, 5.33203125, 20, 18, 6.1640625,
                      20.5, 35.1640625, 6.1640625, 8.1640625, 31.83203125, 11.1640625,
                      12.1640625],
    ("Arial", 10): [2.1640625, 3.1640625, 11.1640625, 2.1640625, 5.5, 21, 16.33203125, 4.5,
                    21.6640625, 35.1640625, 6.1640625, 8.1640625, 28.33203125, 11.1640625,
                    13.1640625],
    ("Times New Roman", 12): [2.1640625, 3.1640625, 11.1640625, 2, 5.6640625, 21.6640625,
                              19.6640625, 6.1640625, 22.83203125, 42.1640625, 6.1640625,
                              8.1640625, 31.83203125, 11.1640625, 14],
}


@pytest.mark.parametrize("default", sorted(FIT_MEASURED))
def test_fitted_widths_are_excel_s_autofit_on_the_probe_s_columns(default):
    """E7's done-when: every probe column of these fonts fits to Excel's own width."""
    book = WorkbookBuilder()
    book.styles_xml = styles_xml(default)
    fit = book.sheet("Fit")
    for n, text in enumerate(FIT_TEXTS):
        fit.cell(f"{chr(65 + n)}1", text)
    wb = Workbook.open(book.to_bytes())
    # the probe's Excel (region Dutch, Belgium) showed 3.14159 as 3,14159
    wb.locale = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                       language="nl-BE", list_separator=";")
    done = wb.set_column_widths(f"Fit!A:{chr(64 + len(FIT_TEXTS))}", "fit")
    got = [done.widths[f"Fit!{chr(65 + n)}"] for n in range(len(FIT_TEXTS))]
    assert got == FIT_MEASURED[default]
    col = _sheet(wb).find(X + "cols")[0]
    assert col.get("bestFit") == "1" and col.get("customWidth") == "1"


#: Excel's AutoFit of four texts in other fonts in a Calibri 11 workbook (measured), for
#: the cell formats 1-9 of styles_xml.
FONT_TEXTS = ["Stock level by warehouse", "0000000000", "WWWWWWWWWW", "iiiiiiiiii"]
FONT_MEASURED = {1: [20.6640625, 11.1640625, 18, 6.1640625],
                 2: [24.83203125, 12.83203125, 21.5, 6.1640625],
                 3: [21, 11.1640625, 16.33203125, 4.5],
                 4: [21.6640625, 11.1640625, 19.6640625, 6.1640625],
                 5: [25.5, 11.1640625, 11.1640625, 11.1640625],
                 6: [19.83203125, 11.1640625, 16.33203125, 4.5],
                 7: [21.5, 11.1640625, 18, 6.1640625],
                 8: [21, 11.1640625, 18, 6.1640625],
                 9: [30.83203125, 14.5, 23.1640625, 7.83203125]}


def test_fitted_widths_in_other_fonts_and_sizes():
    for style, measured in FONT_MEASURED.items():
        got = []
        for text in FONT_TEXTS:
            name, size, *bold = FONTS[style - 1]
            face = Face(name, size, bool(bold and bold[0]))
            got.append(fitted_width(text_pixels(face, text, 6), 6, fit_padding(face)))
        assert got == measured, FONTS[style - 1]
    # 8 and 20 point Calibri: the padding is the cell font's, 5 and 9 px (measured)
    for size, measured in ((8, [14.83203125, 7.5, 12.5, 4.1640625]),
                           (20, [37.33203125, 18.33203125, 32.1640625, 9.83203125])):
        face = Face("Calibri", size)
        assert [fitted_width(text_pixels(face, text, 6), 6, fit_padding(face))
                for text in FONT_TEXTS] == measured


#: Excel's AutoFit padding by the cell font's size (measured 2026-11-09 in Calibri 11 and
#: Aptos 11 workbooks, eight texts a size, every one exact with it).
PADDING_MEASURED = {"Calibri": {6: 5, 7: 5, 8: 5, 9: 7, 10: 7, 11: 7, 12: 7, 14: 7, 16: 7,
                                18: 9, 20: 9, 24: 9, 28: 11, 36: 13},
                    "Aptos": {6: 5, 7: 5, 8: 5, 9: 7, 10: 7, 11: 7, 12: 7, 14: 7, 16: 9,
                              18: 9, 20: 9, 24: 11, 28: 11, 36: 13},
                    "Aptos Narrow": {8: 5, 9: 7, 10: 7, 11: 7, 12: 7, 14: 7, 16: 7, 18: 9,
                                     20: 9, 24: 9}}


def test_the_fit_padding_grows_with_the_cell_font_as_excel_s():
    for name, sizes in PADDING_MEASURED.items():
        for size, padding in sizes.items():
            assert fit_padding(Face(name, size)) == padding, (name, size)
            assert fit_padding(Face(name, size, True)) == padding, (name, size)


#: Excel's AutoFit in Aptos Narrow, Excel's default body face, in an Aptos Narrow 11
#: workbook (measured 2026-11-09), regular and bold.
NARROW_TEXTS = ["0", "a", "Stock", "WWWWW", "iiiiiiiiii", "0000000000",
                "Stock level by warehouse", "x" * 40]
NARROW_MEASURED = {
    (8, False): [1.5, 1.5, 3.83203125, 6.6640625, 4.1640625, 7.5, 13.83203125, 21.1640625],
    (11, False): [2.1640625, 2, 5.33203125, 8.6640625, 4.5, 11.1640625, 19.83203125,
                  28.33203125],
    (24, False): [3.5, 3.5, 10.5, 18.33203125, 9.83203125, 21.83203125, 42.5, 70],
    (11, True): [2.1640625, 2, 5.5, 8.6640625, 6.1640625, 11.1640625, 20.83203125,
                 35.1640625],
}


def test_fitted_widths_in_aptos_narrow():
    for (size, bold), measured in NARROW_MEASURED.items():
        face = Face("Aptos Narrow", size, bold)
        got = [fitted_width(text_pixels(face, text, 6), 6, fit_padding(face))
               for text in NARROW_TEXTS]
        assert got == measured, (size, bold)


def test_a_default_font_without_measured_widths_refuses_a_width_rather_than_fail():
    """The column width unit is the default font's widest digit: unknown for a face with
    no widths at all (refused, naming the fonts that have them); a face measured with a
    clone's (Calibri Light) gives its digit, while its text is not fitted."""
    wb = Workbook.open(stock_book(("Rockwell", 11)).to_bytes())
    for width in (12, "fit"):
        with pytest.raises(EditError) as error:
            wb.set_column_widths("Stock!A:A", width)
        assert error.value.code == "unmeasured_font" and "Rockwell" in str(error.value)
    wb = Workbook.open(stock_book(("Calibri Light", 11)).to_bytes())
    assert wb.set_column_widths("Stock!A:A", 12).widths == {"Stock!A": 12.83203125}


def test_a_fit_refuses_a_font_without_measured_widths_and_lists_what_it_skips():
    """Rule 2: the refusal names the font, the fonts measured and the way out."""
    book = stock_book()
    sheet = book.sheets[0]
    sheet.cell("F1", "Light text")
    sheet.cell("G1", "Wrapped", style=PRE)
    wb = Workbook.open(book.to_bytes())
    wb.format("Stock!F1", font={"name": "Calibri Light"})    # measured as Carlito: 1.3% off
    with pytest.raises(EditError) as error:
        wb.set_column_widths("Stock!F:F", "fit")
    assert error.value.code == "unmeasured_font" and "Calibri Light" in str(error.value)
    assert "Aptos Narrow" in str(error.value)          # among the fonts measured
    assert "Calibri" in str(error.value) and "width instead" in str(error.value)
    wb.format("Stock!G1", alignment={"wrap": True})
    done = wb.set_column_widths("Stock!G:G", "fit")
    assert done.skipped == {"wrapped": ["Stock!G1"]} and done.widths == {}
    assert done.notes


# -- the sheet ---------------------------------------------------------------------------------------


def test_freeze_panes_as_excel_writes_them():
    """Measured: frozen at C4."""
    _, wb = _open()
    wb.freeze_panes("Stock", at="C4")
    view = _sheet(wb).find(X + "sheetViews")[0]
    pane = view.find(X + "pane")
    assert dict(pane.attrib) == {"xSplit": "2", "ySplit": "3", "topLeftCell": "C4",
                                 "activePane": "bottomRight", "state": "frozen"}
    assert [(s.get("pane"), s.get("activeCell"), s.get("sqref"))
            for s in view.findall(X + "selection")] == [
        ("topRight", "C1", "C1"), ("bottomLeft", "A4", "A4"), ("bottomRight", "C4", "C4")]
    wb.freeze_panes("Stock", rows=1)
    view = _sheet(wb).find(X + "sheetViews")[0]
    assert view.find(X + "pane").get("activePane") == "bottomLeft"
    assert wb.freeze_panes("Stock", rows=1).changed == 0
    wb.freeze_panes("Stock", rows=0, columns=0)
    view = _sheet(wb).find(X + "sheetViews")[0]
    assert view.find(X + "pane") is None and not view.findall(X + "selection")
    with pytest.raises(EditError):
        wb.freeze_panes("Stock", at="C4", rows=1)


def test_gridlines_are_hidden_and_shown():
    _, wb = _open()
    wb.set_gridlines("Stock", False)
    assert _sheet(wb).find(X + "sheetViews")[0].get("showGridLines") == "0"
    assert wb.set_gridlines("Stock", False).changed == 0
    wb.set_gridlines("Stock", True)
    assert _sheet(wb).find(X + "sheetViews")[0].get("showGridLines") is None


def test_the_password_hash_is_excel_s():
    """Measured: the hash Excel's Protect Sheet wrote for "secret" with its salt."""
    salt = base64.b64decode("y7yrIbVVO2KmBKSjjnrk5Q==")
    assert password_hash("secret", salt) == (
        "8EGmaL+NbIbUawUy/wCze2HDdrnMMNkAVw3c7H6nulTLZ3HwfE2qHF2/zHCnuxLOnBHQlG8zqOwNHuzUWVLeSw==")


def test_a_sheet_is_protected_and_unprotected():
    _, wb = _open()
    wb.protect_sheet("Stock")
    node = _sheet(wb).find(X + "sheetProtection")
    assert dict(node.attrib) == {"sheet": "1", "objects": "1", "scenarios": "1"}
    assert wb.sheet("Stock").protected
    assert wb.protect_sheet("Stock").changed == 0
    with pytest.raises(EditError) as error:
        wb.set_cells({"Stock!A2": "P-9"})
    assert error.value.code == "protected"
    wb.unprotect_sheet("Stock")
    wb.protect_sheet("Stock", password="s3cret")
    node = _sheet(wb).find(X + "sheetProtection")
    assert node.get("algorithmName") == "SHA-512" and node.get("spinCount") == "100000"
    assert len(base64.b64decode(node.get("saltValue"))) == 16
    assert list(node.attrib)[-3:] == ["sheet", "objects", "scenarios"]
    for wrong in (None, "nope"):
        with pytest.raises(EditError) as error:
            wb.unprotect_sheet("Stock", password=wrong)
        assert error.value.code == "wrong_password"
    wb.unprotect_sheet("Stock", password="s3cret")
    assert _sheet(wb).find(X + "sheetProtection") is None
    assert wb.unprotect_sheet("Stock").changed == 0


# -- the tool -----------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    from ooxml_edit.tools import Toolbox
    from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box):
    session = box.session()
    blob = session.add_blob(stock_book().to_bytes(), name="stock.xlsx")
    assert box.dispatch(session, "open_document", {"blob": blob}).ok
    return session


def test_xlsx_format_formats_cells_columns_and_the_sheet_in_one_step(box):
    """Rule 5: one call for a header row, its columns and the sheet; one undo step."""
    session = _session(box)
    result = box.dispatch(session, "xlsx_format", {
        "doc": "d1", "target": ["Stock!A1:D1"], "bold": True, "color": "#1F3864",
        "fill": "#DDEBF7", "border": "bottom", "horizontal": "center", "fit": True,
        "freeze": "A2", "gridlines": False})
    assert result.ok, result.to_json()
    assert result.data["changed"] >= 6 and set(result.data["widths"]) == {
        "Stock!A", "Stock!B", "Stock!C", "Stock!D"}
    assert session.entry("d1").version == 1
    again = box.dispatch(session, "xlsx_format", {"doc": "d1", "target": ["Stock!A1:D1"],
                                                  "bold": True, "freeze": "A2"})
    assert again.ok and again.data["changed"] == 0                     # rule 3
    assert box.dispatch(session, "undo", {"doc": "d1"}).ok


def test_xlsx_format_refuses_with_every_problem_and_the_way_out(box):
    """Rule 2: each problem listed, with what is valid; nothing changed."""
    session = _session(box)
    result = box.dispatch(session, "xlsx_format", {
        "doc": "d1", "target": ["Stock!A1"], "color": "navy", "style": "Nope"})
    assert result.error.code == "refused"
    codes = {p["code"] for p in result.error.details["problems"]}
    assert codes == {"bad_format", "no_cell_style"}
    assert "#RRGGBB" in result.error.message and "Stock heading" in result.error.message
    assert session.entry("d1").version == 0
    empty = box.dispatch(session, "xlsx_format", {"doc": "d1", "target": ["Stock!A1"]})
    assert empty.error.code == "invalid_arguments"
