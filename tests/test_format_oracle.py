"""The Excel oracle for formatting and layout (E7, production round 44, R53).

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven; without Excel,
or with it in use, every test here skips).  The corner shop's stock list in each variant,
and the stock list of :mod:`test_format` with its named cell styles, formatted with every
option -- fonts, fills, borders, number formats, alignment, protection, a named style,
``like``, whole columns and rows, widths given and fitted, row heights, freeze panes,
gridlines off and a sheet protected with a password -- open in Excel without a repair, and
Excel's re-save keeps what was written.
"""

from __future__ import annotations

import io
import sys
import zipfile

import pytest
from lxml import etree

from ooxml_edit.sheets import Cell

from xlsx_agent import Workbook

import oracle
from shop import FIRST_INPUT, Shop
from test_format import PRE, stock_book

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
VARIANTS = {"plain": {}, "rless": {"rless": True}, "inline": {"shared": False},
            "table": {"table": True}}


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


def _formatted(wb: Workbook, sheet: str, last_row: int) -> None:
    head = f"{sheet}!A1:H1" if sheet != "Stock" or last_row > 9 else f"{sheet}!A1:D1"
    wb.format(head, font={"bold": True, "color": "#1F3864", "size": 12}, fill="#DDEBF7",
              border={"bottom": "medium"}, alignment={"horizontal": "center", "wrap": True})
    wb.format(f"{sheet}!B2:B{last_row}", font={"italic": True, "name": "Arial"},
              number_format="0.0", alignment={"indent": 1})
    wb.format(f"{sheet}!C2:D{last_row}", border={"outline": {"style": "thin",
                                                              "color": "#FF0000"}},
              protection={"locked": False})
    wb.format(f"{sheet}!A{last_row}", like=f"{sheet}!A1")
    wb.format(f"{sheet}!E:E", fill="#FFFF00")
    wb.format(f"{sheet}!{last_row}:{last_row}", font={"underline": True, "strike": True})
    wb.set_column_widths(f"{sheet}!A:A", "fit")
    wb.set_column_widths(f"{sheet}!F:G", 14)
    wb.set_row_heights(f"{sheet}!1:1", 30)
    wb.freeze_panes(sheet, at="B2")
    wb.set_gridlines(sheet, False)
    wb.protect_sheet(sheet, password="s3cret")


def _parts(data: bytes) -> dict[str, etree._Element]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: etree.fromstring(archive.read(name)) for name in archive.namelist()
                if name.endswith(".xml")}


def _cell_format(parts, data: bytes, ref: str) -> dict:
    """A cell's format resolved: number format code, font, fill, border, alignment and
    protection, as canonical text (ids differ between our file and Excel's)."""
    styles = parts["xl/styles.xml"]
    s = Workbook.open(data).editor.cells.sheet("Stock").style(Cell.parse(ref))
    xf = styles.find(X + "cellXfs")[s]
    codes = {n.get("numFmtId"): n.get("formatCode")
             for n in (styles.find(X + "numFmts") if styles.find(X + "numFmts") is not None
                       else [])}

    def item(container, attr):
        node = styles.find(X + container)[int(xf.get(attr, "0"))]
        return sorted((etree.QName(c).localname, sorted(c.attrib.items()),
                       sorted((etree.QName(g).localname, sorted(g.attrib.items())) for g in c))
                      for c in node if c.get("val") != "2" or etree.QName(c).localname != "family")

    def child(tag):
        node = xf.find(X + tag)
        return None if node is None else sorted(node.attrib.items())

    number = xf.get("numFmtId", "0")
    return {"number": codes.get(number, number), "font": item("fonts", "fontId"),
            "fill": item("fills", "fillId"), "border": item("borders", "borderId"),
            "alignment": child("alignment"), "protection": child("protection")}


def _kept(written: bytes, saved: bytes, sheet_part: str, refs) -> None:
    ours, theirs = _parts(written), _parts(saved)
    for ref in refs:
        assert _cell_format(theirs, saved, ref) == _cell_format(ours, written, ref), ref
    sheet = theirs[sheet_part]
    view = sheet.find(X + "sheetViews")[0]
    assert view.get("showGridLines") == "0"
    pane = view.find(X + "pane")
    assert (pane.get("topLeftCell"), pane.get("state")) == ("B2", "frozen")
    protection = sheet.find(X + "sheetProtection")
    mine = ours[sheet_part].find(X + "sheetProtection")
    assert protection.get("hashValue") == mine.get("hashValue")
    widths = {c.get("min"): c.get("width") for c in sheet.find(X + "cols")}
    for c in ours[sheet_part].find(X + "cols"):
        assert widths.get(c.get("min")) == c.get("width"), c.get("min")
    first = [r for r in sheet.iter(X + "row") if r.get("r") == "1"][0]
    assert first.get("ht") == "30" and first.get("customHeight") == "1"


@pytest.mark.oracle
@pytest.mark.parametrize("variant", VARIANTS)
def test_excel_opens_a_formatted_shop_and_keeps_the_formats(excel, variant):
    wb = Workbook.open(Shop(**VARIANTS[variant]).to_bytes())
    last = FIRST_INPUT + 4
    _formatted(wb, "Stock", last)
    written = wb.to_bytes()
    outcome = excel.resave(written, name=f"shop-e7-{variant}")
    assert outcome.ok, outcome
    _kept(written, outcome.saved.read_bytes(), "xl/worksheets/sheet1.xml",
          ["A1", "D1", f"B{FIRST_INPUT}", f"C{FIRST_INPUT + 1}", f"D{last}", f"A{last}",
           f"E{FIRST_INPUT}"])


@pytest.mark.oracle
def test_excel_opens_named_styles_applied_and_keeps_them(excel):
    book = stock_book()
    for row, name in enumerate(["Heading 1", "Stock heading", "Stock note", "Stock locked"], 6):
        book.sheets[0].cell(f"B{row}", 12.5, style=PRE)
    wb = Workbook.open(book.to_bytes())
    for row, name in enumerate(["Heading 1", "Stock heading", "Stock note", "Stock locked"], 6):
        wb.format(f"Stock!B{row}", style=name)
    _formatted(wb, "Stock", 3)
    written = wb.to_bytes()
    outcome = excel.resave(written, name="stock-e7-styles")
    assert outcome.ok, outcome
    _kept(written, outcome.saved.read_bytes(), "xl/worksheets/sheet1.xml",
          ["A1", "B2", "C3", "B6", "B7", "B8", "B9"])
    saved = _parts(outcome.saved.read_bytes())
    styles = saved["xl/styles.xml"]
    names = {n.get("xfId"): n.get("name") for n in styles.find(X + "cellStyles")}
    for ref, name in (("B7", "Stock heading"), ("B8", "Stock note"), ("B9", "Stock locked")):
        cell = [c for c in saved["xl/worksheets/sheet1.xml"].iter(X + "c") if c.get("r") == ref][0]
        xf = styles.find(X + "cellXfs")[int(cell.get("s"))]
        assert names[xf.get("xfId")] == name, ref
