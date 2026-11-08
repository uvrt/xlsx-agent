"""The Excel oracle on round 44's lane T: new workbooks, tables written, rows appended.

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven; without Excel, or
with Excel in use or its lock held, every test here skips).  Made-up workbooks only.

* ``Workbook.new()`` and ``Workbook.from_tables()`` open without a repair, and Excel's
  re-save keeps every part we wrote: what it changes is its own -- the application's name,
  its user as the last modifier, the localised names (``Werkbladen``, ``Standaard``,
  ``Kantoorthema`` in a Dutch Excel), the folder saved in, a revision id, the window's size,
  ``xr:uid`` attributes, ``spans`` and ``x14ac:dyDescent`` on the rows, and ``fullCalcOnLoad``
  taken off once it calculated (round 44 lane T's hand-back lists the diff).
* A table written by ``write_table(as_table=...)`` into a new workbook and into the corner
  shop's stock list (plain, ``r``-less, inline strings, with a table) opens as a table:
  Excel keeps its part -- ``ref``, the AutoFilter, the column names, the style banded by rows.
* Rows appended to a table with a totals row and a calculated column (``Orders``) keep the
  table's formulas and totals: Excel's re-save has the new ``ref`` and AutoFilter, computes
  the calculated column in the new rows and the totals over every row.
"""

from __future__ import annotations

import datetime as dt
import io
import sys
import zipfile

import pytest
from lxml import etree

from xlsx_agent import Workbook

import oracle
from shop import Shop
from test_write_table import _orders

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

pytestmark = pytest.mark.oracle


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


def _part(data: bytes, name: str):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return etree.fromstring(archive.read(name))


def _saved(excel, data: bytes, name: str) -> bytes:
    outcome = excel.resave(data, name=name)
    assert outcome.ok, outcome
    return outcome.saved.read_bytes()


def test_a_new_workbook_opens_and_excel_keeps_its_parts(excel):
    data = Workbook.new(["Stock", "Notes"], author="Stock clerk").to_bytes()
    saved = _saved(excel, data, "r44-new")
    with zipfile.ZipFile(io.BytesIO(data)) as ours, zipfile.ZipFile(io.BytesIO(saved)) as theirs:
        assert sorted(ours.namelist()) == sorted(theirs.namelist())
    for part in ("xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/workbook.xml"):
        mine, excels = _part(data, part), _part(saved, part)
        # Excel's own: the folder saved in (mc:AlternateContent's absPath), a revision id
        assert [etree.QName(c).localname for c in mine] == \
            [etree.QName(c).localname for c in excels
             if etree.QName(c).localname not in ("AlternateContent", "revisionPtr")], part
    font = _part(saved, "xl/styles.xml").find(f"{X}fonts/{X}font")
    assert font.find(f"{X}name").get("val") == "Aptos Narrow"
    assert font.find(f"{X}sz").get("val") == "12"
    assert [s.name for s in Workbook.open(saved).sheets] == ["Stock", "Notes"]


def test_from_tables_opens_and_keeps_its_table_and_values(excel):
    data = Workbook.from_tables({
        "Stock": {"columns": ["Item", "Qty", "Added"], "table": "Stock",
                  "rows": [["Lamp", 3, dt.date(2026, 1, 5)], ["Fan", 4, None]]},
        "Notes": {"columns": ["Note"], "rows": [[{"markdown": "**first**"}]]}}).to_bytes()
    saved = Workbook.open(_saved(excel, data, "r44-from-tables"))
    assert [t.name for t in saved.tables] == ["Stock"]
    rows = [[v.v for v in row] for row in saved.read("Stock!A2:C3")]
    assert rows == [["Lamp", 3, "2026-01-05"], ["Fan", 4, None]]


@pytest.mark.parametrize("options", [{}, {"rless": True}, {"shared": False}, {"table": True}],
                         ids=["plain", "rless", "inline", "table"])
def test_a_written_table_opens_as_a_table_with_filters_and_banding(excel, options):
    wb = Workbook.open(Shop(**options).to_bytes())
    wb.write_table("Stock", "N2", ["Shelf", "Count", "Checked"],
                   [["A", 3, dt.date(2026, 11, 1)], ["B", 4, None]], as_table="Shelves")
    data = wb.to_bytes()
    tag = "-".join(f"{k}" for k in options) or "plain"
    saved = _saved(excel, data, f"r44-table-{tag}")
    reopened = Workbook.open(saved)
    table = next(t for t in reopened._table_parts() if t.name == "Shelves")
    root = _part(saved, table.part)
    assert root.get("ref") == "N2:P4"
    assert root.find(f"{X}autoFilter").get("ref") == "N2:P4"
    assert [c.get("name") for c in root.iter(f"{X}tableColumn")] == ["Shelf", "Count",
                                                                      "Checked"]
    info = root.find(f"{X}tableStyleInfo")
    assert info.get("name") == "TableStyleMedium2" and info.get("showRowStripes") == "1"
    assert reopened.cell("Stock!P3").v == "2026-11-01"


def test_an_appended_table_keeps_its_formulas_and_totals(excel):
    wb = Workbook.open(_orders())
    wb.append_rows("t:Orders", [["Chair", 2, 20], {"Item": "Shelf", "Qty": 1, "Price": 35.5}])
    saved = _saved(excel, wb.to_bytes(), "r44-orders-appended")
    root = _part(saved, "xl/tables/table1.xml")
    assert root.get("ref") == "A1:D7" and root.find(f"{X}autoFilter").get("ref") == "A1:D6"
    reopened = Workbook.open(saved)
    assert reopened.cell("Orders!D5").v == 40 and reopened.cell("Orders!D6").v == 35.5
    assert reopened.cell("Orders!D7").v == 222 and reopened.cell("Orders!A7").v == "Total"
    assert reopened.cell("Orders!F1").v == 222


def test_rows_appended_below_a_sheet_s_values_open(excel):
    wb = Workbook.new(["Log"])
    wb.write_table("Log", "B2", ["When", "What"], [[dt.date(2026, 11, 1), "start"]])
    wb.append_rows("Log!B:C", [[dt.date(2026, 11, 2), "next"]])
    reopened = Workbook.open(_saved(excel, wb.to_bytes(), "r44-log"))
    assert [[v.v for v in row] for row in reopened.read("Log!B3:C4")] == [
        ["2026-11-01", "start"], ["2026-11-02", "next"]]
