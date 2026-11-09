"""The Excel oracle on round 46's writes (R71 protected sheets, R72 dates and times).

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven; a held lock or
an Excel in use skips).  Made-up workbooks only: the staff rota of ``test_protected`` and the
log sheet of ``test_date_times``.  Excel for Mac must open each without asking to repair
it, and its own re-save must keep what was written: the values and serials of dates, dates
and times and times, the formats given (Excel's own ids and a code given), in both date
systems; on a protected sheet, the unlocked cells written, the sheet still protected with
the same hash, salt, spin count and options.
"""

from __future__ import annotations

import datetime as dt
import io
import re
import sys
import zipfile

import pytest
from lxml import etree

from xlsx_agent import Workbook

import oracle
from test_date_times import _book as log_book
from test_protected import _book as rota_book

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


def _saved(excel, data: bytes, name: str) -> bytes:
    outcome = excel.resave(data, name=name)
    assert outcome.outcome == "ok", (outcome.outcome, outcome.detail)
    return outcome.saved.read_bytes()


def _formats(data: bytes, sheet: str = "xl/worksheets/sheet1.xml") -> dict[str, str]:
    """Each cell's number format, as the code (the file's) or the built-in id."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = etree.fromstring(archive.read(sheet))
        styles = etree.fromstring(archive.read("xl/styles.xml"))
    xfs = styles.find(X + "cellXfs")
    codes = {f.get("numFmtId"): f.get("formatCode") for f in styles.iter(X + "numFmt")}
    out = {}
    for node in root.iter(X + "c"):
        number = xfs[int(node.get("s") or 0)].get("numFmtId")
        # A code compared without its escapes: Excel escapes a code's literals (\-, \ ).
        out[node.get("r")] = codes.get(number, number).replace("\\", "")
    return out


DATES = {
    "Log!A2": dt.datetime(2026, 1, 5, 14, 30),
    "Log!A3": {"date": "2026-01-05 14:30:15"},
    "Log!A4": dt.time(14, 30),
    "Log!A5": dt.time(14, 30, 15, 123400),
    "Log!A6": dt.datetime(2026, 1, 5, 23, 59, 59, 999600),
    "Log!A7": {"date": "2026-01-05T08:15", "format": "yyyy-mm-dd hh:mm"},
    "Log!B1": dt.datetime(2026, 1, 5, 9, 0),        # the short date: kept
    "Log!C1": dt.datetime(2026, 1, 5, 9, 0),        # h:mm: Excel's date and time
    "Log!E1": dt.time(9, 0),                        # m/d/yyyy h:mm: Excel's h:mm
}


@pytest.mark.oracle
@pytest.mark.parametrize("date1904", [False, True], ids=["1900", "1904"])
def test_excel_keeps_dates_and_times_as_written(excel, date1904):
    wb = Workbook.open(log_book(date1904))
    wb.set_cells(DATES)
    wb.write_table("Log", "K1", ["Shift", "Start", "Clock"],
                   [["early", dt.datetime(2026, 1, 5, 6, 0), dt.time(6, 0)]], as_table="Shifts")
    wb.append_rows("t:Shifts", [["late", {"date": "2026-01-05 14:30"}, {"date": "14:30:15"}]])
    written = wb.to_bytes()
    saved = _saved(excel, written, f"r72-{'1904' if date1904 else '1900'}")
    again = Workbook.open(saved)
    for ref in [*DATES, "Log!L2", "Log!M2", "Log!L3", "Log!M3"]:
        before, after = wb.cell(ref), again.cell(ref)
        assert (after.t, after.v) == (before.t, before.v), ref
    assert again.date1904 == date1904
    ours, excels = _formats(written), _formats(saved)
    for ref in ("A2", "A3", "A4", "A5", "A6", "A7", "B1", "C1", "E1", "L2", "M2", "L3", "M3"):
        assert excels[ref] == ours[ref], ref
    assert (ours["A2"], ours["A4"], ours["A5"], ours["B1"], ours["C1"], ours["E1"]) == \
        ("22", "20", "21", "14", "22", "20")
    assert ours["A7"] == "yyyy-mm-dd hh:mm"


def _guard(data: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        text = archive.read("xl/worksheets/sheet1.xml").decode()
    node = re.search(r"<sheetProtection[^>]*/>", text).group(0)
    return dict(re.findall(r'(\w+)="([^"]*)"', node))


@pytest.mark.oracle
def test_excel_opens_a_protected_sheet_written_in_its_unlocked_cells(excel):
    data = rota_book()
    wb = Workbook.open(data)
    wb.set_cells({"Rota!B2": 9, "Rota!C2": {"formula": "=B2-1"}, "Rota!E20": "extra",
                  "Rota!F6": "row six", "Rota!B3": dt.date(2026, 1, 5),
                  "Rota!C3": dt.datetime(2026, 1, 5, 9), "Rota!B4": dt.time(9, 30),
                  "Rota!C4": {"markdown": "early\nlate"}})
    wb.append_rows("Rota!E:E", [["one"], ["two"]])
    written = wb.to_bytes()
    saved = _saved(excel, written, "r71-rota")
    again = Workbook.open(saved)
    assert again.sheet("Rota").protected
    assert _guard(saved) == _guard(data)       # the same hash, salt, spin count, options
    for ref in ("Rota!B2", "Rota!E20", "Rota!F6", "Rota!B3", "Rota!C3", "Rota!B4",
                "Rota!C4", "Rota!E5", "Rota!E6"):
        assert (again.cell(ref).t, again.cell(ref).v) == (wb.cell(ref).t, wb.cell(ref).v), ref
    assert again.cell("Rota!C2").formula and again.cell("Rota!C2").v == 8
    assert again.cell("Rota!A2").v == "Ann"
