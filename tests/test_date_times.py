"""Dates and times written (round 46, R72): a date and time, a time of day alone, Excel's
formats for them, its precision, and refusals that say what would do.

Made-up workbooks only: a log sheet whose cells carry the formats a template has.  What
Excel stores and which format it gives were measured in Excel for Mac 16.106 (2026-11-09,
typed into cells of 33 formats; :mod:`xlsx_agent.edit.cells`); the serials here are the
ones Excel saved."""

from __future__ import annotations

import datetime as dt
import io
import zipfile

import pytest
from lxml import etree

from ooxml_common.locales import Locale

from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook
from xlsx_agent.edit import EditError
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from build import WorkbookBuilder

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NL_BE = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")

#: The cells of the log sheet's first row and the format each carries (a number in each).
FORMATS = {"B1": 14, "C1": 20, "D1": 2, "E1": 22, "F1": 49, "G1": 46}
CUSTOM = {"H1": "yyyy-mm-dd", "I1": "hh:mm:ss.000"}


def _book(date1904: bool = False) -> bytes:
    book = WorkbookBuilder(date1904=date1904)
    log = book.sheet("Log")
    log.cell("A1", "Shift")
    for ref, number_format in FORMATS.items():
        log.cell(ref, 7, style=book.style(num_fmt_id=number_format))
    for ref, code in CUSTOM.items():
        log.cell(ref, 7, style=book.style(num_fmt=code))
    return book.to_bytes()


def _open(date1904: bool = False) -> Workbook:
    return Workbook.open(_book(date1904))


def _node(wb: Workbook, ref: str):
    with zipfile.ZipFile(io.BytesIO(wb.to_bytes())) as archive:
        root = etree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        styles = etree.fromstring(archive.read("xl/styles.xml"))
    node = next(c for c in root.iter(X + "c") if c.get("r") == ref)
    xf = styles.find(X + "cellXfs")[int(node.get("s") or 0)]
    codes = {f.get("numFmtId"): f.get("formatCode") for f in styles.iter(X + "numFmt")}
    number_format = xf.get("numFmtId")
    return node.findtext(X + "v"), int(number_format), codes.get(number_format)


def _refused(wb: Workbook, cells, **options) -> EditError:
    before = wb.to_bytes()
    with pytest.raises(EditError) as raised:
        wb.set_cells(cells, **options)
    assert wb.to_bytes() == before
    return raised.value


# -- what is stored ---------------------------------------------------------------------------


@pytest.mark.parametrize("date1904, day", [(False, 46027), (True, 44565)], ids=["1900", "1904"])
def test_a_date_and_time_is_the_serial_excel_stores_in_either_date_system(date1904, day):
    wb = _open(date1904)
    wb.set_cells({"Log!A2": dt.datetime(2026, 1, 5, 14, 30),
                  "Log!A3": {"date": "2026-01-05 14:30:15"},
                  "Log!A4": {"date": "2026-01-05T14:30:15.123"},
                  "Log!A5": dt.time(14, 30), "Log!A6": {"date": "14:30:15"}})
    # Excel's own: 2026-01-05 14:30 typed was saved as 46027.604166666664 (44565.6... in
    # the 1904 system), 14:30 as 0.60416666666666663 -- the same doubles.
    assert float(_node(wb, "A2")[0]) == float(f"{day}.604166666664")
    assert float(_node(wb, "A3")[0]) == float(f"{day}.60434027778")
    assert float(_node(wb, "A4")[0]) == float(f"{day}.604341701386")
    assert float(_node(wb, "A5")[0]) == 0.60416666666666663
    assert float(_node(wb, "A6")[0]) == 0.60434027777777777
    assert (wb.cell("Log!A2").t, wb.cell("Log!A2").v) == ("datetime", "2026-01-05T14:30:00")
    assert (wb.cell("Log!A5").t, wb.cell("Log!A5").v) == ("time", "14:30:00")


def test_a_time_is_kept_to_the_millisecond_rounded_half_up_as_excel_keeps_a_typed_one():
    wb = _open()
    wb.set_cells({"Log!A2": dt.time(14, 30, 15, 123400),
                  "Log!A3": {"date": "14:30:15.12345"},
                  "Log!A4": dt.time(14, 30, 15, 999600),
                  "Log!A5": dt.time(0, 0, 0, 500),
                  "Log!A6": dt.time(0, 0, 0, 400),
                  "Log!A7": dt.datetime(2026, 1, 5, 23, 59, 59, 999600)})
    assert wb.cell("Log!A2").v == wb.cell("Log!A3").v == "14:30:15.123"
    assert wb.cell("Log!A4").v == "14:30:16"
    assert float(_node(wb, "A5")[0]) == 1.1574074074074074e-08     # one millisecond
    assert _node(wb, "A6")[0] == "0"
    # The next midnight: a date, given the short date as one.
    assert (wb.cell("Log!A7").t, wb.cell("Log!A7").v, _node(wb, "A7")[1]) == \
        ("date", "2026-01-06", 14)
    error = _refused(wb, {"Log!A8": dt.time(23, 59, 59, 999600)})
    assert error.code == "bad_date" and "24:00" in str(error)


def test_a_midnight_date_and_time_is_its_date():
    wb = _open()
    written = wb.set_cells({"Log!A2": dt.datetime(2026, 1, 5), "Log!A3": {"date": "2026-01-05 00:00"}})
    assert written.short_date == ("Log!A2", "Log!A3") and not written.time_format
    assert wb.cell("Log!A2").t == wb.cell("Log!A3").t == "date"


# -- Excel's format for what is typed -----------------------------------------------------------


def test_a_general_cell_gets_excel_s_date_time_or_time_format():
    wb = _open()
    written = wb.set_cells({"Log!A2": dt.datetime(2026, 1, 5, 14, 30),
                            "Log!A3": dt.datetime(2026, 1, 5, 14, 30, 15),
                            "Log!A4": dt.time(14, 30), "Log!A5": dt.time(14, 30, 15),
                            "Log!A6": dt.time(14, 30, 0, 250_000)})
    assert [_node(wb, f"A{n}")[1] for n in range(2, 7)] == [22, 22, 20, 21, 21]
    assert written.time_format == ("Log!A2", "Log!A3", "Log!A4", "Log!A5", "Log!A6")
    assert written.as_dict()["time_format"] == list(written.time_format)


def test_a_cell_s_format_is_kept_replaced_or_refused_as_excel_does_for_a_typed_value():
    wb = _open()
    moment, clock = dt.datetime(2026, 1, 5, 14, 30), dt.time(9, 15)
    written = wb.set_cells({"Log!B1": moment,      # the short date: kept, the day shown
                            "Log!C1": moment,      # h:mm: Excel gives the date and time
                            "Log!E1": clock,       # m/d/yyyy h:mm: Excel gives h:mm
                            "Log!G1": clock})      # [h]:mm:ss: a time format, kept
    assert [_node(wb, ref)[1] for ref in ("B1", "C1", "E1", "G1")] == [14, 22, 20, 46]
    assert written.time_format == ("Log!C1", "Log!E1")
    for ref, value, said in (("D1", moment, "serial number"), ("F1", moment, "serial number"),
                             ("I1", moment, "only a time"), ("H1", clock, "day 0"),
                             ("D1", clock, "serial number")):
        error = _refused(wb, {f"Log!{ref}": value})
        assert error.code == "date_format" and said in str(error), ref
        assert '"format": ' in str(error)         # the call that would work


def test_a_format_given_is_honoured_and_checked_against_the_value():
    wb = _open()
    wb.set_cells({"Log!D1": {"date": dt.datetime(2026, 1, 5, 14, 30, 15),
                             "format": "yyyy-mm-dd hh:mm:ss"},
                  "Log!H1": {"date": "14:30", "format": "hh:mm"},
                  "Log!A2": {"date": "2026-01-05 14:30", "format": "d mmm yyyy"}})
    assert _node(wb, "D1")[2] == "yyyy-mm-dd hh:mm:ss"
    assert _node(wb, "H1")[2] == "hh:mm"
    assert wb.cell("Log!D1").v == "2026-01-05T14:30:15"
    assert _refused(wb, {"Log!A3": {"date": "14:30", "format": "yyyy-mm-dd"}}).code \
        == "date_format"
    assert _refused(wb, {"Log!A3": {"date": "2026-01-05 14:30", "format": "0.00"}}).code \
        == "date_format"


# -- text and refusals --------------------------------------------------------------------------


def test_date_time_text_is_iso_or_in_the_locale_s_order_with_a_clock():
    wb = _open()
    wb.locale = NL_BE
    wb.set_cells({"Log!A2": {"date": "5/1/2026 14:30"}, "Log!A3": {"date": "2026-01-05T14:30"},
                  "Log!A4": {"date": " 7:05 "}})
    assert wb.cell("Log!A2").v == wb.cell("Log!A3").v == "2026-01-05T14:30:00"
    assert wb.cell("Log!A4").v == "07:05:00"
    for text in ("25:30", "14:60", "2026-01-05 14:30:61", "2:30 PM", "2026-01-05 14", "14"):
        assert _refused(wb, {"Log!A5": {"date": text}}).code == "bad_date", text


def test_a_time_zone_and_a_duration_are_refused_with_the_way_out():
    wb = _open()
    zone = dt.timezone(dt.timedelta(hours=1))
    for value in (dt.datetime(2026, 1, 5, 14, 30, tzinfo=zone), dt.time(14, 30, tzinfo=zone),
                  {"date": dt.datetime(2026, 1, 5, tzinfo=zone)}):
        error = _refused(wb, {"Log!A2": value})
        assert error.code == "date_time" and "tzinfo" in str(error)
    error = _refused(wb, {"Log!A2": dt.timedelta(hours=26)})
    assert error.code == "date_time" and "duration" in str(error)


# -- a block of rows ----------------------------------------------------------------------------


def test_write_table_and_append_rows_take_date_times_bare_and_as_dates():
    wb = _open()
    wb.write_table("Log", "K1", ["Shift", "Start", "Clock"],
                   [["early", dt.datetime(2026, 1, 5, 6, 0), dt.time(6, 0)],
                    ["late", {"date": dt.datetime(2026, 1, 5, 14, 30)}, {"date": "14:30:15"}]],
                   as_table="Shifts")
    wb.append_rows("t:Shifts", [["night", {"date": "2026-01-05 22:00"}, dt.time(22, 0)]])
    assert [wb.cell(f"Log!L{r}").v for r in (2, 3, 4)] == \
        ["2026-01-05T06:00:00", "2026-01-05T14:30:00", "2026-01-05T22:00:00"]
    assert [wb.cell(f"Log!M{r}").v for r in (2, 3, 4)] == ["06:00:00", "14:30:15", "22:00:00"]
    assert [_node(wb, f"L{r}")[1] for r in (2, 3, 4)] == [22, 22, 22]
    # The appended row's cell takes the format above it (h:mm:ss), as a table's new row does
    # in Excel, which shows a time: kept.
    assert [_node(wb, f"M{r}")[1] for r in (2, 3, 4)] == [20, 21, 21]


def test_a_block_gives_a_cell_whose_format_would_refuse_excel_s_format_for_the_value():
    wb = _open()
    # D1 is 0.00, H1 yyyy-mm-dd: set_cells refuses these, a block of rows formats them.
    wb.write_table("Log", "D1", None, [[dt.datetime(2026, 1, 5, 14, 30)]], overwrite=True)
    wb.write_table("Log", "H1", None, [[dt.time(14, 30)]], overwrite=True)
    assert _node(wb, "D1")[1] == 22 and _node(wb, "H1")[1] == 20


# -- the tool -----------------------------------------------------------------------------------


def test_the_tool_writes_a_date_and_time_and_a_time():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        blob = session.add_blob(_book(), name="log.xlsx")
        assert box.dispatch(session, "open_document", {"blob": blob}).ok
        result = box.dispatch(session, "xlsx_set_cells", {"doc": "d1", "cells": [
            {"address": "Log!A2", "date": "2026-01-05 14:30"},
            {"address": "Log!A3", "date": "14:30"}]})
        assert result.ok, result.to_json()
        assert result.data["written"] == 2
        assert result.data["time_format"] == ["Log!A2", "Log!A3"]
