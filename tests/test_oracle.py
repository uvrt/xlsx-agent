"""The Excel oracle (E1): Excel for Mac opens every edited fixture without a repair prompt.

Opt in with ``pytest -m oracle`` (never under xdist; tests/oracle.py says how Excel is
driven).  Every fixture is the corner shop's stock list (``tests/shop.py``), in each of its
variants: as made, after one step writing every kind of value (text, ``"=..."`` text, a
number, a boolean, a date, a blank, a new formula, a new row), and undone again -- the
original bytes, asked of Excel again under its own name.  Excel's
re-saved file must keep what was written: the values, the date's serial, the formula, the
``x14`` validations and the date system.  A workbook with a shared-string index out of range
is the control: Excel must ask to repair it, or the oracle could not see a repair at all.

E3 (X3): every variant edited with ``set_cells`` -- every kind of value, a date given a
format, a shared formula group's master written over (the group handed to its next member)
and a member given its own formula -- opens without a repair, and Excel's re-save keeps the
values and types and computes each member's line value from the formula it kept.

X4: rows written by key into a declared table's capacity (a value by key, two new rows,
under a write policy) open without a repair, and Excel's re-save keeps them in the table.

Without Excel, or with Excel in use or its lock held, every test here skips.  The lock tests
at the end run in the default suite, without Excel.
"""

from __future__ import annotations

import datetime as dt
import io
import json
import os
import re
import signal
import subprocess
import sys
import time
import zipfile

import pytest

from ooxml_edit.sheets import Cell, Formula

from xlsx_agent import Workbook

import oracle
from shop import FIRST_INPUT, ITEMS, Shop

VARIANTS = {"plain": {}, "rless": {"rless": True}, "1904": {"date1904": True},
            "inline": {"shared": False}, "table": {"table": True}}

WRITES = {
    "Stock!B10": "Spare fan",
    "Stock!C11": "=not a formula",
    "Stock!D10": 12,
    "Stock!G12": True,
    "Stock!H10": dt.date(2026, 10, 20),
    "Stock!B4": None,
    "Stock!J10": Formula("D10*2"),
    "Stock!B15": "Torch, spare",
    "Stock!D15": 2,
}


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


def _edited(options: dict) -> tuple[bytes, bytes, bytes]:
    original = Shop(**options).to_bytes()
    wb = Workbook.open(original)
    with wb.batch():
        for address, value in WRITES.items():
            wb.set_cells({address: value})
    edited = wb.to_bytes()
    wb.undo()
    return original, edited, wb.to_bytes()


@pytest.mark.oracle
@pytest.mark.parametrize("variant", VARIANTS)
def test_excel_opens_every_fixture_edited_and_undone_without_a_repair(excel, variant):
    options = VARIANTS[variant]
    name = f"shop-{variant}"
    original, edited, undone = _edited(options)
    assert undone == original
    for tag, data in (("original", original), ("edited", edited), ("undone", undone)):
        outcome = excel.resave(data, name=f"{name}-{tag}")
        assert outcome.ok, (tag, outcome)
    saved = excel.resave(edited, name=f"{name}-edited")
    again = Workbook.open(saved.saved.read_bytes())
    assert again.cell("Stock!B10").v == "Spare fan"
    assert again.cell("Stock!C11").v == "=not a formula" and again.cell("Stock!C11").t == "text"
    assert again.cell("Stock!D10").v == 12 and again.cell("Stock!G12").v is True
    assert again.cell("Stock!H10").v == "2026-10-20"
    assert again.cell("Stock!J10").formula and again.cell("Stock!B15").v == "Torch, spare"
    assert again.date1904 == bool(options.get("date1904"))
    with zipfile.ZipFile(io.BytesIO(saved.saved.read_bytes())) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml")
    assert b"dataValidation" in sheet and b"Lists!$A$2:$A$5" in sheet


#: What E3's set_cells writes on every variant: each kind of value, a date given its format
#: in a new cell, a shared formula group's master and a member written over (the group
#: handed to the next member), a formula replaced, "=..." text.
SET_CELLS = {
    "Stock!B15": "Spare fan",
    "Stock!C11": "=still text",
    "Stock!D15": 4,
    "Stock!E15": 1.25,
    "Stock!G15": False,
    "Stock!H10": {"date": "2026-10-20"},
    "Stock!H15": {"date": "2026-10-21", "format": "d mmm yyyy"},
    "Stock!B4": None,
    f"Stock!F{FIRST_INPUT}": 50,
    f"Stock!F{FIRST_INPUT + 2}": {"formula": f"=D{FIRST_INPUT + 2}*E{FIRST_INPUT + 2}*2"},
    "Stock!I11": {"formula": "=0.5"},
    "Stock!J15": {"formula": "=D15*E15"},
}

X3_VARIANTS = {**VARIANTS, "array": {"array": True}}


@pytest.mark.oracle
@pytest.mark.parametrize("variant", X3_VARIANTS)
def test_excel_opens_every_set_cells_fixture_and_keeps_what_was_written(excel, variant):
    options = X3_VARIANTS[variant]
    original = Shop(**options).to_bytes()
    wb = Workbook.open(original)
    written = wb.set_cells(SET_CELLS, replace_formulas=True)
    assert written.released == (f"Stock!F{FIRST_INPUT}",)
    edited = wb.to_bytes()
    wb.undo()
    assert wb.to_bytes() == original
    outcome = excel.resave(edited, name=f"shop-x3-{variant}")
    assert outcome.ok, outcome
    again = Workbook.open(outcome.saved.read_bytes())
    assert again.cell("Stock!B15").v == "Spare fan"
    assert (again.cell("Stock!C11").v, again.cell("Stock!C11").t) == ("=still text", "text")
    assert again.cell("Stock!D15").v == 4 and again.cell("Stock!E15").v == 1.25
    assert again.cell("Stock!G15").v is False
    assert again.cell("Stock!H10").v == "2026-10-20" and again.cell("Stock!H15").v \
        == "2026-10-21"
    assert again.cell("Stock!H15").t == "date"
    assert again.cell("Stock!B4").t == "blank"
    assert again.cell(f"Stock!F{FIRST_INPUT}").v == 50
    assert not again.cell(f"Stock!F{FIRST_INPUT}").formula
    # The group handed on: Excel (asked to recalculate on load) computes every member's
    # own line value from the formula each kept.
    values = {v.row: v for v in again.cells(again.sheet("Stock")) if v.column == 6}
    for offset in (1, 3, 4):
        row = FIRST_INPUT + offset
        quantity, price = again.cell(f"Stock!D{row}").v, again.cell(f"Stock!E{row}").v
        assert values[row].formula and values[row].v == pytest.approx(quantity * price), row
    row = FIRST_INPUT + 2
    assert values[row].v == pytest.approx(again.cell(f"Stock!D{row}").v
                                          * again.cell(f"Stock!E{row}").v * 2)
    assert again.cell("Stock!J15").formula and again.cell("Stock!J15").v == pytest.approx(5)
    assert again.cell("Totals!B1").formula
    assert again.date1904 == bool(options.get("date1904"))
    with zipfile.ZipFile(io.BytesIO(outcome.saved.read_bytes())) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml")
    assert b"dataValidation" in sheet and b"Lists!$A$2:$A$5" in sheet


#: What X3b's Markdown writes: rich runs on one line, lines with a list (a wrapping format
#: given), a numbered list nested in a bullet, and a fixed-height row's cell.
MARKDOWN = {
    "Stock!B15": {"markdown": "**Spare** fan, *boxed* ~~old~~ ***new***"},
    f"Stock!B{FIRST_INPUT}": {"markdown": "**Fans**\n\n- desk\n- **floor**\n  1. tall\n  2. short"},
    f"Stock!C{FIRST_INPUT + 1}": {"markdown": "one\ntwo\nthree"},
    "Stock!A1": {"markdown": "Corner shop *stock*\nsecond line"},
}


def _run_fonts(data: bytes, text: str) -> list[tuple[str, dict]]:
    """The runs of the shared string whose text is ``text``: each run's text and its
    properties by name."""
    from lxml import etree

    x = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = etree.fromstring(archive.read("xl/sharedStrings.xml"))
    for item in root.iter(x + "si"):
        runs = item.findall(x + "r")
        if "".join(t.text or "" for r in runs for t in r.iter(x + "t")) == text:
            return [("".join(t.text or "" for t in r.iter(x + "t")),
                     {etree.QName(p).localname: p.get("val", "") for p in
                      (r.find(x + "rPr") if r.find(x + "rPr") is not None else [])})
                    for r in runs]
    return []


@pytest.mark.oracle
@pytest.mark.parametrize("variant", ["plain", "inline"])
def test_excel_opens_markdown_cells_and_keeps_the_runs_lines_and_wrap(excel, variant):
    original = Shop(**VARIANTS[variant]).to_bytes()
    wb = Workbook.open(original)
    written = wb.set_cells(MARKDOWN)
    assert written.fixed_height == ("Stock!A1",)
    edited = wb.to_bytes()
    outcome = excel.resave(edited, name=f"shop-x3b-{variant}")
    assert outcome.ok, outcome
    saved = outcome.saved.read_bytes()
    again = Workbook.open(saved)
    for address in MARKDOWN:
        assert again.cell(address).v == wb.cell(address).v, address
        assert again.cell(address, markdown=True).v == wb.cell(address, markdown=True).v
    runs = _run_fonts(saved, "Spare fan, boxed old new")
    assert [text for text, _ in runs] == ["Spare", " fan, ", "boxed", " ", "old", " ", "new"]
    for text, props in runs:     # Excel kept the cell's font on every run, and the marks
        assert props.get("rFont") == "Aptos" and props.get("sz") == "11", (text, props)
    assert [("b" in p, "i" in p, "strike" in p) for _, p in runs] == [
        (True, False, False), (False, False, False), (False, True, False),
        (False, False, False), (False, False, True), (False, False, False),
        (True, True, False)]
    styles = again.editor.cells.styles
    sheet = again.editor.cells.sheet("Stock")
    for ref in (f"B{FIRST_INPUT}", f"C{FIRST_INPUT + 1}", "A1"):
        assert styles.wraps_text(sheet.style(Cell.parse(ref))), ref
    assert not styles.wraps_text(sheet.style(Cell.parse("B15")))
    with zipfile.ZipFile(io.BytesIO(saved)) as archive:
        rows = archive.read("xl/worksheets/sheet1.xml")
    # Measured: Excel grows a row without a height of its own to its lines on open and
    # saves the height; a row with customHeight keeps its own.
    import re

    heights = {int(m.group(1)): m.group(0) for m in re.finditer(
        rb'<row r="(\d+)"[^>]*>', rows)}
    first = heights[1].decode()
    assert 'customHeight="1"' in first and 'ht="21"' in first
    grown = heights[FIRST_INPUT].decode()
    assert 'ht="' in grown and float(re.search(r'ht="([0-9.]+)"', grown).group(1)) > 40


#: What X4's writes by key put in the stock table (declared over the input rows, its key
#: Item, room for 4 more rows): a value by key, and two new rows in its capacity.
KEYED = {"t:stock/Kettle/Quantity": 7, "t:stock/Lamp/Quantity": 2,
         "t:stock/Lamp/Category": "Hardware", "t:stock/Fan/Unit price": 19.5}


@pytest.mark.oracle
@pytest.mark.parametrize("variant", ["plain", "rless", "inline"])
def test_excel_opens_rows_written_by_key_into_a_table_s_capacity(excel, variant):
    original = Shop(**VARIANTS[variant]).to_bytes()
    wb = Workbook.open(original, policy={"tables": [
        {"table": "stock", "columns": ["Quantity", "Category", "Unit price"],
         "new_rows": True}]})
    wb.declare_tables([{"name": "stock", "sheets": "Stock", "header_row": 9, "key": "Item",
                        "capacity": len(ITEMS) + 4}])
    written = wb.set_cells(KEYED, new_rows=True)
    assert written.new_rows == ("Stock!B15", "Stock!B16")
    edited = wb.to_bytes()
    outcome = excel.resave(edited, name=f"shop-x4-{variant}")
    assert outcome.ok, outcome
    again = Workbook.open(outcome.saved.read_bytes())
    assert again.cell("Stock!D12").v == 7
    assert [again.cell(f"Stock!{c}15").v for c in "BCD"] == ["Lamp", "Hardware", 2]
    assert again.cell("Stock!B16").v == "Fan" and again.cell("Stock!E16").v == 19.5
    again.declare_tables([{"name": "stock", "sheets": "Stock", "header_row": 9,
                           "key": "Item", "capacity": len(ITEMS) + 4}])
    assert again.table("t:stock").rows == len(ITEMS) + 2


@pytest.mark.oracle
def test_excel_asks_to_repair_a_broken_shared_string_index(excel):
    data = Shop().to_bytes()
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, \
            zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            part = source.read(info)
            if info.filename == "xl/worksheets/sheet2.xml":
                part = part.replace(b'<c r="A2" t="s"><v>', b'<c r="A2" t="s"><v>9', 1)
            target.writestr(info, part)
    broken = buffer.getvalue()
    assert [p.code for p in Workbook.open(broken).validate()] == ["shared_string"]
    outcome = excel.resave(broken, name="shop-broken-index")
    if outcome.outcome == "hang" and "screen is locked" in outcome.detail:
        pytest.skip("the screen is locked and the window server showed no prompt")
    assert outcome.outcome == "repair", outcome


#: Two workbooks Excel asks to repair (measured): sheet names it treats as equal, cells
#: out of column order.
def _unrepaired() -> dict[str, bytes]:
    from build import WorkbookBuilder

    book = WorkbookBuilder()
    for name in ("Tally", "TALLY"):
        book.sheet(name).cell("A1", 1)
    swapped = Shop().to_bytes()
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(swapped)) as source, \
            zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            part = source.read(info)
            if info.filename == "xl/worksheets/sheet3.xml":     # Totals: A1, B1 swapped
                part = re.sub(rb'(<c r="A1"[^>]*>.*?</c>)(<c r="B1"[^>]*>.*?</c>)',
                              rb"\2\1", part, count=1)
            target.writestr(info, part)
    return {"names": book.to_bytes(), "order": buffer.getvalue()}


@pytest.mark.oracle
def test_excel_asks_to_repair_a_known_cause_and_the_next_file_is_the_input(excel):
    known = _unrepaired()
    assert oracle.repair_causes(known["order"]) == [
        "cells out of order in sheet 'Totals': B1 listed before A1"]
    told = excel.resave(known["names"], name="shop-known-names")
    assert told.outcome == "preflight" and "'Tally' and 'TALLY'" in told.detail, told
    outcome = excel.resave(known["order"], name=f"shop-known-order-{time.time_ns()}",
                           preflight=False)
    assert outcome.outcome == "repair", outcome
    # After the prompt (unanswered with the screen locked: Excel quit by signal), the next
    # file is opened in a fresh Excel, and what it saved is that file.
    data = Shop(rless=True).to_bytes()
    again = excel.resave(data, name=f"shop-after-repair-{time.time_ns()}")
    assert again.ok, again
    assert oracle.sheet_names(again.saved.read_bytes()) == oracle.sheet_names(data)


# -- the lock, without Excel ---------------------------------------------------------------------


@pytest.fixture()
def container(tmp_path, monkeypatch):
    pytest.importorskip("fcntl")
    monkeypatch.setattr(oracle, "GROUP_CONTAINER", tmp_path)
    monkeypatch.setattr(oracle, "LOCK", tmp_path / "excel-oracle.lock")
    monkeypatch.setattr(oracle, "ORACLE_DIR", tmp_path / "oracle")
    monkeypatch.setattr(oracle, "available", lambda: True)
    return tmp_path


def test_a_held_lock_is_excel_in_use(container):
    import fcntl

    with open(oracle.LOCK, "a+") as other:
        fcntl.flock(other, fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            with pytest.raises(oracle.ExcelBusy, match="lock"):
                with oracle.session():
                    pass
        finally:
            fcntl.flock(other, fcntl.LOCK_UN)


def test_a_running_excel_is_excel_in_use_and_never_quit(container, monkeypatch):
    monkeypatch.setattr(oracle, "excel_pids", lambda: {4242})
    quit_ = []
    monkeypatch.setattr(oracle.Session, "_quit", lambda self, pids: quit_.append(pids))
    with pytest.raises(oracle.ExcelBusy, match="Excel is in use"):
        with oracle.session():
            pass
    assert quit_ == []


def test_recovery_quits_only_what_the_session_started(container, monkeypatch):
    monkeypatch.setattr(oracle, "SETTLE", 0.0)
    running = {101, 202}
    monkeypatch.setattr(oracle, "excel_pids", lambda: set(running))
    quit_ = []

    def fake_quit(self, pids):
        quit_.append(set(pids))
        running.difference_update(pids)

    monkeypatch.setattr(oracle.Session, "_quit", fake_quit)
    session = oracle.Session()
    session.started = {202}
    oracle.ORACLE_DIR.mkdir()
    session.recover()
    assert quit_ == [{202}] and running == {101}


def test_after_a_killed_script_an_excel_that_starts_again_is_quit(container, monkeypatch):
    monkeypatch.setattr(oracle, "SETTLE", 1.0)
    monkeypatch.setattr(oracle.time, "sleep", lambda seconds: None)
    monkeypatch.setattr(oracle, "_start_time", lambda pid: f"start {pid}")
    seen = iter([set(), set(), {303}, {303}])        # a queued AppleEvent starts Excel again
    state = {"running": set()}

    def pids():
        try:
            state["running"] |= next(seen)
        except StopIteration:
            pass
        return set(state["running"])

    monkeypatch.setattr(oracle, "excel_pids", pids)
    quit_ = []

    def fake_quit(self, found):
        quit_.append(set(found))
        state["running"] -= found

    monkeypatch.setattr(oracle.Session, "_quit", fake_quit)
    session = oracle.Session()
    session.orphaned = True
    oracle.ORACLE_DIR.mkdir()
    session.recover()
    assert quit_ == [{303}] and not state["running"] and not session.orphaned



# -- before and after Excel, without it ---------------------------------------------------------


def _book(*names: str) -> bytes:
    from build import WorkbookBuilder

    book = WorkbookBuilder()
    for n, name in enumerate(names):
        book.sheet(name).cell("A1", n)
    return book.to_bytes()


@pytest.mark.parametrize("pair", [
    ("R", "r"), ("XFD9", "xfd9"), ("a\u00b9b", "a\uff11b"), ("a\u00adb", "a\ufeffb"),
    ("ab", "a\u200bb"), ("a b", "a\u00a0b"), ("\u30ab", "\uff76")])
def test_sheet_names_excel_treats_as_equal_are_a_known_cause(pair):
    first, second = pair
    assert oracle.same_name(first) == oracle.same_name(second)
    assert oracle.repair_causes(_book("Stock", first, second)) == [
        f"sheet names Excel treats as equal: {first!r} and {second!r}"]


@pytest.mark.parametrize("pair", [("a-b", "ab"), ("a.b", "a b"), ("x1", "x2"),
                                  ("\u00fcber", "uber")])
def test_sheet_names_excel_tells_apart_are_no_cause(pair):
    assert oracle.repair_causes(_book(*pair)) == []


@pytest.mark.parametrize("options", [{}, {"rless": True}, {"shared": False}, {"table": True},
                                     {"orders": True}])
def test_the_shop_has_no_known_cause_of_a_repair(options):
    data = Shop(**options).to_bytes()
    assert oracle.repair_causes(data) == []
    assert oracle.sheet_names(data)[0] == "Stock"


def test_cells_out_of_column_order_are_a_known_cause():
    assert oracle.repair_causes(_unrepaired()["order"]) == [
        "cells out of order in sheet 'Totals': B1 listed before A1"]


def test_a_cell_without_r_follows_the_cell_before_it():
    sheet = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             b'<sheetData><row r="1"><c r="B1"/><c/><c r="D1"/></row>'
             b'<row><c r="C2"/><c/><c r="D2"/></row></sheetData></worksheet>')
    assert oracle._out_of_order(io.BytesIO(sheet)) == "D2 listed twice"
    sheet = sheet.replace(b'<c r="D2"/>', b'<c r="E2"/>')
    assert oracle._out_of_order(io.BytesIO(sheet)) == ""
    sheet = sheet.replace(b'<c r="E2"/>', b'<c r="AA2"/><c r="Z2"/>')
    assert oracle._out_of_order(io.BytesIO(sheet)) == "AA2 listed before Z2"


@pytest.mark.parametrize("rows, found", [
    (b'<row r="2"/><row r="1"/>', "row 2 listed before row 1"),
    (b'<row r="3"/><row/><row r="4"/>', "row 4 listed twice"),
    (b'<row r="1"/><row r="1"/>', "row 1 listed twice"),
    (b'<row/><row/><row r="5"/>', "")])
def test_rows_out_of_order_are_a_known_cause(rows, found):
    sheet = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             b'<sheetData>' + rows + b'</sheetData></worksheet>')
    assert oracle._out_of_order(io.BytesIO(sheet)) == found


def test_a_file_that_is_no_workbook_has_no_known_cause():
    assert oracle.repair_causes(b"not a zip") == [] and oracle.sheet_names(b"") is None


EXPECTED = ["Stock", "Lists", "Totals"]


@pytest.mark.parametrize("given, outcome, detail", [
    (("done", "", EXPECTED), "ok", ""),
    (("done\n", "", ["Blad1"]), "mismatch",
     "Excel saved the sheets ['Blad1']; the input has ['Stock', 'Lists', 'Totals']"),
    (("wrong workbook: Map1", "", None), "mismatch",
     "Excel has no workbook of the input's name open (it has 'Map1'); nothing saved"),
    (("done", "", None), "error", "Excel reported a save but the saved file cannot be read"),
    (("error -1712: Excel got an error: AppleEvent timed out.", "", None), "hang",
     "Excel did not answer (AppleEvent time-out), with no window open: no known cause in "
     "the file"),
    (("error -50: Parameter error.", "", None), "error", "error -50: Parameter error."),
    (("", "execution error: boom (-2700)", None), "error", "execution error: boom (-2700)"),
])
def test_what_the_script_says_is_an_outcome(given, outcome, detail):
    out, err, saved = given
    assert oracle.judge(out, err, saved=saved, expected=EXPECTED, windows=0) \
        == (outcome, detail)
    assert outcome in oracle.OUTCOMES


def test_a_hang_names_the_known_cause_and_what_excel_showed():
    assert oracle.judge("", "", saved=None, expected=None, timed_out=True, windows=2,
                        causes=["cells out of order in sheet 'S': B1 listed before A1"]) \
        == ("hang", "osascript did not return within 90 s, with 2 windows open: cells out "
            "of order in sheet 'S': B1 listed before A1 (Excel asks to repair such a file)")
    assert oracle.judge("error -1712: x", "", saved=None, expected=None)[1] \
        == "Excel did not answer (AppleEvent time-out): no known cause in the file"


def test_a_hang_with_the_screen_locked_says_a_prompt_may_not_have_been_seen():
    outcome, detail = oracle.judge("error -1712: x", "", saved=None, expected=None,
                                   windows=0, locked=True)
    assert outcome == "hang" and "no window open" not in detail
    assert detail.startswith("Excel did not answer (AppleEvent time-out): no known cause in "
                             "the file; the screen is locked")


@pytest.mark.parametrize("plist, locked", [
    (b"<key>CGSSessionScreenIsLocked</key>\n\t\t\t<true/>", True),
    (b"<key>CGSSessionScreenIsLocked</key><false/>", False), (b"<dict></dict>", False)])
def test_the_screen_lock_is_read_from_ioreg(monkeypatch, plist, locked):
    monkeypatch.setattr(oracle.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(
        a, 0, stdout=plist.decode(), stderr=""))
    assert oracle.screen_locked() is locked


def test_a_known_cause_is_not_given_to_excel(container, monkeypatch):
    def never(*args):
        raise AssertionError("Excel was run")

    monkeypatch.setattr(oracle.Session, "_run", never)
    told = oracle.Session().resave(_book("Data", "DATA"), name="pair")
    assert told.outcome == "preflight" and not told.cached
    assert told.detail == ("sheet names Excel treats as equal: 'Data' and 'DATA' (Excel asks "
                           "to repair such a file; not run: preflight=False runs it)")


def test_a_known_cause_is_given_to_excel_when_asked_and_what_is_not_ok_or_repair_is_not_kept(container,
                                                                           monkeypatch):
    seen = []

    def hang(self, staged, saved, expected, causes):
        seen.append((expected, list(causes)))
        return oracle.Outcome("hang", "named")

    monkeypatch.setattr(oracle.Session, "_run", hang)
    oracle.ORACLE_DIR.mkdir()
    data = _book("Data", "DATA")
    for _ in range(2):
        assert oracle.Session().resave(data, name="pair", preflight=False).outcome == "hang"
    assert len(seen) == 2 and seen[0][0] == ["Data", "DATA"] and seen[0][1]


# -- interrupted, without Excel ------------------------------------------------------------------


class _Script:
    """A stand-in for the ``osascript`` call: still running until killed."""

    def __init__(self):
        self.killed = False

    def poll(self):
        return -9 if self.killed else None

    def kill(self):
        self.killed = True

    def communicate(self, timeout=None):
        return "", ""


def test_an_interrupted_run_stops_the_script_and_owns_the_excel_it_started(container,
                                                                          monkeypatch):
    script = _Script()
    running: set[int] = set()
    monkeypatch.setattr(oracle.subprocess, "Popen", lambda *a, **k: script)
    monkeypatch.setattr(oracle, "_start_time", lambda pid: "Wed Oct 28 10:00:00 2026")
    monkeypatch.setattr(oracle, "screen_locked", lambda: False)

    def pids():
        return set(running)

    def dialog():
        running.add(707)            # Excel starts, then Ctrl-C
        raise KeyboardInterrupt

    monkeypatch.setattr(oracle, "excel_pids", pids)
    monkeypatch.setattr(oracle, "_dialog_buttons", dialog)
    oracle.ORACLE_DIR.mkdir()
    session = oracle.Session()
    with pytest.raises(KeyboardInterrupt):
        session._run(oracle.ORACLE_DIR / "in.xlsx", oracle.ORACLE_DIR / "out.xlsx")
    assert script.killed and session.orphaned and session.started == {707}
    record = json.loads((oracle.ORACLE_DIR / oracle.STARTED).read_text())
    assert record == {"owner": os.getpid(), "excel": {"707": "Wed Oct 28 10:00:00 2026"}}


@pytest.fixture()
def quits(container, monkeypatch):
    """Excel as a set of pids; quitting takes the ones asked for away."""
    monkeypatch.setattr(oracle, "SETTLE", 0.0)
    running: set[int] = set()
    quit_: list[set[int]] = []
    monkeypatch.setattr(oracle, "excel_pids", lambda: set(running))
    monkeypatch.setattr(oracle, "_start_time", lambda pid: f"start {pid}")

    def fake_quit(self, pids):
        quit_.append(set(pids))
        running.difference_update(pids)

    monkeypatch.setattr(oracle.Session, "_quit", fake_quit)
    return running, quit_


@pytest.mark.parametrize("stop", ["interrupt", "SIGTERM", "SIGHUP"])
def test_a_session_stopped_mid_run_quits_its_excel(quits, stop):
    running, quit_ = quits
    if stop != "interrupt" and signal.getsignal(getattr(signal, stop)) is not signal.SIG_DFL:
        pytest.skip(f"{stop} has a handler of the test runner's: the oracle leaves it alone")
    with pytest.raises(KeyboardInterrupt):
        with oracle.session() as current:
            running.add(808)
            current._note({808})
            assert (oracle.ORACLE_DIR / oracle.STARTED).exists()
            if stop == "interrupt":
                raise KeyboardInterrupt
            os.kill(os.getpid(), getattr(signal, stop))
            time.sleep(1)
    assert quit_ == [{808}] and not running and not oracle._LIVE
    assert not (oracle.ORACLE_DIR / oracle.STARTED).exists()
    assert signal.getsignal(signal.SIGTERM) is signal.SIG_DFL


def test_a_signal_while_excel_is_quit_waits_until_it_is(quits, monkeypatch):
    running, quit_ = quits
    session = oracle.Session()
    session.started = {909}
    running.add(909)
    oracle.ORACLE_DIR.mkdir()

    def quit_under_ctrl_c(self, pids):
        os.kill(os.getpid(), signal.SIGINT)
        quit_.append(set(pids))
        running.difference_update(pids)

    monkeypatch.setattr(oracle.Session, "_quit", quit_under_ctrl_c)
    with pytest.raises(KeyboardInterrupt):
        session.recover()
    assert quit_ == [{909}] and not running


def test_at_exit_an_unclosed_session_is_recovered(quits, monkeypatch):
    running, quit_ = quits
    session = oracle.Session()
    session.started = {111}
    running.update({111, 222})
    oracle.ORACLE_DIR.mkdir()
    monkeypatch.setattr(oracle, "_LIVE", [session])
    oracle._recover_live()
    assert quit_ == [{111}] and running == {222} and oracle._LIVE == []


def _dead_pid() -> int:
    done = subprocess.Popen(["true"] if sys.platform != "win32" else ["cmd", "/c", "rem"])
    done.wait()
    return done.pid


@pytest.mark.parametrize("case", ["ours", "owner alive", "restarted", "workbook open",
                                  "not answering", "no record"])
def test_an_excel_left_by_a_killed_run_is_quit_and_no_other(quits, monkeypatch, case):
    running, quit_ = quits
    running.add(515)
    oracle.ORACLE_DIR.mkdir()
    owner = os.getppid() if case == "owner alive" else _dead_pid()
    started = "start 0" if case == "restarted" else "start 515"
    if case != "no record":
        (oracle.ORACLE_DIR / oracle.STARTED).write_text(json.dumps(
            {"owner": owner, "excel": {"515": started}}))

    def workbooks(script, timeout=10):
        assert "count workbooks" in script and "is running" in script
        if case == "not answering":
            raise subprocess.TimeoutExpired(script, timeout)
        return "1" if case == "workbook open" else "0"

    monkeypatch.setattr(oracle, "_osascript", workbooks)
    if case == "ours":
        with oracle.session():
            assert quit_ == [{515}] and not running
        assert not (oracle.ORACLE_DIR / oracle.STARTED).exists()
    else:
        with pytest.raises(oracle.ExcelBusy, match="Excel is in use"):
            with oracle.session():
                pass
        assert quit_ == [] and running == {515}



@pytest.mark.parametrize("line, alert", [
    ("|0|273|274", True), ("|0|273.0|288.0", True), ("Map1|0|1412|842", False),
    ("Book1|0|273|274", False), ("|0|54|54", False), ("|3|273|274", False),
    ("|0|1440|900", False), ("garbage", False), ("", False)])
def test_an_alert_is_an_untitled_window_of_a_dialog_s_size(line, alert):
    assert oracle._is_alert(line) is alert


def test_with_the_screen_locked_a_prompt_the_window_server_shows_is_a_repair(
        container, monkeypatch):
    script = _Script()
    running: set[int] = set()
    monkeypatch.setattr(oracle.subprocess, "Popen", lambda *a, **k: script)
    monkeypatch.setattr(oracle, "excel_pids", lambda: set(running))
    monkeypatch.setattr(oracle, "_start_time", lambda pid: "then")
    monkeypatch.setattr(oracle, "screen_locked", lambda: True)
    monkeypatch.setattr(oracle, "_dialog_buttons", lambda: running.add(404) or [])
    monkeypatch.setattr(oracle, "_alert_shown", lambda: 404 in running)
    monkeypatch.setattr(oracle, "SETTLE", 0.0)
    polite = []
    monkeypatch.setattr(oracle, "_osascript", lambda script, timeout=10: polite.append(script))
    signalled = []

    def kill(args, check=False):
        signalled.append(args[1])
        running.clear()

    monkeypatch.setattr(oracle.subprocess, "run", kill)
    oracle.ORACLE_DIR.mkdir()
    session = oracle.Session()
    outcome = session._run(oracle.ORACLE_DIR / "in.xlsx", oracle.ORACLE_DIR / "out.xlsx")
    assert (outcome.outcome, outcome.detail) == ("repair", oracle.UNANSWERED)
    assert script.killed and polite == [] and signalled == ["-TERM"]
    assert not session.orphaned and not session.modal
