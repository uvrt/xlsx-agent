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

Without Excel, or with Excel in use or its lock held, every test here skips.  The lock tests
at the end run in the default suite, without Excel.
"""

from __future__ import annotations

import datetime as dt
import io
import sys
import zipfile

import pytest

from ooxml_edit.sheets import Cell, Formula

from xlsx_agent import Workbook

import oracle
from shop import FIRST_INPUT, Shop

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
    assert outcome.outcome == "repair", outcome


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
