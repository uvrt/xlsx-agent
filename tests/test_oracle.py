"""The Excel oracle (E1): Excel for Mac opens every edited fixture without a repair prompt.

Opt in with ``pytest -m oracle`` (never under xdist; tests/oracle.py says how Excel is
driven).  Every fixture is the corner shop's stock list (``tests/shop.py``), in each of its
variants: as made, after one step writing every kind of value (text, ``"=..."`` text, a
number, a boolean, a date, a blank, a new formula, a new row), and undone again -- the
original bytes, asked of Excel again under its own name.  Excel's
re-saved file must keep what was written: the values, the date's serial, the formula, the
``x14`` validations and the date system.  A workbook with a shared-string index out of range
is the control: Excel must ask to repair it, or the oracle could not see a repair at all.

Without Excel, or with Excel in use or its lock held, every test here skips.  The lock tests
at the end run in the default suite, without Excel.
"""

from __future__ import annotations

import datetime as dt
import io
import sys
import zipfile

import pytest

from ooxml_edit.sheets import Formula

from xlsx_agent import Workbook

import oracle
from shop import Shop

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
            wb._write(address, value)
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
