"""The Excel oracle on notes and document properties (round 44, lane C: R57, R62).

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven; a held lock or
an Excel in use skips).  Each made-up workbook (tests/notes_book.py) is given notes and
properties as a program would, and Excel for Mac must open it without asking to repair
it, and its own re-save must keep what was written: every note's cell, text and author,
each shape where it was put, a threaded comment beside a note, a comment another program
wrote without a shape, and the core and custom properties with their types.
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from datetime import date

import pytest
from lxml import etree

from xlsx_agent import Workbook

import notes_book
import oracle

V = "{urn:schemas-microsoft-com:vml}"
XL = "{urn:schemas-microsoft-com:office:excel}"


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


def _saved(excel, wb: Workbook, name: str) -> bytes:
    outcome = excel.resave(wb.to_bytes(), name=name)
    assert outcome.outcome == "ok", (outcome.outcome, outcome.detail)
    return outcome.saved.read_bytes()


def _notes(data: bytes, sheet: str) -> dict:
    return {(c.cell.row, c.cell.column): (c.text, c.author)
            for c in Workbook.open(data).comments(sheet)}


def _anchors(data: bytes) -> set[tuple]:
    """Every note shape of the file: (row, column, margin-left, margin-top, anchor)."""
    out = set()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in archive.namelist():
            if not name.endswith(".vml"):
                continue
            for shape in etree.fromstring(archive.read(name)).iter(V + "shape"):
                client = shape.find(XL + "ClientData")
                style = shape.get("style")
                out.add((int(client.find(XL + "Row").text),
                         int(client.find(XL + "Column").text),
                         int(re.search(r"margin-left:(\d+)pt", style).group(1)),
                         int(re.search(r"margin-top:(\d+)pt", style).group(1)),
                         tuple(int(n) for n in client.find(XL + "Anchor").text.split(","))))
    return out


@pytest.mark.oracle
def test_notes_written_replaced_and_removed_are_excel_s(excel):
    wb = Workbook.open(notes_book.plain())
    wb.set_comments({address: f"note on {address}" for address in notes_book.MEASURED},
                    author="Ann Example")
    wb.set_comment("Stock!B3", "Line one\nand two & <three> \x01")
    wb.remove_comment("Orders!C2")
    saved = _saved(excel, wb, "c44-notes")
    assert _notes(saved, "Stock") == _notes(wb.to_bytes(), "Stock")
    assert _notes(saved, "Orders") == _notes(wb.to_bytes(), "Orders")
    assert _anchors(saved) == _anchors(wb.to_bytes())


@pytest.mark.oracle
def test_notes_over_hidden_lines(excel):
    wb = Workbook.open(notes_book.hidden())
    wb.set_comments({address: "note" for address in notes_book.HIDDEN})
    saved = _saved(excel, wb, "c44-hidden")
    assert _anchors(saved) == _anchors(wb.to_bytes())


@pytest.mark.oracle
def test_a_note_beside_a_threaded_comment(excel):
    wb = Workbook.open(notes_book.threaded())
    wb.set_comment("Stock!C3", "beside it")
    saved = _saved(excel, wb, "c44-threaded")
    with zipfile.ZipFile(io.BytesIO(saved)) as archive:
        threads = [n for n in archive.namelist() if "threadedComment" in n]
        assert threads and b'ref="B2"' in archive.read(threads[0])
    notes = _notes(saved, "Stock")
    assert notes[(3, 3)] == ("beside it", "")
    assert (2, 2) in notes


@pytest.mark.oracle
def test_a_comment_without_a_shape_survives_once_shaped(excel):
    wb = Workbook.open(notes_book.unshaped())
    change = wb.set_comment("Stock!B5", "beside old")
    assert change.shaped == ("Stock!A2",)
    saved = _saved(excel, wb, "c44-unshaped")
    assert _notes(saved, "Stock") == {(2, 1): ("old note", "Clerk"), (5, 2): ("beside old", "")}


@pytest.mark.oracle
def test_notes_moved_with_their_rows(excel):
    wb = Workbook.open(notes_book.plain())
    wb.set_comments({"Stock!B3": "three", "Stock!B5": "five"})
    wb.insert_rows("Stock", 2, 2)
    wb.delete_rows("Stock", 5)
    saved = _saved(excel, wb, "c44-moved")
    assert _notes(saved, "Stock") == {(6, 2): ("five", "")}
    # The anchors are Excel's; a moved shape's margins are not rewritten (restructure moves
    # x:Anchor only), and Excel set them from the anchor on saving (measured).
    assert {(r, c, a) for r, c, _, _, a in _anchors(saved)} == \
        {(r, c, a) for r, c, _, _, a in _anchors(wb.to_bytes())}


@pytest.mark.oracle
def test_every_note_removed(excel):
    wb = Workbook.open(notes_book.plain())
    wb.set_comments({"Stock!B3": "a", "Orders!A1": "b"})
    wb.remove_comment(["Stock!", "Orders!"])
    saved = _saved(excel, wb, "c44-none")
    assert _notes(saved, "Stock") == {} and _notes(saved, "Orders") == {}


@pytest.mark.oracle
def test_properties_are_read_and_kept(excel):
    wb = Workbook.open(notes_book.plain())
    wb.set_properties(title="Stock list", subject="Counts", creator="Bea", keywords="a; b",
                      description="Made up", category="Lists", content_status="Draft",
                      custom={"Project": "Alpha", "Run": 42, "Ratio": 2.5, "Checked": True,
                              "Due": date(2026, 3, 14), "Big": 2 ** 40})
    saved = Workbook.open(_saved(excel, wb, "c44-properties")).properties()
    ours = wb.properties()
    for name in ("title", "subject", "creator", "keywords", "description", "category",
                 "content_status", "created"):
        assert saved.core[name] == ours.core[name], name
    assert saved.custom == ours.custom
