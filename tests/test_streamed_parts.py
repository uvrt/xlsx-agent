"""A sheet read from its compressed ZIP entry, a piece at a time: the same as read whole.

ooxml-edit keeps every part it has not edited as its compressed entry, and a pass inflates
the sheet as it goes (``read.stream.SheetBytes``): what it holds is a window of the part,
never the part.  Each read here is compared with the same read of the part held whole --
on test_skim's generated sheets, through pieces and windows small enough to cut a row, a
tag or a closing tag anywhere -- and the memory a pass holds is measured.

Every workbook is made up here, with tests/build.py's parts.
"""

import io
import tracemalloc
import zipfile

import pytest

from build import NS_MAIN, WorkbookBuilder, zip_parts
from test_skim import RANGES, Sheet, _book, _same_scan

from ooxml_common.zipparts import ZipPart
from ooxml_edit.opc import OpcPackage
from xlsx_agent import Workbook
from xlsx_agent.package import OpenError
from xlsx_agent.read import stream
from xlsx_agent.read.stream import SheetBytes, _Facts

SHEET = "xl/worksheets/sheet1.xml"


def _reads(data: bytes) -> tuple:
    wb = Workbook.open(data)
    merges = wb.merges("Stock")
    within = wb.merges("Stock", "A1:D20")
    used = wb.used_range("Stock")
    values = list(wb.cells("Stock", formats=True, values_only=True))
    cells = list(wb.cells("Stock", formats=True))
    found = wb.scan("Stock")
    reads = [list(wb.read(f"Stock!{r}", formats=True)) for r in RANGES]
    texts = list(wb.find_text("crate"))
    return found, (merges, within, used, values, cells, reads, texts)


def _held(monkeypatch):
    """Every part read whole, as before parts were kept compressed."""
    monkeypatch.setattr(OpcPackage, "part_entry", lambda self, path: None)


def _small(monkeypatch, read: int, fill: int, ahead: int) -> None:
    monkeypatch.setattr(stream, "_READ", read)
    monkeypatch.setattr(stream, "_FILL", fill)
    monkeypatch.setattr(stream, "_AHEAD", ahead)


SIZES = [(7, 23, 5), (64, 300, 50), (1 << 16, 1 << 17, 1 << 15)]


@pytest.mark.parametrize("read,fill,ahead", SIZES)
@pytest.mark.parametrize("seed", range(10))
def test_a_streamed_sheet_reads_as_the_sheet_held_whole(monkeypatch, seed, read, fill, ahead):
    data = _book(Sheet(seed).xml())
    _small(monkeypatch, read, fill, ahead)
    streamed_scan, streamed = _reads(data)
    with monkeypatch.context() as held:
        _held(held)
        held_scan, whole = _reads(data)
    _same_scan(streamed_scan, held_scan)
    assert streamed == whole
    assert held_scan.cells > 10


@pytest.mark.parametrize("odd", ["comment", "pi", "xmlns", "cdata", "merge"])
def test_a_streamed_part_that_may_hide_a_tag_reads_as_held(monkeypatch, odd):
    data = _book(Sheet(3, odd=odd).xml())
    _small(monkeypatch, 11, 40, 9)
    streamed_scan, streamed = _reads(data)
    with monkeypatch.context() as held:
        _held(held)
        held_scan, whole = _reads(data)
    _same_scan(streamed_scan, held_scan)
    assert streamed == whole


def _entry(data: bytes) -> ZipPart:
    entry = OpcPackage.open(data).part_entry(SHEET)
    assert entry is not None
    return entry


@pytest.mark.parametrize("read", [3, 8, 9, 10, 64, 1 << 16])
@pytest.mark.parametrize("seed", range(10))
def test_the_facts_of_a_streamed_part_are_those_of_its_bytes(monkeypatch, seed, read):
    monkeypatch.setattr(stream, "_READ", read)
    data = _book(Sheet(seed).xml())
    entry = _entry(data)
    streamed = _Facts.of(SheetBytes(entry))
    whole = _Facts.of(SheetBytes(entry.read()))
    assert streamed == whole and streamed.plan == whole.plan
    assert whole.plan is not None and not whole.merges and not whole.xmlns


@pytest.mark.parametrize("odd", ["comment", "pi", "xmlns", "cdata", "merge"])
def test_odd_parts_have_the_same_facts_streamed(monkeypatch, odd):
    monkeypatch.setattr(stream, "_READ", 5)
    entry = _entry(_book(Sheet(3, odd=odd).xml()))
    assert _Facts.of(SheetBytes(entry)) == _Facts.of(SheetBytes(entry.read()))


@pytest.mark.parametrize("closing", ["</sheetData>", "</sheetData  \n >", "</x:sheetData>"])
def test_a_closing_tag_cut_anywhere_is_found(monkeypatch, closing):
    prefix = "x:" if "x:" in closing else ""
    ns = f'xmlns:x="{NS_MAIN}"' if prefix else f'xmlns="{NS_MAIN}"'
    rows = "".join(f'<{prefix}row r="{n}"><{prefix}c r="A{n}" s="1"/></{prefix}row>'
                   for n in range(1, 40))
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{prefix}worksheet {ns}>'
           f"<{prefix}sheetData>{rows}{closing}</{prefix}worksheet>").encode()
    whole = _Facts.of(SheetBytes(xml))
    assert whole.end is not None and whole.plan is not None
    entry = _entry(_part_book(xml))
    for read in range(1, 40):
        monkeypatch.setattr(stream, "_READ", read)
        assert _Facts.of(SheetBytes(entry)) == whole


def _part_book(xml: bytes) -> bytes:
    book = WorkbookBuilder()
    book.sheet("Stock").cell("A1", "crate")
    parts = book.parts()
    parts[SHEET] = xml
    return zip_parts(parts)


def _large(rows: int, styled: bool = True) -> bytes:
    """A sheet of ``rows`` rows of values, formatted to the bottom as R73's is (or not:
    no row holds only styles, and the part is parsed whole)."""
    body = []
    for n in range(1, rows + 1):
        body.append(f'<row r="{n}"><c r="A{n}"><v>{n}</v></c><c r="B{n}" s="1"/>'
                    f'<c r="C{n}" t="inlineStr"><is><t>item {n % 97}</t></is></c></row>')
    for n in range(rows + 1, rows * 3 if styled else 0):
        body.append(f'<row r="{n}"><c r="A{n}" s="1"/><c r="B{n}" s="1"/></row>')
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS_MAIN}">'
           f'<sheetData>{"".join(body)}</sheetData><mergeCells count="1">'
           f'<mergeCell ref="D1:E1"/></mergeCells></worksheet>').encode()
    return _part_book(xml)


def test_a_pass_holds_a_window_of_the_sheet_never_the_sheet():
    data = _large(40_000)
    wb = Workbook.open(data)
    size = wb.package.part_size(SHEET)
    assert size > 8_000_000
    tracemalloc.start()
    try:
        count = sum(1 for _ in wb.cells("Stock", values_only=True))
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert count == 80_000
    assert peak < 2_000_000 < size / 4
    assert [m.ref for m in wb.merges("Stock")] == ["D1:E1"]
    assert wb.used_range("Stock").ref == "A1:C40000"
    rows = list(wb.read("Stock!A39990:C39991"))     # from a checkpoint, a note of the inflater
    assert [v.v for v in rows[1]][:1] == [39991]
    assert wb.package.part_entry(SHEET) is not None     # still compressed


def test_open_does_not_hold_the_sheet():
    data = _large(20_000)
    tracemalloc.start()
    try:
        wb = Workbook.open(data)
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert peak < len(data) + 2_000_000 < wb.package.part_size(SHEET)


def _damaged(data: bytes, name: str) -> bytes:
    """``data`` with ``name``'s CRC-32 wrong in both headers."""
    out = bytearray(data)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        info = archive.getinfo(name)
    out[info.header_offset + 14] ^= 0xFF
    at = data.find(b"PK\x01\x02")
    while at >= 0:
        size = int.from_bytes(data[at + 28:at + 30], "little")
        if data[at + 46:at + 46 + size] == name.encode():
            out[at + 16] ^= 0xFF
        at = data.find(b"PK\x01\x02", at + 4)
    return bytes(out)


def test_a_damaged_sheet_is_refused_at_open():
    """A sheet's data is checked as it is read: open reads every sheet through (for its
    protection), so a damaged one is refused there, as before, with the part's name."""
    data = _damaged(_large(200), SHEET)
    with pytest.raises(OpenError) as refused:
        Workbook.open(data)
    assert refused.value.code == "damaged" and SHEET in str(refused.value)


def test_damaged_shared_strings_are_refused_when_read():
    book = WorkbookBuilder()
    book.sheet("Stock").cell("A1", "crate").cell("A2", "pallet")
    data = _damaged(zip_parts(book.parts()), "xl/sharedStrings.xml")
    wb = Workbook.open(data)
    with pytest.raises(OpenError) as refused:
        list(wb.cells("Stock"))
    assert refused.value.code == "damaged"


@pytest.mark.parametrize("styled", [True, False])
@pytest.mark.parametrize("read", [1, 2, 5, 13, 64])
def test_the_row_index_of_a_streamed_sheet_is_the_held_one(monkeypatch, read, styled):
    """Row tags cut by a piece's end are found: the checkpoints are where they are in the
    part held whole, whether the rows are read from the window or the part is parsed."""
    data = _large(600, styled)
    with monkeypatch.context() as held:
        _held(held)
        whole = Workbook.open(data).scan("Stock").index
    monkeypatch.setattr(stream, "_READ", read)
    streamed = Workbook.open(data).scan("Stock").index
    assert whole is not None and len(whole.checkpoints) > 3
    assert streamed is not None
    assert (streamed.header, streamed.checkpoints) == (whole.header, whole.checkpoints)
    wb = Workbook.open(data)
    assert list(wb.read("Stock!A500:C502")) == list(Workbook.open(data).read("Stock!A500:C502"))
    wb.scan("Stock")
    with monkeypatch.context() as held:
        _held(held)
        assert list(wb.read("Stock!A500:C502")) == list(Workbook.open(data).read("Stock!A500:C502"))


def test_a_first_pass_inflates_the_sheet_once(monkeypatch):
    """What a pass needs of the whole part was found at open (with the protection), and
    the row index is counted on the way: a scan inflates the sheet once more, no more."""
    data = _large(600)
    wb = Workbook.open(data)
    opened = []
    original = ZipPart.open

    def counted(self, offset=0):
        opened.append((self.name, offset))
        return original(self, offset)

    monkeypatch.setattr(ZipPart, "open", counted)
    found = wb.scan("Stock")
    assert found.index is not None and len(found.index.checkpoints) > 10
    assert [name for name, _ in opened] == [SHEET]
    assert wb.used_range("Stock").ref == "A1:C600" and len(opened) == 1
