"""R73: a sheet formatted to its last row, read the short way.

A row that holds only styles (each cell ``<c r="B9" s="3"/>``, or none) is read from the
part's bytes instead of being parsed; a cell with nothing in it is placed but not typed
by a pass that counts values only (a scan, ``cells(values_only=True)``); a pass over every
cell scans the sheet on the way; merges are read without a pass over the cells.  Each is
checked here against the long way round -- every row parsed, every cell typed, as before
-- on sheets generated to trip it: rows and cells without ``r``, ``$`` and lower-case
references, cells placed in other rows, rows out of order, hidden rows, a prefix on the
main namespace, attributes with prefixes, whitespace between and inside tags, a style-only
``t="s"`` cell, and the parts no row may be skimmed in (a comment, CDATA, a processing
instruction, a namespace bound among the rows).  The part is fed through windows and
pieces small enough to split rows anywhere.

Every workbook is made up here, with tests/build.py's parts.
"""

import random
from collections import deque

import pytest

from build import NS_MAIN, NS_R, WorkbookBuilder, zip_parts

from xlsx_agent import Workbook
from xlsx_agent.package import OpenError
from xlsx_agent.read import stream
from xlsx_agent.read.stream import CHECKPOINT_ROWS, SheetReader

NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
WORDS = ["crate", "pallet", "12", "2026-01-05", "tin", "jar"]


def _book(sheet_xml: str) -> bytes:
    """A workbook of one sheet, ``Stock``, whose part is ``sheet_xml``; six shared strings
    (two of them look like a number and a date) and four cell styles (one a date)."""
    book = WorkbookBuilder()
    book.sheet("Stock").cell("A1", "crate")
    book.style(num_fmt="yyyy-mm-dd")
    book.style(num_fmt="0.00")
    book.style(num_fmt="@")
    parts = book.parts()
    items = "".join(f"<si><t>{w}</t></si>" for w in WORDS)
    parts["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{NS_MAIN}">'
        f"{items}</sst>").encode()
    parts["xl/worksheets/sheet1.xml"] = sheet_xml.encode()
    return zip_parts(parts)


class Sheet:
    """A sheet part made at random (``seed``): see the module's notes."""

    def __init__(self, seed: int, *, rows: int = 400, odd: str | None = None) -> None:
        self.rng = random.Random(seed)
        rng = self.rng
        self.p = rng.choice(["", "", "x:"])
        self.space = rng.random() < 0.3   # newlines and spaces between and inside tags
        self.quote = rng.choice(['"', '"', "'"])
        self.rows = rows
        self.odd = odd

    def attr(self, name: str, value: object) -> str:
        q = self.quote
        if self.space and self.rng.random() < 0.3:
            return f" {name} = {q}{value}{q}"
        return f" {name}={q}{value}{q}"

    def gap(self) -> str:
        return self.rng.choice(["\n", "\n  ", " ", ""]) if self.space else ""

    def cell(self, row: int, column: int, *, blank: bool) -> str:
        rng, p = self.rng, self.p
        letters = "ABCDEFGHIJ"[column - 1]
        roll = rng.random()
        if roll < 0.08:
            ref = ""                                   # placed after the cell before it
        elif roll < 0.12:
            ref = self.attr("r", f"${letters}${row}")
        elif roll < 0.15:
            ref = self.attr("r", f"{letters.lower()}{row}")
        elif roll < 0.17:
            ref = self.attr("r", f"{letters}{max(1, row + rng.choice([-3, 1, 7]))}")
        else:
            ref = self.attr("r", f"{letters}{row}")
        style = self.attr("s", rng.choice([0, 1, 2, 3, 4])) if rng.random() < 0.8 else ""
        if blank:
            kind = rng.random()
            if kind < 0.05:
                return f"<{p}c{ref}{style}{self.attr('t', 's')}/>"   # a string without index
            if kind < 0.08:
                return f"<{p}c{ref}{style}></{p}c>"
            if kind < 0.10:
                return f"<{p}c{style}{ref}/>"                         # s before r
            return f"<{p}c{ref}{style}{self.gap() if self.space else ''}/>"
        kind = rng.random()
        if kind < 0.3:
            return f"<{p}c{ref}{style}><{p}v>{rng.randint(-5, 50000)}</{p}v></{p}c>"
        if kind < 0.45:
            return (f"<{p}c{ref}{style}{self.attr('t', 's')}><{p}v>{rng.randint(0, 6)}</{p}v>"
                    f"</{p}c>")
        if kind < 0.55:
            return f"<{p}c{ref}{style}><{p}v>{rng.random() * 100:.4f}</{p}v></{p}c>"
        if kind < 0.62:
            return (f"<{p}c{ref}{self.attr('t', 'inlineStr')}><{p}is><{p}t>inline {row}</{p}t>"
                    f"</{p}is></{p}c>")
        if kind < 0.7:
            return f"<{p}c{ref}><{p}f>A{row}*2</{p}f><{p}v>{row * 2}</{p}v></{p}c>"
        if kind < 0.75:
            return f"<{p}c{ref}><{p}f>TODAY()</{p}f></{p}c>"
        if kind < 0.8:
            return f"<{p}c{ref}{self.attr('t', 'b')}><{p}v>1</{p}v></{p}c>"
        if kind < 0.85:
            return f"<{p}c{ref}{self.attr('t', 'e')}><{p}v>#N/A</{p}v></{p}c>"
        if kind < 0.9:
            return f"<{p}c{ref}{style}><{p}v>abc</{p}v></{p}c>"     # bad value
        return f"<{p}c{ref}{self.attr('t', 'str')}><{p}v>text {row}</{p}v></{p}c>"

    def repeated(self, number: int) -> str:
        """A row as Excel writes a formatted one, of a few shapes that repeat (so rows are
        read by template), some made to look alike but not be: the number as a style, a
        cell of another row, ``#`` in a value, a column off the grid, columns backwards."""
        rng, p, n = self.rng, self.p, number
        shape = rng.randrange(10)
        cells = "".join(f'<{p}c r="{c}{n}" s="1"/>' for c in "ABCD")
        if shape == 0:
            return f'<{p}row r="{n}" spans="1:4" s="2" customFormat="1"/>'
        if shape == 1:
            return f'<{p}row r="{n}" spans="1:4">{cells}</{p}row>'
        if shape == 2:
            return f'<{p}row r="{n}" spans="1:4" hidden="1" x14ac:dyDescent="0.25">{cells}</{p}row>'
        if shape == 3:
            return f'<{p}row r="{n}"><{p}c r="B{n}" s="{n % 7}"/><{p}c r="C{n}" s="5"/></{p}row>'
        if shape == 4:
            return f'<{p}row r="{n}"><{p}c r="A{n}"/><{p}c r="B{n + 1}" s="1"/></{p}row>'
        if shape == 5:
            return f'<{p}row r="{n}" spans="#1"><{p}c r="A{n}" s="1"/></{p}row>'
        if shape == 6:
            return f'<{p}row r="{n}"><{p}c r="XFE{n}" s="1"/><{p}c r="A{n}" s="1"/></{p}row>'
        if shape == 7:
            return f'<{p}row r="{n}"><{p}c r="C{n}" s="1"/><{p}c r="B{n}" s="1"/></{p}row>'
        if shape == 8:   # a cell of another row whose number ends in this one's
            return f'<{p}row r="{n}"><{p}c r="A{n}" s="1"/><{p}c r="B1{n}" s="1"/></{p}row>'
        return f'<{p}row r="{n}" spans="1:4" s="{n % 3}" customFormat="1">{cells}</{p}row>'

    def row(self, number: int) -> str:
        rng, p = self.rng, self.p
        if rng.random() < 0.35:
            return self.repeated(number)
        attrs = ""
        roll = rng.random()
        if roll > 0.05:
            attrs += self.attr("r", number)
        if rng.random() < 0.5:
            attrs += self.attr("spans", "1:8")
        if rng.random() < 0.1:
            attrs += self.attr("hidden", rng.choice(["1", "true", "0"]))
        if rng.random() < 0.3:
            attrs += self.attr("s", 2) + self.attr("customFormat", 1)
        if rng.random() < 0.2:
            attrs += self.attr("x14ac:dyDescent", "0.25")
        shape = rng.random()
        if shape < 0.25:
            return f"<{p}row{attrs}/>"
        if shape < 0.3:
            return f"<{p}row{attrs}>{self.gap()}</{p}row>"
        blank = shape < 0.75
        width = rng.randint(1, 8)
        cells = self.gap().join(self.cell(number, c, blank=blank and rng.random() < 0.97)
                                for c in range(1, width + 1))
        return f"<{p}row{attrs}>{self.gap()}{cells}{self.gap()}</{p}row>"

    def xml(self) -> str:
        rng, p = self.rng, self.p
        numbers = list(range(1, self.rows + 1))
        if rng.random() < 0.2:          # a few rows out of order
            for _ in range(3):
                i, j = rng.randrange(len(numbers)), rng.randrange(len(numbers))
                numbers[i], numbers[j] = numbers[j], numbers[i]
        body = [self.row(n) for n in numbers]
        if self.odd is not None:
            at = rng.randrange(1, len(body) - 1)
            body[at] = {"comment": "<!-- <row r='9'/> -->" + body[at],
                        "pi": "<?note <row r='9'/> ?>" + body[at],
                        "xmlns": f'<{p}row r="{numbers[at]}" xmlns:y="urn:example:y"/>',
                        "duplicate": f'<{p}row r="{numbers[at]}" r="{numbers[at]}"/>',
                        "unbound": f'<{p}row r="{numbers[at]}" y:s="1"/>',
                        "cdata": f'<{p}row r="{numbers[at]}"><{p}c r="A{numbers[at]}" '
                                 f't="str"><{p}v><![CDATA[<row/>]]></{p}v></{p}c></{p}row>',
                        "merge": f'<{p}mergeCell ref="C3:D9"/>' + body[at]}[self.odd]
        ns = (f'xmlns:x="{NS_MAIN}" xmlns="urn:example:none"' if p else f'xmlns="{NS_MAIN}"')
        merges = "".join(f'<{p}mergeCell ref="{m}"/>' for m in ["A1:C1", "B5:B9", "E2:F3"]
                         if rng.random() < 0.8)
        cols = f'<{p}cols><{p}col min="3" max="3" hidden="1"/></{p}cols>' \
            if rng.random() < 0.3 else ""
        tail = (f'<{p}mergeCells>{merges}</{p}mergeCells>' if merges else "")
        tail += (f'<{p}dataValidations count="1"><{p}dataValidation type="list" sqref="A2:A9">'
                 f'<{p}formula1>"a,b"</{p}formula1></{p}dataValidation></{p}dataValidations>')
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<{p}worksheet {ns} xmlns:r="{NS_R}" xmlns:x14ac="{NS_X14AC}">'
                f'<{p}dimension ref="A1:H9"/>{cols}<{p}sheetData>'
                f"{self.gap().join(body)}</{p}sheetData>{tail}</{p}worksheet>")


def _long_way(monkeypatch, data: bytes, ranges: list[str]):
    """What the reads gave before: every row parsed, every cell typed."""
    monkeypatch.setattr(stream, "SKIM", False)
    original = SheetReader.values

    def every(self, *, blanks=True):
        return original(self, blanks=True)

    monkeypatch.setattr(SheetReader, "values", every)
    try:
        wb = Workbook.open(data)
        found = wb.scan("Stock")
        cells = list(wb.cells("Stock", formats=True))
        reads = [list(wb.read(f"Stock!{r}", formats=True)) for r in ranges]
        texts = list(wb.find_text("crate"))
        return found, cells, reads, wb.merges("Stock"), texts
    finally:
        monkeypatch.setattr(SheetReader, "values", original)
        monkeypatch.setattr(stream, "SKIM", True)


def _short_way(data: bytes, ranges: list[str], *, cells_first: bool):
    wb = Workbook.open(data)
    merges = wb.merges("Stock")       # before any pass: read from the tail
    if cells_first:
        cells = list(wb.cells("Stock", formats=True))
        found = wb.scan("Stock")      # kept from the pass above
    else:
        found = wb.scan("Stock")
        cells = list(wb.cells("Stock", formats=True))
    reads = [list(wb.read(f"Stock!{r}", formats=True)) for r in ranges]
    texts = list(wb.find_text("crate"))
    return found, cells, reads, merges, texts, list(Workbook.open(data).cells(
        "Stock", formats=True, values_only=True))


def _same_scan(one, two) -> None:
    assert one == two     # every field but the index
    if one.index is None or two.index is None:
        assert one.index is two.index is None
    else:
        assert (one.index.header, one.index.checkpoints, one.index.hidden,
                one.index.shared_formulas) == (two.index.header, two.index.checkpoints,
                                               two.index.hidden, two.index.shared_formulas)


RANGES = ["A1:H5", "B100:E140", "A250:H420", "C3:C3", "A1:J400"]


@pytest.mark.parametrize("window,piece", [(1 << 20, 1 << 15), (300, 64), (97, 13)])
@pytest.mark.parametrize("seed", range(12))
def test_the_short_way_gives_what_the_long_way_gives(monkeypatch, seed, window, piece):
    monkeypatch.setattr(stream, "_WINDOW", window)
    monkeypatch.setattr(stream, "_PIECE", piece)
    data = _book(Sheet(seed).xml())
    long_scan, long_cells, long_reads, long_merges, long_texts = _long_way(monkeypatch, data,
                                                                          RANGES)
    for cells_first in (False, True):
        found, cells, reads, merges, texts, values = _short_way(data, RANGES,
                                                                cells_first=cells_first)
        _same_scan(found, long_scan)
        assert cells == long_cells
        assert reads == long_reads
        assert merges == long_merges
        assert texts == long_texts
        assert values == [v for v in long_cells if v.t != "blank" or v.formula]
    assert len(long_cells) > 100 and long_scan.cells > 10


@pytest.mark.parametrize("seed", range(4))
def test_rows_are_skimmed(monkeypatch, seed):
    """The sheets above do have rows read from the bytes, at every prefix."""
    xml = Sheet(seed).xml()
    wb = Workbook.open(_book(xml))
    reader = wb._reader(wb._cell_sheet("Stock"))
    plan = reader.plan()
    assert plan is not None
    kinds = [type(written).__name__ for node, _, written in reader._events(None)
             if node is None]
    assert kinds.count("_Template") > 30 and len(kinds) - kinds.count("_Template") > 30


@pytest.mark.parametrize("odd", ["comment", "pi", "xmlns", "cdata", "merge"])
def test_a_part_that_may_hide_a_tag_is_parsed_whole(monkeypatch, odd):
    data = _book(Sheet(3, odd=odd).xml())
    wb = Workbook.open(data)
    if odd != "merge":
        assert wb._reader(wb._cell_sheet("Stock")).plan() is None
    long_scan, long_cells, long_reads, long_merges, _ = _long_way(monkeypatch, data, RANGES)
    found, cells, reads, merges, _, _ = _short_way(data, RANGES, cells_first=True)
    _same_scan(found, long_scan)
    assert (cells, reads, merges) == (long_cells, long_reads, long_merges)
    if odd == "merge":   # a merged area among the rows: found by the pass
        assert merges[0].ref == "C3:D9"


@pytest.mark.parametrize("odd", ["duplicate", "unbound"])
def test_a_row_the_parser_refuses_is_refused_still(monkeypatch, odd):
    data = _book(Sheet(5, odd=odd).xml())
    wb = Workbook.open(data)
    assert wb._reader(wb._cell_sheet("Stock")).plan() is not None
    for call in (lambda: wb.scan("Stock"), lambda: list(wb.cells("Stock")),
                 lambda: list(wb.cells("Stock", values_only=True))):
        with pytest.raises(OpenError) as caught:
            call()
        assert caught.value.code == "damaged"


# -- the shared pass, values only, merges within a range ----------------------------------------


def formatted_to_the_bottom(rows: int = 300, bottom: int = 3000, *, empty: bool = False) -> bytes:
    """``rows`` rows of four values, then rows to ``bottom`` that hold only styles (each
    four style-only cells, or ``empty``: none, the row styled), and merged areas after."""
    body = []
    for r in range(1, bottom + 1):
        if r <= rows:
            body.append(f'<row r="{r}" spans="1:4"><c r="A{r}" s="1" t="s"><v>{r % 6}</v></c>'
                        f'<c r="B{r}" s="1"><v>{r}</v></c><c r="C{r}" s="2"><v>{45000 + r}</v>'
                        f'</c><c r="D{r}" s="1"><f>B{r}*2</f><v>{2 * r}</v></c></row>')
        elif empty:
            body.append(f'<row r="{r}" spans="1:4" s="1" customFormat="1"/>')
        else:
            body.append(f'<row r="{r}" spans="1:4">'
                        + "".join(f'<c r="{c}{r}" s="1"/>' for c in "ABCD") + "</row>")
    merges = '<mergeCells count="3"><mergeCell ref="A1:D1"/><mergeCell ref="A5:A7"/>' \
             f'<mergeCell ref="B{rows + 50}:C{rows + 50}"/></mergeCells>'
    return _book(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><dimension ref="A1:D{bottom}"/>'
                 f'<sheetData>{"".join(body)}</sheetData>{merges}</worksheet>')


def _count_values(monkeypatch) -> list[int]:
    calls = [0]
    original = SheetReader.value

    def counted(self, *args):
        calls[0] += 1
        return original(self, *args)

    monkeypatch.setattr(SheetReader, "value", counted)
    return calls


@pytest.mark.parametrize("empty", [False, True])
def test_values_only_types_no_formatted_cell(monkeypatch, empty):
    wb = Workbook.open(formatted_to_the_bottom(empty=empty))
    calls = _count_values(monkeypatch)
    values = list(wb.cells("Stock", values_only=True))
    assert len(values) == 300 * 4 and calls[0] == 300 * 4
    assert all(v.t != "blank" for v in values)
    assert values[-1].a == "Stock!D300" and values[-1].formula


def test_one_pass_over_every_cell_scans_the_sheet(monkeypatch):
    wb = Workbook.open(formatted_to_the_bottom())
    passes = []
    original = stream.SheetReader._walk

    def counted(self, *args, **kwargs):
        passes.append(1)
        return original(self, *args, **kwargs)

    monkeypatch.setattr(stream.SheetReader, "_walk", counted)
    every = list(wb.cells("Stock"))
    assert len(every) == 300 * 4 + 2700 * 4 and len(passes) == 1
    assert wb.used_range("Stock").ref == "A1:D300"
    assert [m.ref for m in wb.merges("Stock")] == ["A1:D1", "A5:A7", "B350:C350"]
    assert wb.scan("Stock").cells == 1200
    assert len(passes) == 1   # the used range and merges came from the pass of cells()


def test_a_pass_stopped_early_keeps_no_scan():
    wb = Workbook.open(formatted_to_the_bottom())
    deque(zip(range(10), wb.cells("Stock")), maxlen=0)
    assert "xl/worksheets/sheet1.xml" not in wb._scans
    assert wb.used_range("Stock").ref == "A1:D300"


def test_a_pass_after_an_edit_keeps_no_scan_of_before():
    wb = Workbook.open(formatted_to_the_bottom())
    values = wb.cells("Stock")
    next(values)
    wb.set_cells({"Stock!F2": 5})
    deque(values, maxlen=0)   # the pass reads the bytes it started with
    assert wb.used_range("Stock").ref == "A1:F300"


def test_merges_are_read_without_a_pass_over_the_cells(monkeypatch):
    wb = Workbook.open(formatted_to_the_bottom())
    calls = _count_values(monkeypatch)
    walked = []
    monkeypatch.setattr(stream.SheetReader, "_walk",
                        lambda self, *a, **k: walked.append(1) or iter(()))
    assert [m.ref for m in wb.merges("Stock")] == ["A1:D1", "A5:A7", "B350:C350"]
    assert calls[0] == 0 and walked == [] and "xl/worksheets/sheet1.xml" not in wb._scans


def test_merges_within_a_range():
    wb = Workbook.open(formatted_to_the_bottom())
    assert [m.ref for m in wb.merges("Stock", "A6:B10")] == ["A5:A7"]
    assert [m.ref for m in wb.merges("Stock", "Stock!A1:A1")] == ["A1:D1"]
    assert [m.ref for m in wb.merges("Stock", "C340:C360")] == ["B350:C350"]
    assert wb.merges("Stock", "E1:H9") == []
    assert [m.ref for m in wb.merges("Stock", "A:A")] == ["A1:D1", "A5:A7"]
    every = [m.ref for m in wb.merges("Stock")]
    assert [m.ref for m in wb.merges("Stock", "")] == every   # strict mode's placeholder
    assert [m.ref for m in wb.merges("Stock", "Stock!")] == every
    for wrong in ("not a range", "Stock!A1:A1:B2", "t:Stock!A1"):
        with pytest.raises(ValueError) as caught:
            wb.merges("Stock", wrong)
        assert "A1:S40" in str(caught.value)
    with pytest.raises(KeyError) as missing:   # a sheet there is not: the sheets there are
        wb.merges("Stock", "Other!A1:B2")
    assert "Stock" in str(missing.value)


def test_merges_after_an_edit_are_the_edit_s():
    wb = Workbook.open(formatted_to_the_bottom())
    wb.merge("Stock!F1:G1")
    assert "F1:G1" in [m.ref for m in wb.merges("Stock")]
    wb.undo()
    assert "F1:G1" not in [m.ref for m in wb.merges("Stock")]


def test_a_read_from_a_checkpoint_among_skimmed_rows():
    """A read that starts at a checkpoint the skimmed rows gave."""
    wb = Workbook.open(formatted_to_the_bottom(rows=CHECKPOINT_ROWS * 3, bottom=2000))
    index = wb.scan("Stock").index
    assert index is not None and index.before(1500) is not None
    rows = list(wb.read("Stock!A380:D390"))
    assert rows[0][0].v == WORDS[380 % 6] and rows[0][1].a == "Stock!B380"
    assert all(v.t == "blank" for v in rows[-1])
    found = wb.scan("Stock")
    index, found.index = found.index, None
    try:
        assert list(wb.read("Stock!A380:D390")) == rows
        assert list(wb.read("Stock!A1500:D1502", formats=True)) == \
            list(Workbook.open(formatted_to_the_bottom(rows=CHECKPOINT_ROWS * 3, bottom=2000))
                 .read("Stock!A1500:D1502", formats=True))
    finally:
        found.index = index


@pytest.mark.parametrize("cell", ['<c r="A{n}" s="{n}"/>',          # the number as a style
                                  '<c r="A{n}" s="1"/><c r="B1{n}" s="2"/>',  # another row's
                                  '<c r="A{n}" s="{n}"/><c r="B{n}" s="1"/>'])
def test_rows_alike_but_for_more_than_their_number_are_read_each(monkeypatch, cell):
    """Rows whose bytes differ only where the number stands, but whose number stands
    for something else there too: never read by one template."""
    rows = "".join(f'<row r="{n}">{cell.format(n=n)}</row>' for n in range(1, 4))
    rows += "".join(f'<row r="{n}"><c r="C{n}"><v>{n}</v></c></row>' for n in range(4, 9))
    data = _book(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<worksheet xmlns="{NS_MAIN}"><sheetData>{rows}</sheetData></worksheet>')
    long_scan, long_cells, long_reads, _, _ = _long_way(monkeypatch, data, ["A1:C8"])
    found, cells, reads, _, _, _ = _short_way(data, ["A1:C8"], cells_first=True)
    _same_scan(found, long_scan)
    assert (cells, reads) == (long_cells, long_reads)
    assert len({v.format for v in long_cells if v.column == 1}) > 1 or "B1" in cell
