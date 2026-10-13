"""P1, performance: each faster path gives what the path it replaced gave.

The timings themselves are tools/bench.py's; here every shortcut is checked against the
long way round on sheets made to trip it: a read started at a checkpoint against one from
the top.

Every workbook is built with tests/build.py, the sheets it cannot write by hand.
"""

import pytest

from build import NS_MAIN, NS_R, WorkbookBuilder, zip_parts

from xlsx_agent import Workbook
from xlsx_agent.read.stream import CHECKPOINT_ROWS, RowIndex


def _with_sheet(body: str, *, before: str = "", after: str = "", prefix: str = "",
                strings: int = 0) -> bytes:
    """A workbook of one sheet, ``Data``, written by hand (``prefix``: the main namespace
    bound to it, as ``x``), with ``strings`` shared strings ``s0``, ``s1``..."""
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    sheet.cell("A1", "s0")  # so that the builder writes a shared strings part
    book.style(num_fmt="yyyy-mm-dd")
    parts = book.parts()
    items = "".join(f"<si><t>s{k}</t></si>" for k in range(max(strings, 1)))
    parts["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{NS_MAIN}">'
        f"{items}</sst>").encode()
    ns = f'xmlns:{prefix}="{NS_MAIN}"' if prefix else f'xmlns="{NS_MAIN}"'
    p = f"{prefix}:" if prefix else ""
    parts["xl/worksheets/sheet1.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<{p}worksheet {ns} xmlns:r="{NS_R}">{before}<{p}sheetData>{body}</{p}sheetData>'
        f"{after}</{p}worksheet>").encode()
    return zip_parts(parts)


def _messy_rows(rows: int, prefix: str = "") -> str:
    """Rows to trip a resumed read: hidden rows (some with no cells), rows and cells
    without ``r`` (a row at a checkpoint among them), ``$`` in a reference, shared
    formulas whose first cell is far above, text, dates and errors; and an empty hidden
    row written early whose cell comes late, in a row of its own."""
    p = f"{prefix}:" if prefix else ""
    out = []
    late = rows - 2
    for r in range(1, rows + 1):
        if r == 300:
            out.append(f'<{p}row r="{late}" hidden="1"/>')
        if r == late:
            continue
        if r == late + 1:
            out.append(f'<{p}row r="{r}"><{p}c r="A{late}"><{p}v>1</{p}v></{p}c></{p}row>')
            continue
        attrs = "" if r % 97 == 5 or r % CHECKPOINT_ROWS == 1 else f' r="{r}"'
        if r % 41 == 0:
            out.append(f'<{p}row r="{r}" hidden="1"/>')
            continue
        if r % 13 == 0:
            attrs += ' hidden="1"'
        cells = [f'<{p}c r="A{r}"><{p}v>{r}</{p}v></{p}c>',
                 f'<{p}c r="$B${r}" t="s"><{p}v>{r % 5}</{p}v></{p}c>',
                 f'<{p}c><{p}v>{r * 0.5}</{p}v></{p}c>',  # C, by position
                 f'<{p}c r="D{r}" s="1"><{p}v>{45000 + r}</{p}v></{p}c>']
        if r % 29 == 0:
            cells.append(f'<{p}c r="E{r}" t="e"><{p}v>#N/A</{p}v></{p}c>')
        if r == 3:
            cells.append(f'<{p}c r="F3"><{p}f t="shared" ref="F3:F{rows}" si="0">A3*2</{p}f>'
                         f'<{p}v>6</{p}v></{p}c>')
        elif r > 3:
            cells.append(f'<{p}c r="F{r}"><{p}f t="shared" si="0"/><{p}v>{r * 2}</{p}v></{p}c>')
        out.append(f"<{p}row{attrs}>{''.join(cells)}</{p}row>")
    return "".join(out)


def _from_top(wb: Workbook, address: str) -> list:
    """``address`` read as it was before the index: every pass from the top."""
    found = wb.scan("Data")
    index, found.index = found.index, None
    try:
        return list(wb.read(address))
    finally:
        found.index = index


@pytest.mark.parametrize("prefix", ["", "x"])
def test_a_read_from_a_checkpoint_gives_what_a_read_from_the_top_gives(prefix):
    rows = CHECKPOINT_ROWS * 12 + 7
    wb = Workbook.open(_with_sheet(_messy_rows(rows, prefix), prefix=prefix, strings=5,
                                   before=f'<{prefix + ":" if prefix else ""}cols>'
                                          f'<{prefix + ":" if prefix else ""}col min="3" '
                                          f'max="3" hidden="1"/>'
                                          f'</{prefix + ":" if prefix else ""}cols>'))
    index = wb.scan("Data").index
    assert index is not None and len(index.checkpoints) == 13
    for top in [1, 2, CHECKPOINT_ROWS, CHECKPOINT_ROWS + 1, 500, 777, 1000, rows - 3, rows]:
        assert index.before(top) is None or index.before(top).last_row < top
        for address in [f"Data!A{top}:F{min(top + 60, rows + 5)}", f"Data!C{top}:E{top}",
                        f"Data!A{rows - 2}:F{rows}"]:
            fast = list(wb.read(address))
            assert fast == _from_top(wb, address), address
    assert index.before(1000) is not None   # the reads above did start part-way
    late = list(wb.read(f"Data!A{rows - 2}:F{rows}"))
    assert late[0][0].hidden and late[0][0].v == 1   # its row was hidden before the checkpoint
    assert [v.hidden for v in late[2]] == [False, False, True, False, False, False]
    assert late[2][5].shared_from is None   # as from the top: the group's first cell unread


@pytest.mark.parametrize("extra", ["<!-- <row r='9'/> -->", "<?pi <row r='9'/> ?>",
                                   "<![CDATA[<row r='9'/>]]>"])
def test_a_sheet_whose_bytes_may_hide_a_tag_gets_no_index(extra):
    rows = _messy_rows(CHECKPOINT_ROWS * 3)
    if "CDATA" in extra:   # where character data may stand: in a cell's value
        rows = rows.replace("<v>7</v>", f"<v>7{extra}</v>", 1)
    else:
        rows = rows.replace("</row>", "</row>" + extra, 1)
    wb = Workbook.open(_with_sheet(rows, strings=5))
    assert wb.scan("Data").index is None
    address = f"Data!A{CHECKPOINT_ROWS * 2}:F{CHECKPOINT_ROWS * 2 + 3}"
    assert len(list(wb.read(address))) == 4


def test_a_sheet_out_of_order_or_short_gets_no_index():
    rows = "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>'
                   for r in [*range(1, 300), 2])
    assert Workbook.open(_with_sheet(rows)).scan("Data").index is None
    short = "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>' for r in range(1, 50))
    assert Workbook.open(_with_sheet(short)).scan("Data").index is None


def test_an_index_holds_a_checkpoint_per_block_of_rows_not_the_rows():
    wb = Workbook.open(_with_sheet(_messy_rows(CHECKPOINT_ROWS * 40)))
    index = wb.scan("Data").index
    assert isinstance(index, RowIndex) and len(index.checkpoints) == 40
    assert len(index.header) < 1000
