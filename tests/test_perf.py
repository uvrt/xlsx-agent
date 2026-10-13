"""P1, performance: each faster path gives what the path it replaced gave.

The timings themselves are tools/bench.py's; here every shortcut is checked against the
long way round on sheets made to trip it: a read started at a checkpoint against one from
the top.

Every workbook is built with tests/build.py, the sheets it cannot write by hand.
"""

import datetime as dt

import pytest

from build import NS_MAIN, NS_R, Formula, WorkbookBuilder, zip_parts
from ooxml_edit.tools import Toolbox

from xlsx_agent import Workbook, find, query, workbook
from xlsx_agent.read import stream
from xlsx_agent.read.regions import Regions
from xlsx_agent.read.stream import CHECKPOINT_ROWS, RowIndex
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS


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
                                   "<![CDATA[<row r='9'/>]]>", "before"])
def test_a_sheet_whose_bytes_may_hide_a_tag_gets_no_index(extra):
    rows = _messy_rows(CHECKPOINT_ROWS * 3)
    before = ""
    if extra == "before":   # an element of another namespace named alike, before the cells
        before = '<o:sheetData xmlns:o="urn:example:other"/>'
    elif "CDATA" in extra:   # where character data may stand: in a cell's value
        rows = rows.replace("<v>7</v>", f"<v>7{extra}</v>", 1)
    else:
        rows = rows.replace("</row>", "</row>" + extra, 1)
    wb = Workbook.open(_with_sheet(rows, strings=5, before=before))
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


# -- every tool, the fast way and the long way round -------------------------------------------

ROWS = CHECKPOINT_ROWS * 20 + 11
CATEGORIES = ["Hardware", "Tools", "Garden", "Paint"]


def shop_floor(*, shared: bool, changed: bool = False) -> bytes:
    """A table of :data:`ROWS` rows over several checkpoints (codes and categories as
    shared strings, amounts, dates, formulas, one with text), hidden rows and a hidden
    column, a block below it and a table below that; ``shared``: two columns are one
    shared formula group whose first cell is at the top (so rows far apart are read in
    one pass), and a block stands to the table's right part-way down (so no row of it is
    covered across the sheet); ``changed``: some amounts and a code differ."""
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    columns = ["Code", "Category", "Amount", "Day", "Twice", "Again", "Label"]
    for r in range(2, ROWS + 2):
        i = r - 2
        amount = round(i * 0.37 % 1000, 2) + (1 if changed and i % 211 == 0 else 0)
        if not shared:
            twice = again = Formula(f"C{r}*2", cached=amount * 2)
        elif r == 2:
            twice = Formula("C2*2", cached=amount * 2, kind="shared", ref=f"E2:F{ROWS + 1}",
                            si=0)
            again = Formula("", cached=amount * 2, kind="shared", si=0)
        else:
            twice = again = Formula("", cached=amount * 2, kind="shared", si=0)
        code = f"SKU-{i:06d}" if not (changed and i == 77) else "SKU-NEW"
        sheet.rows(f"A{r}", [[code, CATEGORIES[i % 4], amount, 45000 + i % 900, twice, again,
                              Formula(f'"L"&A{r}', cached=f"L{code}")]])
        if i % 37 == 5:
            sheet.hide_row(r)
    sheet.hide_column("C")
    sheet.table("Big", f"A1:G{ROWS + 1}", columns)
    if shared:
        sheet.rows("J900", [["Note", "Qty"], ["Hardware here", 1], ["more", 2]])
    sheet.rows(f"A{ROWS + 5}", [["Item", "Price"], ["I-1", 1.5], ["Hardware", 2.5]])
    sheet.rows(f"A{ROWS + 11}", [[f"T-{k}", k, "Garden"] for k in range(10)])
    sheet.table("Tail", f"A{ROWS + 10}:C{ROWS + 20}", ["Tag", "N", "Kind"])
    return book.to_bytes()


def _calls(ident: str, other: str | None):
    """The tool calls compared: ``(name, arguments)``, ``"next"`` for the page after the
    last call's."""
    late = f"Data!A{ROWS - 40}:G{ROWS + 6}"
    out = [("describe", {}),
           ("xlsx_read", {"target": "t:Big"}), "next", "next",
           ("xlsx_read", {"target": late, "formulas": True, "merges": True}),
           ("xlsx_read", {"target": "Data!E1500:F1510", "formulas": True}),
           ("xlsx_read", {"target": "t:Big/SKU-002000", "key_column": "Code"}),
           ("find_text", {"text": "SKU-0001"}), "next", "next", "next", "next",
           ("find_text", {"text": "Hardware"}), "next",
           ("find_text", {"text": "Garden", "range": f"Data!A1500:G{ROWS}"}), "next",
           ("find_text", {"text": r"L?SKU-00(1|2)\d5", "regex": True}), "next",
           ("xlsx_query", {"table": "t:Big", "where": [{"column": "Category", "op": "eq",
                                                         "value": "Tools"}],
                           "sort": [{"column": "Amount", "desc": True}]}), "next", "next",
           ("xlsx_query", {"table": "t:Big", "columns": ["Code", "Twice", "Label"],
                           "sort": [{"column": "Day"}]}), "next",
           # Again's group's first cell is in Twice: read on the way, or not
           ("xlsx_query", {"table": "t:Big", "columns": ["Again"], "sort": [{"column": "Day"}]}),
           ("xlsx_query", {"table": "t:Big", "columns": ["Again"],
                           "where": [{"column": "Category", "op": "eq", "value": "Tools"}]}),
           ("xlsx_query", {"table": "t:Big", "columns": ["Twice"],
                           "where": [{"column": "Amount", "op": "gte", "value": "900"}]}),
           ("xlsx_profile", {"target": "t:Big", "columns": ["Day", "Label"]}),
           ("xlsx_derive", {"from": "t:Big", "id": "g", "group_by": ["Category"],
                            "measures": [{"fn": "sum", "column": "Amount", "as": "total"},
                                         {"fn": "median", "column": "Twice", "as": "mid"}]}),
           ("xlsx_validate_data", {"schema": {"tables": [{"table": "t:Big", "key": ["Code"],
                                                          "columns": [
               {"name": "Amount", "type": "number", "min": "1"},
               {"name": "Label", "type": "text", "pattern": "LSKU-0.*"}]}]}})]
    if other is not None:
        out.append(("xlsx_compare", {"a": ident, "b": other, "tables": ["t:Big"],
                                         "key": "Code"}))
    return out


class _Everything:
    def __contains__(self, item) -> bool:
        return True


def _long_way_round(monkeypatch) -> None:
    """Every shortcut off: no row index (so no read starts part-way, and find_regions and
    query read every row), every cell typed, find_text typing every cell, the key columns
    found in a pass of their own."""
    monkeypatch.setattr(stream.RowIndex, "build", classmethod(lambda cls, data, reader: None))
    read_area = stream.read_area
    monkeypatch.setattr(workbook, "read_area", lambda reader, area, ordered, index=None,
                        columns=None: read_area(reader, area, ordered, index, None))
    monkeypatch.setattr(Regions, "key_finder", lambda self, region: None)
    monkeypatch.setattr(find, "TEXT_KINDS", _Everything())
    monkeypatch.setattr(stream.SheetReader, "shared_index", staticmethod(lambda node: None))


def _run(data: bytes, other: bytes, *, slow: bool) -> list:
    out = []
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session(clock=lambda: dt.datetime(2026, 10, 13, tzinfo=dt.timezone.utc))
        for name, blob in (("a.xlsx", data), ("b.xlsx", other)):
            box.dispatch(session, "open_document", {"blob": session.add_blob(blob, name=name)})
        last = None
        for call in _calls("d1", "d2"):
            if call == "next":
                if last is None or last[1].next_cursor is None:
                    continue
                name, arguments = last[0], {**last[2], "cursor": last[1].next_cursor}
            else:
                name, arguments = call
            if "doc" not in arguments and name not in ("xlsx_compare",):
                arguments = {"doc": "d1", **arguments}
            if slow:  # nothing kept from an earlier page
                session.entry("d1").document._finds.clear()
            result = box.dispatch(session, name, arguments)
            assert result.ok, (name, result.error)
            out.append((name, arguments.get("cursor"), result.to_json()))
            last = (name, result, arguments)
    return out


@pytest.mark.parametrize("shared", [False, True])
def test_every_tool_gives_what_it_gave_the_long_way_round(monkeypatch, shared):
    # a query's pages read their rows, but for one that keeps its few cells
    monkeypatch.setattr(query, "KEEP_CELLS", 1000)
    data, other = shop_floor(shared=shared), shop_floor(shared=shared, changed=True)
    fast = _run(data, other, slow=False)
    index = Workbook.open(data).scan("Data").index
    assert index is not None and index.shared_formulas is shared
    with monkeypatch.context() as patched:
        _long_way_round(patched)
        assert Workbook.open(data).scan("Data").index is None
        slow = _run(data, other, slow=True)
    assert len(fast) == len(slow) > 20
    for one, two in zip(fast, slow):
        assert one == two, one[:2]


def _layout(kind: str) -> bytes:
    """A sheet of tables and blocks laid out to test which rows find_regions may skip."""
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    n = CHECKPOINT_ROWS * 4

    def table(name: str, corner: str, width: int, rows: int) -> None:
        top, left = int(corner[1:]), ord(corner[0]) - 64
        names = [f"{name}{k}" for k in range(width)]
        sheet.rows(corner, [names] + [[f"{name}-{r}-{k}" if k == 0 else r * k
                                       for k in range(width)] for r in range(rows)])
        sheet.table(name, f"{corner}:{chr(64 + left + width - 1)}{top + rows}", names)

    if kind == "beside":       # a block in the last column, beside a table, part-way down
        table("T", "A1", 4, n)
        sheet.rows(f"E{n // 2}", [["Note"], ["x"], ["y"]])
    elif kind == "between":    # blocks between and below tables that span the sheet
        table("T", "A1", 4, n)
        sheet.rows(f"A{n + 3}", [["Item", "Qty"], ["a", 1]])
        table("U", f"A{n + 7}", 4, n)
        sheet.rows(f"B{2 * n + 10}", [["Title"], ["Code", "Qty"], ["c", 2]])
    elif kind == "side":       # two tables side by side, together the sheet's width
        table("T", "A1", 2, n)
        table("U", "C1", 2, n // 2)
        sheet.rows(f"C{n // 2 + 4}", [["Late", "Qty"], ["d", 4]])
    else:                      # a table narrower than the sheet, a block to its left
        table("T", "C1", 2, n)
        sheet.rows(f"A{n - 5}", [["Left"], ["e"]])
    return book.to_bytes()


@pytest.mark.parametrize("kind", ["beside", "between", "side", "left"])
def test_find_regions_skips_only_rows_that_hold_nothing_outside_the_tables(monkeypatch, kind):
    data = _layout(kind)
    fast = Workbook.open(data)
    assert fast.scan("Data").index is not None
    found = fast.find_regions("Data")
    with monkeypatch.context() as patched:
        _long_way_round(patched)
        assert Workbook.open(data).find_regions("Data") == found
    assert found   # every layout has a block outside its tables
