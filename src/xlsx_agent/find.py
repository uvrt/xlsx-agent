"""Finding text: where words are in a workbook, by address and kind (what the shared
``find_text`` tool returns for a workbook)::

    for found in wb.find_text("Hardware"):            # a string, or a compiled re.Pattern
        found.address, found.kind, found.context()     # 'Stock!B7', 'cell', '...Hardware...'
    wb.find_text(re.compile(r"SKU-00\\d\\d"), within="t:Products")
    wb.find_text("Prices", within="Stock!")            # one sheet

**What is searched**, and the kind of each match:

* ``cell``: text cells, as the reader gives them (:mod:`xlsx_agent.read`): shared, inline
  and rich strings, a formula's cached text, numbers and dates stored as text.  A number,
  date, time, percent, boolean or error is **not** searched: no display string is made in
  0.1 (ROADMAP.md, "Values and types"), and the reader's form (``12.5``, ``2026-01-05``)
  is not what the sheet shows, so a match or a miss there would mislead; ``xlsx_query``
  finds values by type.
* ``table``: an Excel table's name; ``region``: a declared region's name.
* ``header``: an Excel table's column names (its ``tableColumn`` names, the ones its
  addresses use), each at ``t:Products[Price]``, with its header cell; the cells of a
  table's header row are searched as headers, not again as cells.
* ``comment``: a comment's text, at its cell (its author is not searched).
* ``name``: a defined name's name (its formula is not searched), hidden and built-in names
  among them.

**Formula text is not searched**: the shared tool has no option for it, and a formula's
text is given only to a call that asks for it (``xlsx_read`` with ``formulas``).  Nor are
hyperlinks, data-validation lists or number-format codes.

**Order**: sheet by sheet in the workbook's order; on a sheet, in reading order (row by
row, left to right; the file's order for a sheet whose rows are out of order, with its
comments, headers, tables and names after its cells): at one cell a table's or name's
name first, then a header, the cell, its comment.  Names that are not a range (a formula,
a constant, an unknown sheet) come last.  Each occurrence is a match; one of no length
(an empty regex match) is none.

**Hidden**: matches on a hidden or very hidden sheet, in a hidden row or column, and hidden
defined names are found like any other, and say ``hidden``.

**within**: a sheet (``Stock!``), a cell or range (``Stock!A1:H40``, ``Stock!A:C``), a
table, name or region (``t:Products``: all of it, header and totals rows too) or a part of
one (``t:Products[Price]``, ``t:Products/#5``: its data cells).  A table or name is found
when its range overlaps ``within``; a comment or header when its cell is inside it; a name
that is not a range never is.

**Content.**  A match holds the text it was found in (:attr:`Found.text`, kept out of its
``repr``); a :class:`Found` names it by address, kind and position.  Errors name the
problem, never the text searched for.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Iterator

from ooxml_edit.charts.workbook import Area, Cell

from .address import (AddressError, RangeAddress, TableAddress, parse_address, quote_sheet)

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: The kinds of match, in the order they come at one cell.
KINDS = ("table", "region", "name", "header", "cell", "comment")

#: Characters of context on either side of a match, and the most of a match shown.
CONTEXT = 40
MATCH_SHOWN = 80

#: Examples of what ``within`` takes, offered when it is something else.
WITHIN_FORMS = ["Stock!", "Stock!A1:H40", "t:Products"]

#: The cell types (ST_CellType) the reader may give as text: shared and inline strings and
#: a formula's text (:meth:`~xlsx_agent.read.stream.SheetReader.value`).
TEXT_KINDS = ("s", "inlineStr", "str")

_ORDER = {"table": 0, "region": 0, "name": 1, "header": 2, "cell": 3, "comment": 4}
_CELL_SHEETS = ("worksheet", "macrosheet")


@dataclass(frozen=True)
class Found:
    """One match: ``kind`` (:data:`KINDS`); ``address`` where it is (a cell, ``t:Products``,
    ``t:Products[Price]``, ``t:Stock!Prices``; ``None`` for a defined name an address
    cannot hold, whose ``name`` is then given); ``cell`` a header's cell, ``range`` the
    cells a table, region or name covers; ``hidden`` on a hidden sheet, row or column, or a
    hidden name.  ``text`` is what was searched (content: not in the ``repr``), ``start``
    and ``end`` the match in it."""

    kind: str
    address: str | None
    text: str = field(repr=False)
    start: int
    end: int
    hidden: bool = False
    cell: str | None = None
    range: str | None = None
    name: str | None = None

    def context(self, width: int = CONTEXT, shown: int = MATCH_SHOWN) -> str:
        """The match with ``width`` characters either side, ``…`` where the text goes on; a
        match longer than ``shown`` characters is cut there."""
        text, start, end = self.text, self.start, self.end
        left = max(0, start - width)
        if end - start > shown:
            return ("…" if left else "") + text[left:start + shown] + "…"
        right = min(len(text), end + width)
        return ("…" if left else "") + text[left:right] + ("…" if right < len(text) else "")

    def as_dict(self) -> dict:
        """``{"address", "kind", "context"}`` and whichever of the rest are set."""
        out: dict = {"address": self.address, "kind": self.kind}
        for key in ("name", "cell", "range"):
            value = getattr(self, key)
            if value is not None:
                out[key] = value
        out["context"] = self.context()
        if self.hidden:
            out["hidden"] = True
        return out


def pattern_of(needle: "str | re.Pattern") -> re.Pattern:
    """``needle`` as a pattern: a string found as it is written (case counts), or a compiled
    pattern as it is."""
    if isinstance(needle, re.Pattern):
        return needle
    if not isinstance(needle, str) or not needle:
        raise TypeError("find_text takes a non-empty string or a compiled re.Pattern")
    return re.compile(re.escape(needle))


def find_text(workbook: "Workbook", needle: "str | re.Pattern",
              within: "str | RangeAddress | TableAddress | None" = None) -> Iterator[Found]:
    """See :meth:`Workbook.find_text <xlsx_agent.workbook.Workbook.find_text>`."""
    pattern = pattern_of(needle)
    limit = _within(workbook, within) if within is not None else None
    for sheet in workbook.sheets:
        if sheet.part is None or sheet.type not in _CELL_SHEETS:
            continue
        if limit is not None and limit[0] != sheet:
            continue
        yield from _sheet(workbook, sheet, pattern, None if limit is None else limit[1])
    if limit is None:
        for name in workbook.names:
            if name.range is None or _cell_sheet_of(workbook, name.range) is None:
                yield from _name(name, pattern, sheet_hidden=False)


def _within(workbook: "Workbook", within) -> "tuple[Sheet, Area | None]":
    """The sheet and the cells ``within`` names (``None``: none of them)."""
    address = parse_address(within) if isinstance(within, str) else within
    if isinstance(address, TableAddress):
        if address.column is None and address.key is None and address.position is None:
            region = workbook.region(address)
            return workbook._cell_sheet(workbook.sheet_of(region.range)), region.area
        selection = workbook.resolve(address)
        sheet = workbook._cell_sheet(workbook.sheet_of(selection.region.range))
        return sheet, (selection.range.area if selection.range is not None else None)
    if isinstance(address, RangeAddress):
        return workbook._cell_sheet(workbook.sheet_of(address)), address.area
    raise AddressError(f"find_text looks within a sheet, range or table, not a "
                       f"{getattr(address, 'kind', 'value')}", text=str(within),
                       valid_options=WITHIN_FORMS)


def _cell_sheet_of(workbook: "Workbook", address: RangeAddress) -> "Sheet | None":
    try:
        sheet = workbook.sheet_of(address)
    except KeyError:
        return None
    return sheet if sheet.type in _CELL_SHEETS else None


def _inside(cell: Cell, area: Area) -> bool:
    return area.first.row <= cell.row <= area.last.row \
        and area.first.column <= cell.column <= area.last.column


def _overlaps(one: Area, two: Area) -> bool:
    return not (one.last.row < two.first.row or two.last.row < one.first.row
                or one.last.column < two.first.column or two.last.column < one.first.column)


def _occurrences(pattern: re.Pattern, text: str) -> Iterator[tuple[int, int]]:
    for match in pattern.finditer(text):
        if match.end() > match.start():
            yield match.start(), match.end()


def _name(name, pattern: re.Pattern, *, sheet_hidden: bool) -> Iterator[Found]:
    address = name.address
    for start, end in _occurrences(pattern, name.name):
        yield Found("name", address.text if address is not None else None, name.name, start,
                    end, hidden=name.hidden or sheet_hidden,
                    range=name.range.text if name.range is not None else None,
                    name=None if address is not None else name.name)


def _sheet(workbook: "Workbook", sheet: "Sheet", pattern: re.Pattern,
           area: Area | None) -> Iterator[Found]:
    """One sheet's matches: its cells streamed, the rest (tables, regions, names, headers,
    comments) merged in at their cells."""
    hidden_sheet = not sheet.visible
    written = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
    # What lies at a cell besides the cell itself: (row, column, order, kind, make), where
    # make(hidden) gives its matches once whether its row or column is hidden is known.
    extras: list[tuple[int, int, int, int, object]] = []
    headers: list[Area] = []

    def add(cell: Cell, kind: str, make) -> None:
        extras.append((cell.row, cell.column, _ORDER[kind], len(extras), make))

    for region in [*workbook.tables, *workbook.regions.declared]:
        if region.sheet != sheet.name or (area is not None and not _overlaps(region.area, area)):
            continue
        kind = "table" if region.kind == "table" else "region"
        text = region.address.text
        span = region.range.text
        add(region.area.first, kind, lambda hidden, r=region, k=kind, a=text, s=span: [
            Found(k, a, r.name, start, end, hidden=hidden_sheet, range=s)
            for start, end in _occurrences(pattern, r.name)])
        if region.kind != "table":
            continue
        header = region.header
        if header is not None:
            headers.append(header)
        row = header.first.row if header is not None else region.area.first.row
        for index, column in enumerate(region.columns):
            cell = Cell(row, region.area.first.column + index)
            if area is not None and not _inside(cell, area):
                continue
            where = f"{written}!{cell.ref}" if header is not None else None
            add(cell, "header", lambda hidden, c=column, w=where, r=region: [
                Found("header", replace(r.address, column=c).text, c, start, end,
                      hidden=hidden, cell=w)
                for start, end in _occurrences(pattern, c)])
    for name in workbook.names:
        if name.range is None or (area is not None and not _overlaps(name.range.area, area)):
            continue
        found = _cell_sheet_of(workbook, name.range)
        if found is None or found != sheet:
            continue
        add(name.range.area.first, "name",
            lambda hidden, n=name: list(_name(n, pattern, sheet_hidden=hidden_sheet)))
    for comment in workbook.comments(sheet):
        if area is not None and not _inside(comment.cell, area):
            continue
        add(comment.cell, "comment", lambda hidden, c=comment, w=f"{written}!{comment.cell.ref}": [
            Found("comment", w, c.text, start, end, hidden=hidden)
            for start, end in _occurrences(pattern, c.text)])
    extras.sort(key=lambda item: item[:4])

    reader = workbook._reader(sheet)
    scanned = workbook.scan(sheet)
    ordered = scanned.ordered
    # Within an area of an ordered sheet, the pass starts at the last checkpoint above it.
    start = None
    if area is not None and ordered and scanned.index is not None:
        checkpoint = scanned.index.before(area.first.row)
        start = (scanned.index, checkpoint) if checkpoint is not None else None

    def flush(upto: tuple[int, int, int] | None) -> Iterator[Found]:
        nonlocal next_extra
        while next_extra < len(extras) and (upto is None or extras[next_extra][:3] < upto):
            row, column = extras[next_extra][:2]
            make = extras[next_extra][4]
            next_extra += 1
            hidden = hidden_sheet or row in reader.hidden_rows or reader.column_hidden(column)
            yield from make(hidden)

    # Whether each shared string holds a match, as it is first met: a cell whose string
    # holds none is not read further.
    strings = workbook.strings
    holds: dict[int, bool] = {}
    next_extra = 0
    for row, column, row_hidden, node in workbook._stream(sheet, lambda: reader._walk(start)):
        if area is not None and not _inside(Cell(row, column), area):
            if ordered and row > area.last.row:
                break
            continue
        if ordered and next_extra < len(extras):
            yield from flush((row, column, _ORDER["cell"]))
        kind = node.get("t")
        if kind not in TEXT_KINDS:   # a cell of no other type is ever text
            continue
        if kind == "s":
            index = reader.shared_index(node)
            if index is not None and 0 <= index < len(strings):
                found = holds.get(index)
                if found is None:
                    found = holds[index] = next(_occurrences(pattern, strings[index]),
                                                None) is not None
                if not found:
                    continue
        value = reader.value(row, column, row_hidden, node)
        if value.t != "text" or not isinstance(value.v, str):
            continue
        if headers and any(_inside(Cell(row, column), h) for h in headers):
            continue
        for begin, end in _occurrences(pattern, value.v):
            yield Found("cell", value.a, value.v, begin, end, hidden=value.hidden)
    yield from flush(None)
