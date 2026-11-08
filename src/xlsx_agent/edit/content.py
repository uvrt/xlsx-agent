"""Which cells of an area hold a value or a formula, read from the tree an edit works on.

A block is written (:mod:`.tables`) only where no cell holds a value or a formula, and rows
are appended below the last one that holds one.  The streamed pass over a sheet
(:meth:`Workbook.cells <xlsx_agent.Workbook.cells>`) types every cell to say so -- seconds
on a sheet of 50,000 rows, for an edit that writes a few.  Here the sheet's tree (the one
the cell writer parses for the write anyway) is read instead: its rows numbered
(:func:`rows`), only the rows inside the area placed cell by cell (:func:`columns`), and
only their cells inside the area asked whether they hold something, as the streamed pass
would type them (:func:`holds`).
"""

from __future__ import annotations

from typing import Any, Iterator

from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area

from ..read.stream import _C, _COLUMNS, _F, _IS, _REF, _ROW, _V, _column, _int

__all__ = ["columns", "holds", "rows", "with_content"]


def rows(root: Any) -> list[tuple[int, Any]]:
    """Every row of a parsed sheet with its number, in the file's order: a row without
    ``r`` (or with one that is no row) is the row after the one before it, as the streamed
    pass and the cell writer number it."""
    out = []
    row = 0
    for node in root.iter(*_ROW):
        number = _int(node.get("r"))
        row = number if number is not None and 1 <= number <= MAX_ROW else row + 1
        out.append((row, node))
    return out


def columns(row: int, node: Any) -> Iterator[tuple[int, Any]]:
    """``(column, c element)`` for each cell of row ``row`` (its ``row`` element
    ``node``), in the file's order: a cell without ``r`` is the column after the cell
    before it; a cell off the grid is left out.  A cell is taken to be in its row, as the
    cell writer takes it (one whose reference names another row is a structure issue,
    ``cell_row``)."""
    column = 0
    known = _COLUMNS
    match_ref = _REF.fullmatch
    digits = str(row)
    cut = -len(digits)
    for cell in node:
        if cell.tag not in _C:
            continue
        ref = cell.get("r")
        if ref and ref.endswith(digits) and ref[:cut] in known:
            column = known[ref[:cut]]
        elif (match := match_ref(ref) if ref else None) is not None:
            letters = match.group(1)
            column = known.get(letters) or _column(letters)
        else:
            column += 1
        if 1 <= column <= MAX_COLUMN:
            yield column, cell


def holds(node: Any, reader: Any, strings: int, row: int, column: int) -> bool:
    """Whether a cell holds a value or a formula as the streamed pass types it (a value
    that is not ``blank``, or a formula): ``strings`` is how many shared strings there
    are; ``reader`` a :class:`~xlsx_agent.read.stream.SheetReader` of the sheet, which
    types any cell but a shared string's."""
    has_value = False
    for child in node:
        tag = child.tag
        if tag in _F:
            return True
        if tag in _V or tag in _IS:
            has_value = True
    if not has_value:
        return False        # nothing a type could be read from: blank
    if (node.get("t") or "n") == "s":
        v_text = None
        for child in node:
            if child.tag in _V:
                v_text = child.text or ""
        index = _int(v_text)
        return index is not None and 0 <= index < strings
    return reader.value(row, column, False, node).t != "blank"


def with_content(workbook: Any, sheet: Any, area: Area, *,
                 last: bool = False) -> list[tuple[int, int]]:
    """The cells of ``area`` on ``sheet`` holding a value or a formula, as ``(row,
    column)`` in order; with ``last``, only the last of them (by row, then column), or
    none.  The sheet's tree is parsed when it is not yet (as the write that follows
    parses it); only the rows inside ``area`` are read, from the last one up for
    ``last`` when the rows are in order."""
    root = workbook.package.tree(sheet.part)
    if root is None:
        return []
    first, end = area.first, area.last
    numbered = rows(root)
    inside = [(number, node) for number, node in numbered if first.row <= number <= end.row]
    if not inside:
        return []
    ordered = all(a[0] < b[0] for a, b in zip(numbered, numbered[1:]))
    reader = workbook._reader(sheet, data=b"")
    strings = workbook._string_count()
    found: list[tuple[int, int]] = []
    for number, node in (reversed(inside) if last and ordered else inside):
        for column, cell in columns(number, node):
            if first.column <= column <= end.column and holds(cell, reader, strings, number,
                                                               column):
                found.append((number, column))
        if found and last and ordered:
            break
    found.sort()
    return found[-1:] if last else found
