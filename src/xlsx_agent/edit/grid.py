"""Inserting, deleting and moving rows and columns (round 44): Excel's own operations, with
every reference to the cells rewritten as Excel rewrites it.

::

    wb.insert_rows("Stock", 5, 2)          # two empty rows before row 5
    wb.delete_columns("Stock", "C", 1)     # column C goes; D becomes C
    wb.move_columns("Stock", "J", "A")     # Excel's Cut + "Insert Cut Cells": J before A
    wb.move_rows("Stock", "12:13", 2)      # rows 12 and 13 placed before row 2

Each is one step (or part of the open ``wb.batch()``), undoable, and refused before
anything changes -- every problem listed (:class:`~xlsx_agent.edit.EditError`, its
``problems``) -- when Excel refuses the same operation or when a dependent's adjustment has
not been measured from Excel yet.  Nothing is ever written half-updated.

**What moves, measured** (Excel for Mac 16.106, 2026-11-08; ``tests/structure_book.py`` and
``tests/test_structure_oracle.py``; the table of what was measured is in ROADMAP.md, round 44):

* *References* (in cells on every sheet, defined names, data validation and conditional
  format formulas, chart series, table formulas, pivot cache sources): a cell reference
  follows its cell; a reference to a deleted cell becomes ``#REF!`` (``Stock!#REF!`` with
  its sheet); a range grows when lines are inserted after its first line and up to its
  last, moves when they are inserted at or before its first line, shrinks when some of its
  lines are deleted and becomes ``#REF!`` when all are.  A reference pushed off the grid
  is ``#REF!``; a range whose end is pushed off is cut at the grid's edge
  (``B2:B1048576`` stays).  Whole columns do not change when rows change, nor whole rows
  when columns do.  3-D references (``Stock:Notes!B2``), other workbooks' cells, text
  (``INDIRECT("Stock!B3")``) and a defined name's relative references (``Rel``:
  ``Stock!B3``) are left as they are -- Excel leaves them.
* *A table's data rows* (rows only): a range whose rows are exactly a table's data rows
  and whose columns lie within the table's (``B2:B7``, ``$A$2:$F$7``) grows when rows are
  inserted at the table's first data row, where any other range moves.
* *Structured references* to a deleted table column, or to a table every column of which
  was deleted, become ``#REF!``.
* *Inserted lines take the format of the line before them* (above, or to the left):
  a row's height and style, a column's width and style, and the data validations and
  conditional formats of that line -- not hiding, and nothing from a hidden line.
* *Merged cells, data validations, conditional formats, hyperlinks, comments, filters,
  tables, print areas and titles, frozen panes, page breaks* move, grow and shrink with
  their cells; what is wholly deleted goes (a merge left one cell goes too).  A
  conditional format's or validation's formula keeps its meaning for every cell it covers
  (Excel splits a block whose cells would no longer share one formula; so does this).
* *Tables*: inserting columns inside a table adds table columns (``Column1``, the first
  free such name, as an English Excel names them); deleting a table's columns removes
  them, and deleting all of them the table.  Deleting every data row of a table keeps one
  empty data row, as Excel does.
* *Moves* are Excel's Cut and "Insert Cut Cells": the cells, their formats and everything
  attached go to the new place, and the lines between close up.

**Refused, as Excel refuses**: cutting through an array formula (inserting inside it,
deleting or moving part of it); pushing cells that hold something off the grid; deleting a
table's header row; moving columns of a table.  **Refused, not measured yet**: a data
table, a pivot table the change would cut through, a structured reference to a column
range a deletion reaches, a defined name with a mixed reference to the sheet, a moved
block that cuts through a merged area or crosses a table's edge, a table's totals row,
and parts of a sheet this module does not rewrite (form objects, scenarios, data
consolidation, custom views, threaded comments are rewritten).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable

from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, column_index, column_letters

from ..formula import ShiftError, shift as shift_formula, tokens

_M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
X = f"{{{_M}}}"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
XM = "{http://schemas.microsoft.com/office/excel/2006/main}"
X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
VX = "{urn:schemas-microsoft-com:office:excel}"
TC = "{http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments}"

ROW, COLUMN = "row", "column"
#: A reference that no longer names a cell.
REF = "#REF!"
_THREE_D = object()

_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})")
_COLUMN = re.compile(r"(\$?)([A-Za-z]{1,3})")
_ROW = re.compile(r"(\$?)([0-9]{1,7})")
_PLAIN_CELL = re.compile(r"([A-Z]{1,3})([0-9]{1,7})")


class NotDone(ValueError):
    """A dependent that cannot be rewritten as Excel would: the operation is refused.
    ``code`` is an :data:`~xlsx_agent.edit.EDIT_CODES` code, ``where`` an address or a
    part, ``detail`` what it is (never content)."""

    def __init__(self, code: str, where: str, detail: str) -> None:
        super().__init__(f"{code} at {where}: {detail}")
        self.code, self.where, self.detail = code, where, detail


# -- the change along one axis -----------------------------------------------------------------


@dataclass(frozen=True)
class Change:
    """One operation along one axis of one sheet: ``kind`` ``"insert"``, ``"delete"`` or
    ``"move"``; ``axis`` ``"row"`` or ``"column"``; the lines ``first`` .. ``first +
    count - 1`` (inserted before ``first``; deleted; moved), and for a move ``before``: the
    line (numbered as before the move) the block is placed before.  ``bodies`` are the
    tables' data bodies on the sheet as ``(first row, first column, last row, last
    column)``: a range equal to one in rows, within it in columns, is the table's (the
    first-data-row rule of the module's notes)."""

    kind: str
    axis: str
    first: int
    count: int
    before: int = 0
    bodies: tuple[tuple[int, int, int, int], ...] = ()

    @property
    def last(self) -> int:
        return self.first + self.count - 1

    @property
    def limit(self) -> int:
        return MAX_ROW if self.axis == ROW else MAX_COLUMN

    @property
    def target(self) -> int:
        """A move's first line once moved."""
        return self.before if self.before < self.first else self.before - self.count

    def index(self, i: int) -> int | None:
        """Where line ``i`` goes; ``None`` when it is deleted or pushed off the grid."""
        if self.kind == "insert":
            j = i if i < self.first else i + self.count
            return j if j <= self.limit else None
        if self.kind == "delete":
            if i < self.first:
                return i
            return None if i <= self.last else i - self.count
        first, last, before = self.first, self.last, self.before
        if before < first:                       # up
            if before <= i < first:
                return i + self.count
            if first <= i <= last:
                return before + i - first
            return i
        if last < i < before:                    # down
            return i - self.count
        if first <= i <= last:
            return before - self.count + i - first
        return i

    def origin(self, j: int) -> int | None:
        """The line that ``j`` was before the change; ``None`` for an inserted line."""
        if self.kind == "insert":
            if j < self.first:
                return j
            return None if j < self.first + self.count else j - self.count
        if self.kind == "delete":
            return j if j < self.first else j + self.count
        for i in self._candidates(j):
            if self.index(i) == j:
                return i
        return j

    def _candidates(self, j: int) -> Iterable[int]:
        yield j
        yield j - self.count
        yield j + self.count
        yield j - (self.target - self.first)
        yield j + (self.target - self.first)

    def span(self, a: int, b: int, across: "tuple[int, int] | None" = None
             ) -> tuple[int, int] | None:
        """Where the range of lines ``a`` .. ``b`` goes (``None``: ``#REF!``);
        ``across`` its lines on the other axis (``None`` for whole lines)."""
        if self.kind == "insert":
            if (self.axis == ROW and across is not None and a == self.first
                    and any(a == r1 and b == r2 and c1 <= across[0] and across[1] <= c2
                            for r1, c1, r2, c2 in self.bodies)):
                return a, min(b + self.count, self.limit)     # the table's data rows
            na = a if a < self.first else a + self.count
            if na > self.limit:
                return None
            nb = b if b < self.first else b + self.count
            return na, min(nb, self.limit)
        if self.kind == "delete":
            first, last = self.first, self.last
            if first <= a and b <= last:
                return None
            na = a if a < first else (first if a <= last else a - self.count)
            nb = b if b < first else (first - 1 if b <= last else b - self.count)
            if b == self.limit and nb >= na:
                nb = self.limit          # the grid refills from its edge (B2:B1048576 stays)
            return na, nb
        return self._moved_span(a, b, across)

    def _moved_span(self, a: int, b: int, across: "tuple[int, int] | None"
                    ) -> tuple[int, int] | None:
        """A range through a move, as measured (``tests/structure_book.py``'s grid):

        * a range inside the moved block moves with it;
        * a range the block lies in, the block placed inside it again (after its first
          line, up to its last), stays as it is;
        * a range one of whose ends is in the block: that end follows the block, unless
          the range would turn over;
        * otherwise -- and when it would turn over -- the range is shrunk as by deleting
          the block, then grown or moved as by inserting it at its new place (where the
          table rule of the module's notes applies)."""
        first, last, before = self.first, self.last, self.before
        ia, ib = self.index(a), self.index(b)
        assert ia is not None and ib is not None
        if first <= a and b <= last:
            return ia, ib
        if a <= first and last <= b:
            if a < before <= b:
                return a, b
            if (first == a or last == b) and ia <= ib:
                return ia, ib
            return self._delete_insert(a, b, across)
        if b < first or last < a:
            return self._delete_insert(a, b, across)
        if ia <= ib:
            return ia, ib
        return self._delete_insert(a, b, across)

    def _delete_insert(self, a: int, b: int, across: "tuple[int, int] | None"
                       ) -> tuple[int, int] | None:
        removed = Change("delete", self.axis, self.first, self.count)
        deleted = removed.span(a, b)
        if deleted is None:
            return None
        bodies = []
        for r1, c1, r2, c2 in self.bodies:
            rows = removed.span(r1, r2) if self.axis == ROW else (r1, r2)
            if rows is not None:
                bodies.append((rows[0], c1, rows[1], c2))
        return Change("insert", self.axis, self.target, self.count,
                      bodies=tuple(bodies)).span(*deleted, across)

    def pieces(self, a: int, b: int, across: "tuple[int, int] | None" = None
               ) -> list[tuple[int, int]]:
        """Where the cells of lines ``a`` .. ``b`` go, as runs of lines: what a format, a
        rule or a table's cells follow (for a move, the cells themselves; for an insert
        or a delete, :meth:`span`)."""
        if self.kind != "move":
            span = self.span(a, b, across)
            return [] if span is None else [span]
        cuts = sorted({a, b + 1} | {x for x in (self.first, self.last + 1, self.before)
                                     if a < x <= b})
        out = []
        for lo, hi in zip(cuts, cuts[1:]):
            x, y = self.index(lo), self.index(hi - 1)
            assert x is not None and y is not None
            out.append((x, y))
        return sorted(out)

    def splits(self, a: int, b: int) -> bool:
        """Whether the lines ``a`` .. ``b`` (a block that changes only as a whole: an
        array formula) would be cut through."""
        if self.kind == "insert":
            return a < self.first <= b
        if self.kind == "delete":
            overlap = not (b < self.first or self.last < a)
            return overlap and not (self.first <= a and b <= self.last)
        inside = self.first <= a and b <= self.last
        outside = b < self.first or self.last < a
        if not (inside or outside):
            return True
        return outside and a < self.before <= b


# -- references --------------------------------------------------------------------------------


@dataclass
class _Part:
    column_anchor: str
    column: int | None
    row_anchor: str
    row: int | None

    def text(self) -> str:
        out = ""
        if self.column is not None:
            out += self.column_anchor + column_letters(self.column)
        if self.row is not None:
            out += self.row_anchor + str(self.row)
        return out


def _parse_part(text: str) -> _Part | None:
    cell = _CELL.fullmatch(text)
    if cell is not None:
        return _Part(cell.group(1), column_index(cell.group(2)), cell.group(3),
                     int(cell.group(4)))
    column = _COLUMN.fullmatch(text)
    if column is not None:
        return _Part(column.group(1), column_index(column.group(2)), "", None)
    row = _ROW.fullmatch(text)
    if row is not None:
        return _Part("", None, row.group(1), int(row.group(2)))
    return None


def sheet_key(name: str) -> str:
    """A sheet's name as Excel compares names in references: case aside."""
    return name.casefold()


def _prefix_key(prefix: str) -> Any:
    name = prefix[:-1] if prefix.endswith("!") else prefix
    if name.startswith("'") and name.endswith("'") and len(name) > 1:
        name = name[1:-1].replace("''", "'")
    if ":" in name:
        return _THREE_D
    return sheet_key(name)


@dataclass
class Tables:
    """What happened to tables, for structured references: by table name (case folded),
    the columns deleted (names, case folded), or ``None`` when the whole table went."""

    deleted: dict[str, "set[str] | None"] = field(default_factory=dict)


class Rewriter:
    """Rewrites references to one sheet (:attr:`target`, case folded) through a
    :class:`Change`."""

    def __init__(self, change: Change, target: str, tables: Tables | None = None) -> None:
        self.change = change
        self.target = sheet_key(target)
        self.name = target
        self.tables = tables or Tables()
        self._needles = {target.casefold(), target.replace("'", "''").casefold()}

    def mentions(self, text: str) -> bool:
        """Whether ``text`` may refer to the target sheet by name (a fast first look)."""
        folded = text.casefold()
        return any(needle in folded for needle in self._needles)

    def formula(self, text: str, home: str | None, table: str | None = None) -> str:
        """``text`` (a formula, no ``=``) with its references to the target sheet
        rewritten; ``home`` the sheet the formula is on (case folded; its unqualified
        references are to it), ``None`` for one on no sheet; ``table`` the table a
        structured reference without a table name is to.  :class:`ShiftError` when the text
        cannot be read; :class:`NotDone` for what is not measured."""
        on_home = home is not None and home == self.target
        if not on_home and not self.mentions(text) and not self._structured_matter(text):
            return text
        out: list[str] = []
        changed = False
        for token in tokens(text):
            new = token.text
            if token.kind == "ref" and not token.book:
                key = _prefix_key(token.sheet) if token.sheet else None
                if key is _THREE_D:
                    pass
                elif (key is None and on_home) or (key is not None and key == self.target):
                    new = token.sheet + self.body(token.body)
            elif token.kind == "structured" and self.tables.deleted:
                new = self._structured(token.text, table)
            changed = changed or new != token.text
            out.append(new)
        return "".join(out) if changed else text

    def _structured_matter(self, text: str) -> bool:
        return bool(self.tables.deleted) and "[" in text

    def body(self, body: str) -> str:
        """A reference's body (``$B$2:C7``, ``B:B``, ``2:7``) rewritten; ``#REF!``."""
        pieces = body.split(":")
        if len(pieces) > 2:
            raise NotDone("not_measured", body, "a reference of more than two parts")
        parts = [_parse_part(piece) for piece in pieces]
        if any(part is None for part in parts):
            raise NotDone("not_measured", body, "a reference that cannot be read")
        change = self.change
        row_axis = change.axis == ROW

        def along(part: _Part) -> int | None:
            return part.row if row_axis else part.column

        def set_along(part: _Part, value: int) -> None:
            if row_axis:
                part.row = value
            else:
                part.column = value

        if len(parts) == 1:
            part = parts[0]
            assert part is not None
            position = along(part)
            if position is None:
                return body
            if change.kind == "move" and position == change.limit \
                    and not change.first <= position <= change.last:
                return REF       # a move inserts first: the last line is pushed off
            moved = change.index(position)
            if moved is None:
                return REF
            set_along(part, moved)
            return part.text()
        first, second = parts
        assert first is not None and second is not None
        a, b = along(first), along(second)
        if a is None or b is None:
            return body
        if min(a, b) == 1 and max(a, b) == change.limit:
            return body          # every line: Excel reads B1:B1048576 as B:B (measured)
        if a > b:
            first, second, a, b = second, first, b, a
        cross_a = first.column if row_axis else first.row
        cross_b = second.column if row_axis else second.row
        across = None if cross_a is None or cross_b is None else \
            (min(cross_a, cross_b), max(cross_a, cross_b))
        span = change.span(a, b, across)
        if span is None:
            return REF
        set_along(first, span[0])
        set_along(second, span[1])
        return f"{first.text()}:{second.text()}"

    def _structured(self, text: str, table: str | None) -> str:
        cut = text.find("[")
        name = text[:cut].strip().casefold() or (table or "").casefold()
        if name not in self.tables.deleted:
            return text
        gone = self.tables.deleted[name]
        if gone is None:
            return REF
        items, ranged = _structured_columns(text[cut:])
        hit = [item for item in items if item.casefold() in gone]
        if hit and ranged:
            raise NotDone("not_measured", text[:cut] or "a table",
                          "a structured reference to a range of columns, one of them deleted")
        return REF if hit else text


def _structured_columns(text: str) -> tuple[list[str], bool]:
    """The column names a structured reference's brackets name, and whether it names a
    range of columns (``[[Qty]:[Price]]``)."""
    names: list[str] = []
    depth, start, i, ranged = 0, -1, 0, False
    while i < len(text):
        char = text[i]
        if char == "'":
            i += 2
            continue
        if char == "[":
            depth += 1
            start = i + 1
        elif char == "]":
            if start >= 0:
                item = re.sub(r"'(.)", r"\1", text[start:i]).strip()
                if item.startswith("@"):
                    item = item[1:].strip()
                    if item.startswith("[") and item.endswith("]"):
                        item = item[1:-1]
                if item and not item.startswith("#") and "[" not in item:
                    names.append(item)
                start = -1
            depth -= 1
        elif char == ":" and depth == 1:
            ranged = True
        i += 1
    return names, ranged


# -- anchored formulas (data validations, conditional formats) ---------------------------------


@dataclass(frozen=True)
class Area:
    """A rectangle: rows ``r1`` .. ``r2``, columns ``c1`` .. ``c2``."""

    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def ref(self) -> str:
        first = f"{column_letters(self.c1)}{self.r1}"
        if (self.r1, self.c1) == (self.r2, self.c2):
            return first
        return f"{first}:{column_letters(self.c2)}{self.r2}"

    def along(self, axis: str) -> tuple[int, int]:
        return (self.r1, self.r2) if axis == ROW else (self.c1, self.c2)

    def across(self, axis: str) -> tuple[int, int]:
        return (self.c1, self.c2) if axis == ROW else (self.r1, self.r2)

    def with_along(self, axis: str, a: int, b: int) -> "Area":
        return Area(a, self.c1, b, self.c2) if axis == ROW else Area(self.r1, a, self.r2, b)


def parse_area(text: str) -> Area | None:
    """``B2``, ``B2:C7`` (anchors allowed) as an :class:`Area`; ``None`` otherwise."""
    pieces = text.replace("$", "").split(":")
    if len(pieces) == 1:
        pieces = pieces * 2
    if len(pieces) != 2:
        return None
    first, second = _CELL.fullmatch(pieces[0]), _CELL.fullmatch(pieces[1])
    if first is None or second is None:
        return None
    r1, r2 = int(first.group(4)), int(second.group(4))
    c1, c2 = column_index(first.group(2)), column_index(second.group(2))
    return Area(min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def map_area(area: Area, change: Change) -> Area | None:
    """Where a rectangle of cells goes as a reference to it (``None``: ``#REF!``)."""
    a, b = area.along(change.axis)
    span = change.span(a, b, area.across(change.axis))
    if span is None:
        return None
    return area.with_along(change.axis, *span)


def _linked_move(area: Area, change: Change) -> bool:
    """A rule's area that is a table's data rows, and rows moved into the table from
    outside: it grows as a reference does (measured)."""
    a, b = area.along(change.axis)
    return change.kind == "move" and change.axis == ROW and any(
        (a, b) == (r1, r2) and c1 <= area.c1 and area.c2 <= c2
        for r1, c1, r2, c2 in change.bodies) and (b < change.first or change.last < a)


def area_cells(area: Area, change: Change, *, linked: bool = False) -> list[Area]:
    """Where a rectangle's cells go, as rectangles (a format, a rule, a merged area or a
    table follows its cells): for a move, each run of its lines to its place; for an
    insert or a delete, as :func:`map_area`.  ``linked``: a rule's area that is a table's
    data rows grows with rows moved into the table, as a reference does (measured)."""
    a, b = area.along(change.axis)
    if linked and _linked_move(area, change):
        span = change.span(a, b, area.across(change.axis))
        return [] if span is None else [area.with_along(change.axis, *span)]
    return [area.with_along(change.axis, x, y)
            for x, y in change.pieces(a, b, area.across(change.axis))]


def _relative_offsets(text: str, axis: str) -> set[int]:
    """The positions along ``axis`` a formula's relative references name (as written at
    its anchor): where the cells it is copied to may fall on either side of a change."""
    found: set[int] = set()
    for token in tokens(text):
        if token.kind != "ref" or token.book:
            continue
        for piece in token.body.split(":"):
            part = _parse_part(piece)
            if part is None:
                continue
            if axis == ROW and part.row is not None and not part.row_anchor:
                found.add(part.row)
            if axis == COLUMN and part.column is not None and not part.column_anchor:
                found.add(part.column)
    return found


def _breaks(change: Change) -> list[int]:
    """The lines at which a change's behaviour changes along its axis."""
    if change.kind == "insert":
        return [change.first]
    if change.kind == "delete":
        return [change.first, change.last + 1]
    return sorted({change.first, change.last + 1, change.before})


def anchored(areas: list[Area], formulas: list[str], change: Change, rewriter: Rewriter,
             home: str, *, linked: bool = True) -> list[tuple[list[Area], list[str]]]:
    """A rule over ``areas`` whose ``formulas`` are written for the top-left cell of the
    areas' bounding box, through ``change``: groups of ``(areas, formulas)``, each group's
    formulas written for its own top-left cell -- one group when every cell keeps one
    formula, as many as the cells' formulas now differ (Excel splits such a rule)."""
    axis = change.axis
    top = min(a.r1 for a in areas), min(a.c1 for a in areas)
    pieces: list[tuple[Area, list[str]]] = []
    breaks = _breaks(change)
    for area in areas:
        dr, dc = area.r1 - top[0], area.c1 - top[1]
        local = [shift_formula(f, dr, dc) if f else f for f in formulas]
        lo, hi = area.along(axis)
        anchor = lo
        cuts = {lo}
        whole = linked and _linked_move(area, change)
        if change.kind == "move" and not whole:   # the cells part at a move's edges
            cuts |= {line for line in breaks if lo < line <= hi}
        for text in ([] if whole else local):
            if not text:
                continue
            for position in _relative_offsets(text, axis):
                offset = position - anchor
                for line in breaks:
                    cut = line - offset
                    if lo < cut <= hi:
                        cuts.add(cut)
        edges = sorted(cuts) + [hi + 1]
        for start, end in zip(edges, edges[1:]):
            piece = area.with_along(axis, start, end - 1)
            step = start - anchor
            texts = [shift_formula(f, step if axis == ROW else 0, step if axis == COLUMN else 0)
                     if f else f for f in local]
            written = [rewriter.formula(t, home) if t else t for t in texts]
            for moved in area_cells(piece, change, linked=linked):
                went = change.index(piece.along(axis)[0])
                start = moved.along(axis)[0]
                here = written
                if went is not None and went != start:
                    delta = start - went
                    here = [shift_formula(t, delta if axis == ROW else 0,
                                          delta if axis == COLUMN else 0) if t else t
                            for t in written]
                pieces.append((moved, here))
    return _grouped(pieces)


def _grouped(pieces: list[tuple[Area, list[str]]]) -> list[tuple[list[Area], list[str]]]:
    """Pieces that read alike from one top-left cell, together (in order of first
    appearance); each group's formulas written for its own top-left cell."""
    groups: list[tuple[list[Area], list[tuple[Area, list[str]]]]] = []
    for area, texts in pieces:
        for areas, members in groups:
            if _agree(members + [(area, texts)]):
                areas.append(area)
                members.append((area, texts))
                break
        else:
            groups.append(([area], [(area, texts)]))
    out = []
    for areas, members in groups:
        top = min(a.r1 for a in areas), min(a.c1 for a in areas)
        area, texts = members[0]
        out.append((_merged(areas), [shift_formula(t, top[0] - area.r1, top[1] - area.c1)
                                     if t else t for t in texts]))
    return out


def _agree(members: list[tuple[Area, list[str]]]) -> bool:
    top = min(a.r1 for a, _ in members), min(a.c1 for a, _ in members)
    seen: list[str] | None = None
    for area, texts in members:
        try:
            here = [shift_formula(t, top[0] - area.r1, top[1] - area.c1) if t else t
                    for t in texts]
        except ShiftError:
            return False
        if seen is None:
            seen = here
        elif here != seen:
            return False
    return True


def _merged(areas: list[Area]) -> list[Area]:
    """Areas that touch along a line, joined (``C2:C7`` and ``C8`` as ``C2:C8``)."""
    out: list[Area] = []
    for area in sorted(areas, key=lambda a: (a.c1, a.c2, a.r1)):
        last = out[-1] if out else None
        if last is not None and (last.c1, last.c2) == (area.c1, area.c2) \
                and last.r2 + 1 >= area.r1:
            out[-1] = Area(last.r1, last.c1, max(last.r2, area.r2), last.c2)
            continue
        out.append(area)
    joined: list[Area] = []
    for area in sorted(out, key=lambda a: (a.r1, a.r2, a.c1)):
        last = joined[-1] if joined else None
        if last is not None and (last.r1, last.r2) == (area.r1, area.r2) \
                and last.c2 + 1 >= area.c1:
            joined[-1] = Area(last.r1, last.c1, last.r2, max(last.c2, area.c2))
            continue
        joined.append(area)
    return joined


def sqref_areas(text: str) -> list[Area] | None:
    areas = [parse_area(piece) for piece in text.split()]
    if any(area is None for area in areas):
        return None
    return [area for area in areas if area is not None]


def sqref_text(areas: Iterable[Area]) -> str:
    return " ".join(area.ref for area in areas)


__all__ = ["Area", "COLUMN", "Change", "NotDone", "REF", "ROW", "Rewriter", "Tables",
           "anchored", "map_area", "parse_area", "sheet_key", "sqref_areas", "sqref_text"]
