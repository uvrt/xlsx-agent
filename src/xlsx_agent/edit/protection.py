"""Protected sheets (round 46, R71): which cells may still be written, and what else a
sheet's protection allows.

A protected sheet (``sheetProtection sheet="1"``, ECMA-376 Part 1, 18.3.1.85) keeps its
*locked* cells as they are; its *unlocked* cells stay open to input.  A cell is locked
unless its own cell format (``cellXfs``) says ``<protection locked="0"/>`` (18.8.33); a
position that holds no cell has the format a cell made there would take -- its row's
(``customFormat``), else its column's (ooxml-edit's ``style(cell, inherited=True)``).
``applyProtection`` and the cell style the format is based on (``xfId``) change nothing.

**Measured** (Excel for Mac 16.106, 2026-11-09, through its object model, on made-up sheets
protected with Excel's defaults -- ``sheet objects scenarios`` -- and with every "allow"
option set):

* *values and formulas* went into unlocked cells -- by their own format (with or without
  ``applyProtection``), their row's, their column's -- and were cleared there; a locked
  cell kept its value: one whose format says nothing (its cell style unlocked or not), a
  locked cell in an unlocked column.  So :meth:`Workbook.set_cells
  <xlsx_agent.Workbook.set_cells>`, ``write_table`` and ``append_rows`` write unlocked cells
  and refuse locked ones (``protected``: every locked cell named, the sheet's unlocked
  cells listed, and the way out).  A date, a date and time, a time or a line break typed
  into an unlocked General cell got Excel's format for it even where the protection allows
  no formatting.  There is no override for a locked cell: unprotecting the sheet is what
  that takes (:meth:`Workbook.unprotect_sheet <xlsx_agent.Workbook.unprotect_sheet>`).
* *tables*: none was made over unlocked cells, and none grew from a row typed below it or
  was resized, whatever the options: ``write_table(as_table=...)`` and ``append_rows`` to a
  table are refused there.
* *formatting* (bold, a number format) only where ``formatCells`` allows it, unlocked cells
  too (:mod:`.format`; widths and heights by ``formatColumns``/``formatRows``).
* *notes* were added to locked and unlocked cells alike under ``objects="0"`` and refused
  under ``objects="1"`` (:mod:`.notes`).
* *hyperlinks* only into an unlocked cell and only under ``insertHyperlinks="0"``
  (xlsx-agent writes none yet).
* *merging* was refused whatever the options (:mod:`.merges`).
* *inserting* rows or columns only under ``insertRows``/``insertColumns`` ``"0"``;
  *deleting* them only under ``deleteRows``/``deleteColumns`` ``"0"`` and only when every
  cell of the lines is unlocked (:func:`lines_problem`); inserting cells (shifting them
  down) was refused even then.

The ``sheetProtection`` element -- its hash, salt, spin count and options -- is never
touched by any of these.
"""

from __future__ import annotations

from typing import Any, Iterable

from ooxml_edit.sheets import Area, Cell

from ..address import format_cell, format_range

_X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

#: The most unlocked ranges a refusal lists as the way out.
MAX_HINTS = 8


def protection(workbook: Any, sheet: Any) -> Any | None:
    """The sheet's ``sheetProtection`` element when it turns protection on, else ``None``."""
    if sheet.part is None or not sheet.protected:
        return None
    root = workbook.package.tree(sheet.part)
    node = root.find(_X + "sheetProtection") if root is not None else None
    if node is None or node.get("sheet", "").strip().lower() not in ("1", "true"):
        return None
    return node


class Locks:
    """Which cells of the workbook's protected sheets are locked, the cell formats read
    once per call."""

    def __init__(self, workbook: Any) -> None:
        self.workbook = workbook
        self._locked: dict[int, bool] = {}
        styles = workbook.editor.cells.styles
        self._xfs = styles._xfs() if styles is not None else []

    def style_locked(self, style: int) -> bool:
        """Whether cell format ``style`` locks its cells (see the module)."""
        found = self._locked.get(style)
        if found is None:
            found = True
            if 0 <= style < len(self._xfs):
                node = self._xfs[style].find(_X + "protection")
                if node is not None:
                    found = _true(node.get("locked"), True)
            self._locked[style] = found
        return found

    def locked(self, writer: Any, cell: Cell) -> bool:
        """Whether ``cell`` of the sheet ``writer`` writes is locked."""
        return self.style_locked(writer.style(cell, inherited=True))

    def unlocked(self, sheet: Any, writer: Any,
                 limit: int = MAX_HINTS) -> tuple[list[str], bool]:
        """The sheet's unlocked cells as ranges (whole columns and rows a format unlocks,
        then blocks of unlocked cells), at most ``limit``, and whether there are more."""
        from .structure import _label

        label = _label(sheet)
        root = self.workbook.package.tree(sheet.part)
        found: list[str] = []
        if root is None:
            return found, False
        columns = root.find(_X + "cols")
        unlocked_columns: list[tuple[int, int]] = []
        if columns is not None:
            for column in columns.iterchildren(_X + "col"):
                low, high, style = column.get("min", ""), column.get("max", ""), \
                    column.get("style", "")
                if low.isdigit() and high.isdigit() and style.isdigit() \
                        and not self.style_locked(int(style)):
                    unlocked_columns.append((int(low), int(high)))
                    found.append(_columns(label, int(low), int(high)))
        data = root.find(_X + "sheetData")
        # Runs of unlocked cells by row, then runs alike in rows one below the other merged:
        # (first column, last column) -> [first row, last row], open while the rows follow.
        open_: dict[tuple[int, int], list[int]] = {}
        areas: list[tuple[int, int, int, int]] = []
        for row in (data.iterchildren(_X + "row") if data is not None else ()):
            number = row.get("r", "")
            if not number.isdigit():
                continue
            r = int(number)
            style = row.get("s", "")
            if row.get("customFormat", "").strip().lower() in ("1", "true") \
                    and style.isdigit() and not self.style_locked(int(style)):
                found.append(f"{label}!{r}:{r}")
                continue
            runs: list[tuple[int, int]] = []
            for cell in _cells(row, r):
                if any(low <= cell.column <= high for low, high in unlocked_columns) \
                        or self.locked(writer, cell):
                    continue
                if runs and runs[-1][1] == cell.column - 1:
                    runs[-1] = (runs[-1][0], cell.column)
                else:
                    runs.append((cell.column, cell.column))
            for run in list(open_):
                if run not in runs or open_[run][1] != r - 1:
                    first, last = open_.pop(run)
                    areas.append((first, run[0], last, run[1]))
            for run in runs:
                if run in open_:
                    open_[run][1] = r
                else:
                    open_[run] = [r, r]
            if len(found) + len(areas) > limit:
                break
        areas.extend((first, run[0], last, run[1]) for run, (first, last) in open_.items())
        for top, left, bottom, right in sorted(areas):
            area = Area(Cell(top, left), Cell(bottom, right))
            found.append(format_range(label, area) if area.first != area.last
                         else format_cell(label, top, left))
        more = len(found) > limit
        return found[:limit], more


def refusal(locks: Locks, sheet: Any, writer: Any) -> str:
    """A locked cell's refusal detail: why, the sheet's unlocked cells, and the way out."""
    from .structure import _label

    shown, more = locks.unlocked(sheet, writer)
    where = (f"; its unlocked cells: {', '.join(shown)}" + (" and more" if more else "")
             if shown else "; it has no unlocked cells")
    return (f"the cell is locked (sheet {_label(sheet)}'s sheetProtection): only unlocked "
            f"cells are written{where}. To write a locked cell, unprotect the sheet "
            "(wb.unprotect_sheet), unlock the cells (wb.format(..., protection={'locked': "
            "False})) and protect it again")


def _true(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in ("1", "true")


def _cells(row: Any, r: int) -> Iterable[Cell]:
    column = 0
    for node in row.iterchildren(_X + "c"):
        ref = node.get("r")
        if ref:
            try:
                cell = Area.parse(ref).first
            except ValueError:
                continue
            column = cell.column
        else:
            column += 1
        yield Cell(r, column)


def _columns(label: str, low: int, high: int) -> str:
    from ..read.types import _letters

    return f"{label}!{_letters(low)}:{_letters(high)}"


def _run(label: str, row: int, start: int, end: int) -> str:
    if start == end:
        return format_cell(label, row, start)
    return format_range(label, Area(Cell(row, start), Cell(row, end)))


__all__ = ["Locks", "MAX_HINTS", "protection", "refusal"]


#: The protection option that allows each insert and delete of the grid (ECMA-376 Part 1,
#: 18.3.1.85), by (kind, axis).
LINE_OPTIONS = {("insert", "row"): "insertRows", ("insert", "column"): "insertColumns",
                ("delete", "row"): "deleteRows", ("delete", "column"): "deleteColumns"}


def lines_problem(workbook: Any, sheet: Any, kind: str, axis: str, first: int,
                  last: int) -> str | None:
    """Why rows or columns ``first`` to ``last`` of the protected ``sheet`` are not
    inserted, deleted or moved as Excel would refuse it (see the module), with the way
    out; ``None`` when Excel would do it.

    Measured (Excel for Mac 16.106, 2026-11-09, through its object model): a sheet whose
    protection allows inserting rows (columns) took them, one whose protection does not
    refused them; a deletion it allows was done only when every cell of the rows (columns)
    was unlocked -- each cell's own format, else its row's (``customFormat``), else its
    column's, else locked: a row of unlocked cells under an unlocked row format, and one
    whose cells and every column were unlocked, were deleted; a row holding one locked cell,
    and a column whose positions below its cells had no unlocked format, were not.  Moving
    rows or columns (cut and insert) is not measured on a protected sheet: refused."""
    from .format import protection_allows
    from .structure import _label

    noun = f"{axis}s"
    if kind == "move":
        return (f"{noun} are not moved on a protected sheet: unprotect it first "
                "(wb.unprotect_sheet)")
    option = LINE_OPTIONS[(kind, axis)]
    if not protection_allows(workbook, sheet, option):
        return (f"its protection does not allow {kind}ing {noun} ({option}): unprotect the "
                f"sheet (wb.unprotect_sheet), or protect it allowing {option}")
    if kind != "delete":
        return None
    writer = workbook.editor.cells.sheet(sheet.stored_name)
    if writer is None:
        return None
    locked = _first_locked(workbook, sheet, writer, axis, first, last)
    if locked is None:
        return None
    return (f"{format_cell(_label(sheet), *locked)} is locked: Excel deletes {noun} of a "
            f"protected sheet only when every cell of them is unlocked (the cells' formats, "
            f"else the rows', else the columns')")


def _first_locked(workbook: Any, sheet: Any, writer: Any, axis: str, first: int,
                  last: int) -> tuple[int, int] | None:
    """The first locked position (row, column) of rows (``axis`` ``"row"``) or columns
    ``first`` to ``last``, by each position's format (see :func:`lines_problem`)."""
    from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW

    locks = Locks(workbook)
    root = workbook.package.tree(sheet.part)
    columns: list[tuple[int, int, bool]] = []     # (low, high, unlocked)
    node = root.find(_X + "cols")
    for column in (node.iterchildren(_X + "col") if node is not None else ()):
        low, high, style = column.get("min", ""), column.get("max", ""), column.get("style", "")
        if low.isdigit() and high.isdigit():
            columns.append((int(low), int(high),
                            not locks.style_locked(int(style) if style.isdigit() else 0)))
    columns.sort()

    def column_unlocked(number: int) -> bool:
        return any(low <= number <= high and open_ for low, high, open_ in columns)

    def first_locked_column(low: int, high: int) -> int | None:
        at = low
        for start, end, open_ in columns:
            if end < at or not open_:
                continue
            if start > at:
                return at
            at = end + 1
            if at > high:
                return None
        return at if at <= high else None

    data = root.find(_X + "sheetData")
    rows = {}
    for row in (data.iterchildren(_X + "row") if data is not None else ()):
        number = row.get("r", "")
        if number.isdigit():
            rows[int(number)] = row

    def row_format(row: Any) -> int | None:
        if row is not None and row.get("customFormat", "").strip().lower() in ("1", "true"):
            style = row.get("s", "")
            return int(style) if style.isdigit() else 0
        return None

    if axis == "row":
        for r in range(first, last + 1):
            row = rows.get(r)
            taken: set[int] = set()
            for cell in (_cells(row, r) if row is not None else ()):
                taken.add(cell.column)
                if locks.locked(writer, cell):
                    return (r, cell.column)
            style = row_format(row)
            if style is not None:
                if locks.style_locked(style) and len(taken) < MAX_COLUMN:
                    return (r, next(c for c in range(1, MAX_COLUMN + 1) if c not in taken))
                continue
            gap = first_locked_column(1, MAX_COLUMN)
            while gap is not None and gap in taken:
                gap = first_locked_column(gap + 1, MAX_COLUMN)
            if gap is not None:
                return (r, gap)
        return None
    for number, row in sorted(rows.items()):
        taken = set()
        for cell in _cells(row, number):
            if first <= cell.column <= last:
                taken.add(cell.column)
                if locks.locked(writer, cell):
                    return (number, cell.column)
        style = row_format(row)
        if style is not None and locks.style_locked(style):
            for c in range(first, last + 1):
                if c not in taken:
                    return (number, c)
    for c in range(first, last + 1):
        if not column_unlocked(c):
            # A position no row format covers, below or between the rows written.
            r = next(n for n in range(1, MAX_ROW + 1)
                     if row_format(rows.get(n)) is None
                     and not (n in rows and any(cell.column == c
                                                for cell in _cells(rows[n], n))))
            return (r, c)
    return None


__all__ += ["LINE_OPTIONS", "lines_problem"]
