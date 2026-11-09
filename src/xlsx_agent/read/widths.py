"""Column widths read back (production R76): each column's width as Excel's Column Width
shows it, and whether it is the column's own or the sheet's default.

::

    wb.column_widths("Stock!A:H")      # columns A to H
    wb.column_widths("Stock")          # every column up to the last one used or sized
    wb.column_widths("'Q3 sales'!C")   # one column; a range (Stock!B2:D9) reads its columns

Each column (:class:`ColumnWidth`) has its ``width`` in characters -- the unit
:meth:`Workbook.set_column_widths <xlsx_agent.Workbook.set_column_widths>` takes and Excel's
Column Width dialog shows -- with ``stored``, the width the file gives the column (``None``
when it takes the sheet's default; ``default`` is then true), and ``hidden``.  A hidden
column has the width it shows with once shown again; one stored as ``0`` is hidden (Excel
hides it and writes ``hidden``) and its width is 0.

**Measured** (Excel for Mac 16.106, 2026-11-09: each column's ``column width`` through its
object model, under the default fonts Aptos Narrow 11 and 12, Calibri 11 and 16):

* a stored width ``w`` shows as ``(round(w * digit) - 5) / digit`` characters, to two
  decimals (half up): ``12.83203125`` is 12 under a 6-pixel digit, ``9.140625`` 8.33, and
  under Calibri 16 (an 8-pixel digit) ``12.875`` is 12.25 and ``9.140625`` 8.5 -- ``digit``
  being the default font's widest digit in pixels (:func:`xlsx_agent.edit.layout.
  digit_width`; the 5 is Excel's padding, :data:`~xlsx_agent.edit.layout.PADDING`);
* a sheet's default: its ``sheetFormatPr defaultColWidth``, shown by the same rule
  (``12`` is 11.17 under a 6-pixel digit, 11.38 under an 8-pixel one; it wins over a
  ``baseColWidth``); without one, ``baseColWidth`` exactly (8 when the sheet format gives
  none, 10 and 4 as given) -- stored as :func:`~xlsx_agent.edit.layout.stored_width` writes
  it (Excel wrote ``defaultColWidth="8.83203125"`` for base 8 under a 6-pixel digit); a
  sheet with no ``sheetFormatPr`` at all showed 10 (:data:`~xlsx_agent.edit.layout.
  BASE_COLUMN_WIDTH_ABSENT`; Excel for Windows not measured);
* a hidden column's width reads 0 in Excel while it is hidden; its stored width is kept
  (``width="20" hidden="1"`` saved as it was) and is what it shows with again; a hidden
  column with no width shows with the default; a width of 0 is hidden.

A workbook whose default font has no measured widths has ``width`` ``None`` throughout (the
unit, its widest digit, is not known) and a note; ``stored`` is still given.  At most
``limit`` columns are listed (default :data:`DEFAULT_COLUMNS`, at most
:data:`MAX_COLUMNS`, the whole grid); the rest is counted in ``more`` and ``next`` is the
address that reads on.  Reading a sheet alone reads its used range (a pass over its
cells) to know its last column.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any

from ooxml_edit.sheets import MAX_COLUMN, column_letters

from ..xml import iterparse_xml

_X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

#: The columns one call lists unless asked for more, and the most it lists.
DEFAULT_COLUMNS = 256
MAX_COLUMNS = MAX_COLUMN


@dataclass(frozen=True)
class ColumnWidth:
    """One column's width: ``width`` in characters as Excel shows it (``None`` when the
    default font is not measured), ``stored`` the file's width for it in the file's unit
    (``None``: the sheet's default, and ``default`` is true), ``hidden``."""

    column: str
    width: float | None
    stored: float | None
    default: bool
    hidden: bool = False

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"column": self.column, "width": self.width,
                               "source": "default" if self.default else "column"}
        if self.stored is not None:
            out["stored"] = self.stored
        if self.hidden:
            out["hidden"] = True
        return out


@dataclass(frozen=True)
class ColumnWidths:
    """What :meth:`Workbook.column_widths <xlsx_agent.Workbook.column_widths>` read: the
    sheet, its columns in order, the sheet's default width (``default_width`` in
    characters, ``default_stored`` in the file's unit; each ``None`` when the default
    font is not measured and the sheet gives no ``defaultColWidth``), the default font's widest digit in
    pixels (``digit``; ``None`` when not measured), ``more`` columns not listed and the
    address ``next`` that reads them, and ``notes``."""

    sheet: str
    columns: tuple[ColumnWidth, ...]
    default_width: float | None
    default_stored: float | None
    digit: int | None
    more: int = 0
    next: str | None = None
    notes: tuple[str, ...] = field(default_factory=tuple)

    def __getitem__(self, column: str) -> ColumnWidth:
        """The column of these letters (``"C"``)."""
        wanted = column.upper()
        for item in self.columns:
            if item.column == wanted:
                return item
        raise KeyError(f"column {column!r} is not among {self.columns[0].column}:"
                       f"{self.columns[-1].column}" if self.columns else "no columns read")

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"sheet": self.sheet, "default_width": self.default_width,
                               "default_stored": self.default_stored,
                               "columns": [c.as_dict() for c in self.columns]}
        if self.more:
            out["more"] = self.more
            out["next"] = self.next
        if self.notes:
            out["notes"] = list(self.notes)
        return out


def shown_width(stored: float, digit: int) -> float:
    """A stored width as Excel's Column Width shows it, in characters (see the module)."""
    from ..edit.layout import PADDING, _round

    pixels = _round(stored * digit)
    return max(0.0, math.floor((pixels - PADDING) / digit * 100 + 0.5) / 100)


def sheet_head(workbook: Any, part: str) -> tuple[dict[str, str] | None,
                                                  list[tuple[int, int, dict[str, str]]]]:
    """A worksheet's ``sheetFormatPr`` attributes (``None`` without one) and its ``col``
    runs (first, last, attributes), from the tree an edit holds, or read from the part's
    start up to its cells (never the cells)."""
    trees = getattr(workbook.package, "_trees", None)
    root = trees.get(part) if isinstance(trees, dict) else None
    fmt: dict[str, str] | None = None
    runs: list[tuple[int, int, dict[str, str]]] = []

    def col(node: Any) -> None:
        low, high = node.get("min", ""), node.get("max", "")
        if low.isdigit() and high.isdigit():
            runs.append((int(low), int(high), {k: v for k, v in node.attrib.items()
                                               if k not in ("min", "max")}))

    if root is not None:
        node = root.find(_X + "sheetFormatPr")
        fmt = dict(node.attrib) if node is not None else None
        cols = root.find(_X + "cols")
        for item in cols.iterchildren(_X + "col") if cols is not None else ():
            col(item)
        return fmt, runs
    data = workbook.package.read(part) or b""
    for event, node in iterparse_xml(data, events=("start", "end")):
        if event == "start":
            if node.tag == _X + "sheetData":
                break
            continue
        if node.tag == _X + "sheetFormatPr":
            fmt = dict(node.attrib)
        elif node.tag == _X + "col":
            col(node)
    return fmt, runs


def column_widths(workbook: Any, target: Any, *, limit: int = DEFAULT_COLUMNS) -> ColumnWidths:
    """The widths of the columns ``target`` names (see the module)."""
    from ..address import format_sheet
    from ..edit import EditError
    from ..edit.format import Problems, targets
    from ..edit.layout import (BASE_COLUMN_WIDTH, BASE_COLUMN_WIDTH_ABSENT, MEASURED_FONTS,
                               _quoted, digit_known, digit_width, face_from_styles,
                               stored_width)

    if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= MAX_COLUMNS:
        raise EditError("bad_value", f"limit is a number of columns, 1 to {MAX_COLUMNS:,}, "
                        f"not {limit!r}")
    if not isinstance(target, str) or not target:
        raise EditError("bad_target", "columns with their sheet, as Stock!A:H or Stock!C, "
                        "or a sheet's name, as Stock")
    address = target if "!" in target else _quoted(target) + "!"
    head, bang, tail = address.rpartition("!")
    letters = tail.strip().replace("$", "")
    if bang and 1 <= len(letters) <= 3 and letters.isascii() and letters.isalpha():
        address = f"{head}!{letters}:{letters}"      # one column, as Stock!C (C:C)
    problems = Problems()
    found = targets(workbook, address, problems, spans=("cells", "columns", "sheet"))
    problems.raise_any()
    item = found[0]
    sheet = item.sheet
    fmt, runs = sheet_head(workbook, sheet.part)
    styles = workbook._related("styles")
    face = face_from_styles(workbook.package.tree(styles) if styles else None)
    digit = digit_width(face) if digit_known(face) else None
    notes: list[str] = []
    if digit is None:
        notes.append(f"the workbook's default font, {face.name}, has no measured widths: "
                     "its widest digit, the unit of a width in characters, is not known, so "
                     "only stored widths are given; fonts that are measured: "
                     + ", ".join(MEASURED_FONTS))
    base = BASE_COLUMN_WIDTH if fmt is not None else BASE_COLUMN_WIDTH_ABSENT
    raw_base = (fmt or {}).get("baseColWidth") or ""
    if raw_base.isdigit():
        base = int(raw_base)
    default_stored: float | None = None
    raw_default = (fmt or {}).get("defaultColWidth")
    if raw_default:
        try:
            default_stored = float(raw_default)
        except ValueError:
            default_stored = None
        if default_stored is not None and not (math.isfinite(default_stored)
                                               and default_stored >= 0):
            default_stored = None
    if default_stored is not None:
        default_width = shown_width(default_stored, digit) if digit else None
    else:
        default_width = float(base) if digit else None
        default_stored = stored_width(base, digit) if digit else None
    first, last = item.area.first.column, item.area.last.column
    if item.span == "sheet":
        first, last = 1, _last_column(workbook, sheet, runs)
    columns: list[ColumnWidth] = []
    for number in range(first, min(last, first + limit - 1) + 1):
        attrs: dict[str, str] = {}
        for low, high, run in runs:
            if low <= number <= high:
                attrs = run
                break
        stored = _width(attrs.get("width"))
        hidden = (attrs.get("hidden") or "").strip().lower() in ("1", "true") \
            or stored == 0
        if stored is None:
            width = default_width
        else:
            width = shown_width(stored, digit) if digit else None
        columns.append(ColumnWidth(column_letters(number), width, stored, stored is None,
                                   hidden))
    more = max(0, last - first + 1 - limit)
    label = format_sheet(sheet.name, sheet.id)
    after = None
    if more:
        start = first + limit
        after = f"{label}!{column_letters(start)}:{column_letters(last)}"
        notes.append(f"{more:,} more column(s) not listed (limit={limit}): read them with "
                     f"column_widths({after!r}) or a larger limit, up to {MAX_COLUMNS:,}")
    return ColumnWidths(sheet.name, tuple(columns), default_width, default_stored, digit,
                        more, after, tuple(notes))


def _width(text: str | None) -> float | None:
    if text is None:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) and value >= 0 else None


def _last_column(workbook: Any, sheet: Any, runs: list[tuple[int, int, dict[str, str]]]) -> int:
    """A sheet's last column worth reading: the used range's, or the last one given a
    width or hidden (a run reaching the grid's end counts as its first column)."""
    used = workbook.used_range(sheet)
    last = used.last.column if used is not None else 1
    for low, high, attrs in runs:
        if "width" in attrs or "hidden" in attrs:
            last = max(last, low if high >= MAX_COLUMN else high)
    return last
