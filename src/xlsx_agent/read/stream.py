"""The streaming pass over a sheet: every cell typed, the sheet never parsed whole.

A worksheet may hold a million rows (ROADMAP.md, "Design rules", 6), so it is read with
:func:`xlsx_agent.xml.iterparse_xml` (``parse_xml``'s safe options): each cell is typed as
its row ends, and the row is then dropped, so what is held at any moment is one row.  What the pass meets on the way -- hidden columns (``cols`` comes before the
cells), hidden rows, merges (``mergeCells`` comes after them), how the rows are ordered --
is collected as it goes.

Positions are computed as Excel computes them (ROADMAP.md, "Messy workbooks"): a row
without ``r`` is the row after the one before it, a cell without ``r`` the column after
the one before it in its row (the first column in a row's first cell).  A ``dimension``
is never trusted: :class:`SheetScan` computes the used range from the cells.

How a cell's value is typed is :mod:`xlsx_agent.read.types`'s; the cell types are
ECMA-376 Part 1, §18.18.11 (ST_CellType): ``b``, ``d``, ``e``, ``inlineStr``, ``n``,
``s``, ``str``.
"""

from __future__ import annotations

import math
import re
from bisect import bisect_right
from dataclasses import dataclass, field
from typing import Iterator

from ooxml_edit.charts.workbook import MAX_COLUMN, MAX_ROW, Area, Cell

from ..xml import iterparse_xml
from .strings import SharedStrings, item_text
from .styles import DATE_KINDS, Styles
from .types import (TEXT_LIMIT, Locale, Problem, Value, looks_like_date, looks_like_number,
                    serial_text)

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")


def _tags(local: str) -> frozenset[str]:
    return frozenset(f"{{{ns}}}{local}" for ns in NS_MAIN)


_ROW, _C, _COL, _MERGE = _tags("row"), _tags("c"), _tags("col"), _tags("mergeCell")
_DIMENSION = _tags("dimension")
_V, _F, _IS = _tags("v"), _tags("f"), _tags("is")
_STREAMED = tuple(_ROW | _COL | _MERGE | _DIMENSION)

_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})")
_INTEGER = re.compile(r"[+-]?[0-9]+")
_TIME = r"\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?"
_ISO_DATETIME = re.compile(rf"(?:(\d{{4}}-\d{{2}}-\d{{2}})(?:T({_TIME}))?|({_TIME}))"
                           r"(?:Z|[+-]\d{2}:\d{2})?")

#: The functions whose value changes at every recalculation (ROADMAP.md, "Values and
#: types"): a cached value of a formula that calls one is as of the last save.
VOLATILE_FUNCTIONS = ("NOW", "TODAY", "RAND", "OFFSET", "INDIRECT")
_VOLATILE = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?(?:" + "|".join(VOLATILE_FUNCTIONS)
                       + r")\s*\(", re.IGNORECASE)
_QUOTED = re.compile(r'"[^"]*"|\'[^\']*\'')

_COLUMNS: dict[str, int] = {}
_MISSING = object()
_CACHE_LIMIT = 1024


def _column(letters: str) -> int:
    index = _COLUMNS.get(letters)
    if index is None:
        index = 0
        for char in letters.upper():
            index = index * 26 + ord(char) - 64
        _COLUMNS[letters] = index
    return index


def is_volatile(formula: str) -> bool:
    """Whether a formula's text calls one of :data:`VOLATILE_FUNCTIONS` (outside its
    string literals and quoted sheet names)."""
    return bool(_VOLATILE.search(_QUOTED.sub("", formula)))


@dataclass
class SheetScan:
    """What one full pass over a sheet found.

    ``used`` is the used range computed from the cells that hold a value or a formula
    (``None`` for a sheet with none; a cell with only a style is blank and does not count);
    ``dimension`` the ``dimension`` as written, never trusted; ``cells`` how many cells
    hold a value or a formula; ``merges`` the merged areas; ``hidden_rows`` how many rows
    are hidden; ``ordered`` whether rows and cells come in the order of the sheet;
    ``issues`` each value issue (``no_cached_value``, ``number_as_text``...), ``at_limit``
    and ``volatile`` by code, as ``(count, first address)``."""

    sheet: str
    used: Area | None = None
    dimension: str | None = None
    cells: int = 0
    merges: list[Area] = field(default_factory=list)
    hidden_rows: int = 0
    ordered: bool = True
    issues: dict[str, tuple[int, str]] = field(default_factory=dict)

    @property
    def problems(self) -> list[Problem]:
        """The issues as :class:`~xlsx_agent.read.types.Problem`, by code."""
        return [Problem(code, address, count) for code, (count, address) in sorted(self.issues.items())]

    def _count(self, code: str, value: Value) -> None:
        seen = self.issues.get(code)
        self.issues[code] = (1, value.a) if seen is None else (seen[0] + 1, seen[1])


class SheetReader:
    """One sheet's cells, streamed from the part's bytes.

    ``sheet`` is the sheet as an address writes it; ``hidden`` whether the sheet is hidden
    (every value from it is then marked ``hidden``); ``strings``, ``styles``, ``date1904``
    and ``locale`` are the workbook's.  ``formats`` puts each value's format code on it.
    A reader makes one pass at a time; each pass starts from the top."""

    def __init__(self, data: bytes, *, sheet: str, hidden: bool, strings: SharedStrings,
                 styles: Styles, date1904: bool, locale: Locale, formats: bool = False) -> None:
        self.data = data
        self.sheet = sheet
        self.hidden = hidden
        self.strings = strings
        self.styles = styles
        self.date1904 = date1904
        self.locale = locale
        self.formats = formats
        #: Hidden columns as sorted, non-overlapping ``(first, last)`` spans, known once
        #: the pass has passed ``cols``.
        self.hidden_columns: list[tuple[int, int]] = []
        #: Hidden rows met so far.
        self.hidden_rows: set[int] = set()
        self.merges: list[Area] = []
        self.dimension: str | None = None
        self.ordered = True
        self._shared: dict[str, tuple[str, bool]] = {}
        self._format_cache: dict[str | None, object] = {}
        # Each shared string's text and what it looks like, once per pass (as many as the
        # table, which is held already); dates by serial, at most _CACHE_LIMIT at a time.
        self._string_cache: dict[int, tuple[str, str | None, bool]] = {}
        self._date_cache: dict[tuple, str | None] = {}

    # -- the pass ----------------------------------------------------------------------

    def values(self) -> Iterator[Value]:
        """Every cell in the sheet, typed, in the order of the file (a cell with only a
        style comes as ``blank``)."""
        value = self.value
        for row, column, row_hidden, node in self._walk():
            yield value(row, column, row_hidden, node)

    def _walk(self):
        """``(row, column, row hidden, c element)`` for each cell.  The pass asks for the
        ``end`` of each row (and of ``col``, ``mergeCell``, ``dimension``), takes the row's
        cells, then drops the row: one row is held at a time."""
        if not self.data:
            return
        row = 0
        last_row = last_column = 0
        hidden_rows = self.hidden_rows
        hidden_spans: list[tuple[int, int]] = []
        match_ref = _REF.fullmatch
        columns = _COLUMNS
        cell_tags, row_tags = _C, _ROW
        for _, node in iterparse_xml(self.data, tag=_STREAMED):
            tag = node.tag
            if tag in row_tags:
                number = _int(node.get("r"))
                row = number if number is not None and 1 <= number <= MAX_ROW else row + 1
                row_hidden = _true(node.get("hidden"))
                if row_hidden:
                    hidden_rows.add(row)
                column = 0
                for cell in node:
                    if cell.tag not in cell_tags:
                        continue
                    ref = cell.get("r")
                    match = match_ref(ref) if ref else None
                    if match is not None:
                        letters = match.group(1)
                        column = columns.get(letters) or _column(letters)
                        cell_row = int(match.group(2))
                        cell_hidden = row_hidden if cell_row == row else cell_row in hidden_rows
                    else:
                        cell_row, column, cell_hidden = row, column + 1, row_hidden
                    if not (1 <= cell_row <= MAX_ROW and 1 <= column <= MAX_COLUMN):
                        continue
                    if cell_row < last_row or (cell_row == last_row and column <= last_column):
                        self.ordered = False
                    last_row, last_column = cell_row, column
                    yield cell_row, column, cell_hidden, cell
                node.clear()
                parent = node.getparent()
                if parent is not None:
                    while node.getprevious() is not None:
                        del parent[0]
            elif tag in _COL:
                if _true(node.get("hidden")):
                    first, end = _int(node.get("min")), _int(node.get("max"))
                    if first is not None and end is not None and 1 <= first <= end:
                        hidden_spans.append((first, min(end, MAX_COLUMN)))
                        self.hidden_columns = _merge_spans(hidden_spans)
            elif tag in _MERGE:
                area = _area(node.get("ref"))
                if area is not None:
                    self.merges.append(area)
            elif tag in _DIMENSION:
                self.dimension = node.get("ref")

    def column_hidden(self, column: int) -> bool:
        spans = self.hidden_columns
        if not spans:
            return False
        k = bisect_right(spans, (column, MAX_COLUMN + 1)) - 1
        return k >= 0 and spans[k][0] <= column <= spans[k][1]

    def _format(self, style: str | None):
        number_format = self._format_cache.get(style)
        if number_format is None:
            number_format = self._format_cache[style] = self.styles.format(_int(style) or 0)
        return number_format

    # -- one cell ----------------------------------------------------------------------

    def value(self, row: int, column: int, row_hidden: bool, node) -> Value:
        """The typed value of a ``c`` element at ``row``, ``column``."""
        hidden = self.hidden or row_hidden or (bool(self.hidden_columns) and self.column_hidden(column))
        value = Value(self.sheet, row, column, hidden=hidden)
        kind = node.get("t") or "n"
        v_text = f_node = is_node = None
        for child in node:
            tag = child.tag
            if tag in _V:
                v_text = child.text or ""
            elif tag in _F:
                f_node = child
            elif tag in _IS:
                is_node = child
        if f_node is not None:
            self._formula(value, f_node)
        if self.formats:
            value.format = self._format(node.get("s")).code

        if kind == "n":
            self._number(value, v_text, node)
        elif kind == "s":
            self._shared_string(value, v_text)
        elif kind == "inlineStr":
            if is_node is not None:
                self._text(value, item_text(is_node))
            elif v_text is not None:
                self._text(value, v_text)
        elif kind == "str":
            if v_text is not None:
                self._text(value, v_text)
        elif kind == "b":
            flag = (v_text or "").strip()
            if flag in ("1", "true"):
                value.v, value.t = True, "bool"
            elif flag in ("0", "false"):
                value.v, value.t = False, "bool"
            elif v_text is not None and not value.formula:
                value.issue = "bad_value"
        elif kind == "e":
            if v_text is not None and v_text.strip():
                value.v, value.t = v_text.strip(), "error"
        elif kind == "d":
            self._iso(value, v_text)
        else:  # not an ST_CellType: read as a number, as the default type is
            self._number(value, v_text, node)
        if value.formula and value.t == "blank" and value.issue is None:
            value.issue = "no_cached_value"
        return value

    def _shared_string(self, value: Value, v_text: str | None) -> None:
        index = _int(v_text)
        cached = self._string_cache.get(index) if index is not None else None
        if cached is None:
            try:
                if index is None or index < 0:
                    raise IndexError
                text = self.strings[index]
            except IndexError:
                value.issue = "missing_string"
                return
            cached = self._string_cache[index] = (text, self._looks(text), len(text) >= TEXT_LIMIT)
        value.v, value.t, value.at_limit = cached[0], "text", cached[2]
        if not value.formula:
            value.issue = cached[1]

    def _looks(self, text: str) -> str | None:
        """``"number_as_text"`` or ``"date_as_text"`` when text looks like one, in the
        locale."""
        if looks_like_number(text, self.locale):
            return "number_as_text"
        if looks_like_date(text, self.locale):
            return "date_as_text"
        return None

    def _text(self, value: Value, text: str) -> None:
        value.v, value.t = text, "text"
        if len(text) >= TEXT_LIMIT:
            value.at_limit = True
        if not value.formula:
            value.issue = self._looks(text)

    def _number(self, value: Value, v_text: str | None, node) -> None:
        if not v_text:
            return
        text = v_text.strip()
        try:
            if not text.isascii():
                raise ValueError
            number = int(text) if _INTEGER.fullmatch(text) else float(text)
        except ValueError:
            value.issue = "bad_value"
            return
        if isinstance(number, float) and not math.isfinite(number):
            value.issue = "bad_value"
            return
        value.v, value.t = number, "number"
        number_format = self._format(node.get("s"))
        kind = number_format.kind
        if kind == "general" or kind == "number" or kind == "text":
            return
        if kind == "percent":
            value.t = "percent"
        elif kind in DATE_KINDS:
            key = (number, kind, number_format.elapsed)
            iso = self._date_cache.get(key, _MISSING)
            if iso is _MISSING:
                if len(self._date_cache) >= _CACHE_LIMIT:
                    self._date_cache.clear()
                iso = self._date_cache[key] = serial_text(float(number), kind, self.date1904,
                                                          number_format.elapsed)
            if iso is None:
                value.issue = "date_out_of_range"
            else:
                value.v, value.t, value.serial = iso, kind, float(number)
        elif kind == "ambiguous":
            value.issue = "ambiguous_format"
        elif kind == "unknown":
            value.issue = "unknown_format"

    def _iso(self, value: Value, v_text: str | None) -> None:
        """A ``t="d"`` cell: an ISO 8601 date, time or both (§18.17.4.1), kept as given."""
        if v_text is None or not v_text.strip():
            return
        text = v_text.strip()
        match = _ISO_DATETIME.fullmatch(text)
        if match is None:
            value.issue = "bad_value"
            return
        value.v = text
        value.t = "time" if match.group(3) else "datetime" if match.group(2) else "date"

    def _formula(self, value: Value, f_node) -> None:
        value.formula = True
        text = f_node.text
        kind = f_node.get("t")
        if kind == "shared":
            group = f_node.get("si")
            if text:
                value.formula_text = text
                value.volatile = is_volatile(text)
                if group is not None:
                    self._shared[group] = (value.a, value.volatile)
            elif group in self._shared:
                value.shared_from, value.volatile = self._shared[group]
            return
        if kind == "array":
            value.array = f_node.get("ref") or value.ref
        if text:
            value.formula_text = text
            value.volatile = is_volatile(text)


def scan(reader: SheetReader) -> SheetScan:
    """One full pass over ``reader``'s sheet: the used range, merges, ordering and the
    issues met (see :class:`SheetScan`)."""
    result = SheetScan(reader.sheet)
    top = left = MAX_ROW + MAX_COLUMN
    bottom = right = 0
    for value in reader.values():
        if value.t == "blank" and not value.formula:
            if value.issue is not None:
                result._count(value.issue, value)
            continue
        result.cells += 1
        top, bottom = min(top, value.row), max(bottom, value.row)
        left, right = min(left, value.column), max(right, value.column)
        if value.issue is not None:
            result._count(value.issue, value)
        if value.at_limit:
            result._count("at_limit", value)
        if value.volatile:
            result._count("volatile", value)
    if result.cells:
        result.used = Area(Cell(top, left), Cell(bottom, right))
    result.dimension = reader.dimension
    result.merges = list(reader.merges)
    result.hidden_rows = len(reader.hidden_rows)
    result.ordered = reader.ordered
    return result


def read_area(reader: SheetReader, area: Area, ordered: bool) -> Iterator[list[Value]]:
    """The rows of ``area``, each a list of values from its first column to its last,
    blanks filled in.  With ``ordered`` (the sheet's rows and cells in order, as a scan
    found) the pass stops after the area's last row; otherwise the area's cells are
    gathered first."""
    top, left = area.first.row, area.first.column
    bottom, right = area.last.row, area.last.column
    width = right - left + 1
    pending: dict[int, dict[int, Value]] = {}
    if not ordered:
        for value in reader.values():
            if top <= value.row <= bottom and left <= value.column <= right:
                pending.setdefault(value.row, {})[value.column] = value
        for row in range(top, bottom + 1):
            yield _row(reader, row, left, width, pending.pop(row, None))
        return
    next_row = top
    current: dict[int, Value] | None = None
    current_row = None
    for row, column, row_hidden, node in reader._walk():
        if row < top:
            continue
        if row != current_row:
            if current_row is not None:
                yield _row(reader, current_row, left, width, current)
                next_row = current_row + 1
            if row > bottom:
                current_row = None
                break
            while next_row < row:
                yield _row(reader, next_row, left, width, None)
                next_row += 1
            current_row, current = row, {}
        if left <= column <= right:
            current[column] = reader.value(row, column, row_hidden, node)
    if current_row is not None:
        yield _row(reader, current_row, left, width, current)
        next_row = current_row + 1
    while next_row <= bottom:
        yield _row(reader, next_row, left, width, None)
        next_row += 1


def _row(reader: SheetReader, row: int, left: int, width: int,
         cells: dict[int, Value] | None) -> list[Value]:
    out = []
    row_hidden = reader.hidden or row in reader.hidden_rows
    for column in range(left, left + width):
        value = cells.get(column) if cells else None
        if value is None:
            value = Value(reader.sheet, row, column,
                          hidden=row_hidden or reader.column_hidden(column))
        out.append(value)
    return out


def _merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for first, last in sorted(spans):
        if out and first <= out[-1][1] + 1:
            out[-1] = (out[-1][0], max(out[-1][1], last))
        else:
            out.append((first, last))
    return out


def _area(ref: str | None) -> Area | None:
    if not ref:
        return None
    first, _, last = ref.partition(":")
    a, b = _REF.fullmatch(first), _REF.fullmatch(last or first)
    if a is None or b is None:
        return None
    one = Cell(int(a.group(2)), _column(a.group(1)))
    two = Cell(int(b.group(2)), _column(b.group(1)))
    if not (1 <= min(one.row, two.row) and max(one.row, two.row) <= MAX_ROW
            and 1 <= min(one.column, two.column) and max(one.column, two.column) <= MAX_COLUMN):
        return None
    return Area(Cell(min(one.row, two.row), min(one.column, two.column)),
                Cell(max(one.row, two.row), max(one.column, two.column)))


def _true(value: str | None) -> bool:
    return value is not None and value.strip() in ("1", "true")


def _int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None
