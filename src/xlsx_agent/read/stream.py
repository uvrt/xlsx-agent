"""The streaming pass over a sheet: every cell typed, the sheet never parsed whole.

A worksheet may hold a million rows (ROADMAP.md, "Design rules", 6), so it is read with
:func:`xlsx_agent.xml.iterparse_xml` (``parse_xml``'s safe options): each cell is typed as
its row ends, and the row is then dropped, so what is held at any moment is one row.  What
the pass meets on the way -- hidden columns (``cols`` comes before the cells), hidden rows,
array formulas' blocks, merges, data-validation rules and hyperlinks (the last three come
after the cells), how the rows are ordered -- is collected as it goes.

**Hyperlinks** (ECMA-376 Part 1, §18.3.1.47 ``hyperlink``): the cells each covers
(``ref``), its relationship id (``r:id``, an external target in the sheet's relationships)
and its ``location`` (a place in the workbook), as written.  Content: returned by a read
when asked for, never in a message.

**Data validation** (ECMA-376 Part 1, §18.3.1.32 ``dataValidation``): each rule's type
(ST_DataValidationType), the cells it covers (``sqref``) and its first formula, as text,
for a program only (formulas hold literals: never in a message).  A rule Excel keeps in
the worksheet's ``extLst`` (the ``x14:dataValidation`` of Microsoft's [MS-XLSX]
extensions, whose ``xm:sqref`` and ``x14:formula1/xm:f`` carry the same) is read alike.

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
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, field
from typing import Any, Iterator

from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell

from ..formula import is_external
from ..xml import iterparse_xml
from .strings import SharedStrings, item_text
from .styles import DATE_KINDS, Styles
from .types import (TEXT_LIMIT, Problem, Value, _letters, looks_like_date, looks_like_number,
                    serial_text)

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")


def _tags(local: str) -> frozenset[str]:
    return frozenset(f"{{{ns}}}{local}" for ns in NS_MAIN)


_ROW, _C, _COL, _MERGE = _tags("row"), _tags("c"), _tags("col"), _tags("mergeCell")
_DIMENSION, _HYPERLINK = _tags("dimension"), _tags("hyperlink")
#: ``r:id``, transitional and strict.
_R_ID = tuple(f"{{{ns}}}id" for ns in (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships"))
_VALIDATION, _FORMULA1 = _tags("dataValidation"), _tags("formula1")
#: [MS-XLSX]'s namespaces of the x14 data-validation extension and of its ``xm:`` children.
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"
_X14_VALIDATION = frozenset({f"{{{NS_X14}}}dataValidation"})
_X14_FORMULA1, _XM_F, _XM_SQREF = f"{{{NS_X14}}}formula1", f"{{{NS_XM}}}f", f"{{{NS_XM}}}sqref"
_V, _F, _IS = _tags("v"), _tags("f"), _tags("is")
_STREAMED = tuple(_ROW | _COL | _MERGE | _DIMENSION | _VALIDATION | _X14_VALIDATION
                  | _HYPERLINK)

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

#: A sheet's :class:`RowIndex` notes where every this many rows start.
CHECKPOINT_ROWS = 128
#: A ``row`` start tag and the ``sheetData`` start tag, as bytes, with any prefix.
_ROW_TAG = re.compile(rb"<(?:[A-Za-z_][-\w.]*:)?row[\s/>]")
_SHEET_DATA_TAG = re.compile(rb"<(?:[A-Za-z_][-\w.]*:)?sheetData(?:\s[^>]*)?>")


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


@dataclass(frozen=True)
class Validation:
    """One data-validation rule: ``type`` its ST_DataValidationType as written (``"list"``,
    ``"whole"``, ``"decimal"``...; ``"none"`` when the file gives none), ``areas`` the
    cells it covers (``sqref``; a reference that is not one is left out), ``formula1`` its
    first formula as text (a list's ``"Red,Green"`` literal, a range or a name), for a
    program only."""

    type: str
    areas: tuple[Area, ...]
    formula1: str | None = field(default=None, repr=False)


@dataclass(frozen=True)
class Hyperlink:
    """One hyperlink as the sheet writes it: ``area`` the cells it covers, ``rel_id`` its
    relationship (an external target), ``location`` a place in the workbook (``Sheet2!A1``,
    a defined name), and ``url`` the relationship's target, which
    :meth:`Workbook.hyperlinks <xlsx_agent.workbook.Workbook.hyperlinks>` fills in; each
    ``None`` when not given.  Content, for a result only: kept off the repr."""

    area: Area
    rel_id: str | None = field(default=None, repr=False)
    location: str | None = field(default=None, repr=False)
    url: str | None = field(default=None, repr=False)


@dataclass
class SheetScan:
    """What one full pass over a sheet found.

    ``used`` is the used range computed from the cells that hold a value or a formula
    (``None`` for a sheet with none; a cell with only a style is blank and does not count);
    ``dimension`` the ``dimension`` as written, never trusted; ``cells`` how many cells
    hold a value or a formula; ``merges`` the merged areas; ``hidden_rows`` how many rows
    are hidden; ``ordered`` whether rows and cells come in the order of the sheet;
    ``issues`` each value issue (``no_cached_value``, ``number_as_text``...), ``at_limit``,
    ``volatile`` and ``external_formula`` (a formula that refers to another workbook) by
    code, as ``(count, first address)``; ``validations`` the sheet's
    data-validation rules, in the file's order; ``arrays`` the blocks the array formulas
    (and dynamic arrays' spills) fill, each its formula's ``ref``, in reading order;
    ``hyperlinks`` the sheet's hyperlinks, in the file's order."""

    sheet: str
    used: Area | None = None
    dimension: str | None = None
    cells: int = 0
    merges: list[Area] = field(default_factory=list)
    hidden_rows: int = 0
    ordered: bool = True
    issues: dict[str, tuple[int, str]] = field(default_factory=dict)
    validations: list[Validation] = field(default_factory=list)
    arrays: list[Area] = field(default_factory=list)
    hyperlinks: list[Hyperlink] = field(default_factory=list)
    #: Where the rows start, for a later read to start near them (an ordered sheet's;
    #: ``None`` otherwise).
    index: "RowIndex | None" = field(default=None, repr=False, compare=False)

    @property
    def problems(self) -> list[Problem]:
        """The issues as :class:`~xlsx_agent.read.types.Problem`, by code."""
        return [Problem(code, address, count) for code, (count, address) in sorted(self.issues.items())]

    def _count(self, code: str, value: Value) -> None:
        seen = self.issues.get(code)
        self.issues[code] = (1, value.a) if seen is None else (seen[0] + 1, seen[1])


@dataclass(frozen=True)
class Checkpoint:
    """Where a row element starts in a sheet part (``offset``, in bytes) and what a pass
    from the top knows there: the row number before it (``row``: a row without ``r`` is
    the next), the last cell's row and column, and how many hidden rows it has met."""

    offset: int
    row: int
    last_row: int
    last_column: int
    hidden: int


class RowIndex:
    """Where a sheet's rows start: a :class:`Checkpoint` every :data:`CHECKPOINT_ROWS`
    rows, so that a read of later rows starts at the last checkpoint before them rather
    than at the top (an ordered sheet only: its cells come in reading order, so none after
    a checkpoint is above it).  The part is parsed from there as its ``header`` (every byte
    up to and including the ``sheetData`` start tag: the namespaces, ``cols``,
    ``dimension``) followed by the bytes from the checkpoint on: the same document without
    the rows before.  Made by :func:`scan` from the marks its pass noted; ``None`` there
    for a part whose bytes it cannot match to the rows the parser met."""

    def __init__(self, header: bytes, checkpoints: list[Checkpoint], hidden: list[int],
                 shared_formulas: bool = False) -> None:
        self.header = header
        self.checkpoints = checkpoints
        #: The hidden rows in the order the pass met them.
        self.hidden = hidden
        #: Whether the sheet has a shared formula group: a read through rows notes each
        #: group's first cell for the cells after it, which a read started below misses.
        self.shared_formulas = shared_formulas
        self._last_rows = [c.last_row for c in checkpoints]

    def before(self, row: int) -> Checkpoint | None:
        """The last checkpoint with every cell before it above ``row``; ``None`` when the
        first is not (start at the top)."""
        k = bisect_left(self._last_rows, row) - 1
        return self.checkpoints[k] if k > 0 else None

    @classmethod
    def build(cls, data: bytes, reader: "SheetReader") -> "RowIndex | None":
        """The index of a part from what a whole pass over it noted (``reader.marks``):
        each checkpoint's row element found by its start tag in the bytes, the
        ``n``-th tag being the ``n``-th row the parser met.  ``None`` when that cannot be
        relied on: a comment, CDATA section, document type or processing instruction
        (whose text may hold a tag), ``sheetData`` written other than once as a start tag
        and once as an end tag, or as many tags as rows not found."""
        marks = reader.marks
        if not marks or len(marks) < 2 or b"<!" in data or data.count(b"<?") > 1 \
                or data.find(b"<?") not in (-1, 0, 3) or data.count(b"sheetData") != 2:
            return None
        found = _SHEET_DATA_TAG.search(data)
        if found is None:
            return None
        wanted = {mark[0]: mark for mark in marks}
        checkpoints = []
        count = 0
        for match in _ROW_TAG.finditer(data, found.end()):
            mark = wanted.get(count)
            if mark is not None:
                checkpoints.append(Checkpoint(match.start(), *mark[1:]))
            count += 1
        if count != reader.row_elements or len(checkpoints) != len(marks):
            return None
        return cls(data[:found.end()], checkpoints, list(reader.hidden_order),
                   bool(reader._shared))


class _Resumed:
    """A part's bytes as a binary file read from a checkpoint: ``header``, then ``data``
    from ``offset`` on, read in pieces (the part is never copied whole)."""

    def __init__(self, header: bytes, data: bytes, offset: int) -> None:
        self._parts = [memoryview(header), memoryview(data)[offset:]]
        self._part = self._at = 0

    def read(self, size: int = -1) -> bytes:
        out = []
        left = size if size is not None and size >= 0 else None
        while self._part < len(self._parts) and (left is None or left > 0):
            part = self._parts[self._part]
            end = len(part) if left is None else min(len(part), self._at + left)
            out.append(part[self._at:end])
            if left is not None:
                left -= end - self._at
            self._at = end
            if end == len(part):
                self._part, self._at = self._part + 1, 0
        return b"".join(out)


class SheetReader:
    """One sheet's cells, streamed from the part's bytes.

    ``sheet`` is the sheet as an address writes it; ``hidden`` whether the sheet is hidden
    (every value from it is then marked ``hidden``); ``strings``, ``styles``, ``date1904``
    and ``locale`` (with its reader settings ``grouping`` and ``local_dates``,
    :mod:`xlsx_agent.read.types`) are the workbook's.  ``formats`` puts each value's format
    code on it.
    A reader makes one pass at a time; each pass starts from the top."""

    def __init__(self, data: bytes, *, sheet: str, hidden: bool, strings: SharedStrings,
                 styles: Styles, date1904: bool, locale: Any, formats: bool = False,
                 grouping: bool = True, local_dates: bool = False,
                 redaction: Any = None, cache: str = "cached_unverified",
                 external_names: "frozenset[str]" = frozenset()) -> None:
        self.data = data
        self.sheet = sheet
        self.hidden = hidden
        self.strings = strings
        self.styles = styles
        self.date1904 = date1904
        self.locale = locale
        self.grouping = grouping
        self.local_dates = local_dates
        self.formats = formats
        #: The workbook's redaction (:mod:`xlsx_agent.redact`): applied to every text (the
        #: shared strings come redacted), a number format's code, and a formula's text,
        #: which is withheld when it holds a match.
        self.redaction = redaction
        #: The status of every cached formula value the workbook holds
        #: (:attr:`Workbook.calculation_status <xlsx_agent.Workbook.calculation_status>`):
        #: ``cached_unverified`` or ``invalidated``; a formula without one is ``absent``.
        self.cache = cache
        #: The defined names (case folded) that refer to another workbook: a formula using
        #: one is ``external`` as one with a workbook index is.
        self.external_names = external_names
        #: Hidden columns as sorted, non-overlapping ``(first, last)`` spans, known once
        #: the pass has passed ``cols``.
        self.hidden_columns: list[tuple[int, int]] = []
        #: Hidden rows met so far.
        self.hidden_rows: set[int] = set()
        self.merges: list[Area] = []
        self.validations: list[Validation] = []
        #: The blocks of the array formulas met so far, and the sheet's hyperlinks.
        self.arrays: list[Area] = []
        self.hyperlinks: list[Hyperlink] = []
        self.dimension: str | None = None
        self.ordered = True
        #: Set to a list by :func:`scan`: the pass then notes a checkpoint every
        #: :data:`CHECKPOINT_ROWS` rows, and the hidden rows in the order met, from which
        #: :class:`RowIndex` is made.
        self.marks: list[tuple[int, int, int, int, int]] | None = None
        self.hidden_order: list[int] = []
        #: How many row elements the pass met.
        self.row_elements = 0
        self._shared: dict[str, tuple[str, bool, bool]] = {}
        #: Each shared formula group met, by its first cell's address: the group's text and
        #: the range it fills (its ``ref``), for what its cells reach
        #: (:meth:`Regions.formula_withheld <xlsx_agent.read.regions.Regions.formula_withheld>`).
        self.shared_groups: dict[str, tuple[str, str | None]] = {}
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

    def _walk(self, start: "tuple[RowIndex, Checkpoint] | None" = None):
        """``(row, column, row hidden, c element)`` for each cell.  The pass asks for the
        ``end`` of each row (and of ``col``, ``mergeCell``, ``dimension``), takes the row's
        cells, then drops the row: one row is held at a time.  ``start``: a sheet's
        :class:`RowIndex` and one of its checkpoints, where the pass starts instead of the
        top, as it would be there."""
        if not self.data:
            return
        row = 0
        last_row = last_column = 0
        hidden_rows = self.hidden_rows
        hidden_spans: list[tuple[int, int]] = []
        match_ref = _REF.fullmatch
        columns = _COLUMNS
        cell_tags, row_tags = _C, _ROW
        marks, hidden_order = self.marks, self.hidden_order
        elements = 0
        source = self.data
        if start is not None:
            index, checkpoint = start
            source = _Resumed(index.header, self.data, checkpoint.offset)
            row, last_row, last_column = checkpoint.row, checkpoint.last_row, checkpoint.last_column
            hidden_rows.update(index.hidden[:checkpoint.hidden])
        for _, node in iterparse_xml(source, tag=_STREAMED):
            tag = node.tag
            if tag in row_tags:
                if marks is not None and elements % CHECKPOINT_ROWS == 0:
                    marks.append((elements, row, last_row, last_column, len(hidden_order)))
                elements += 1
                self.row_elements = elements
                number = _int(node.get("r"))
                row = number if number is not None and 1 <= number <= MAX_ROW else row + 1
                row_hidden = _true(node.get("hidden"))
                if row_hidden and row not in hidden_rows:
                    hidden_rows.add(row)
                    hidden_order.append(row)
                column = 0
                # A cell's ``r`` is nearly always its column's letters and this row's
                # number: those are looked up, the rest matched.
                digits = str(row)
                cut = -len(digits)
                for cell in node:
                    if cell.tag not in cell_tags:
                        continue
                    ref = cell.get("r")
                    if ref and ref.endswith(digits) and ref[:cut] in columns:
                        column, cell_row, cell_hidden = columns[ref[:cut]], row, row_hidden
                    elif (match := match_ref(ref) if ref else None) is not None:
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
            elif tag in _HYPERLINK:
                area = _area(node.get("ref"))
                if area is not None:
                    rel_id = next((node.get(a) for a in _R_ID if node.get(a)), None)
                    self.hyperlinks.append(Hyperlink(area, rel_id, node.get("location") or None))
            elif tag in _VALIDATION:
                formula = next((child.text for child in node if child.tag in _FORMULA1), None)
                self.validations.append(_validation(node.get("type"), node.get("sqref"), formula))
            elif tag in _X14_VALIDATION:
                formula = sqref = None
                for child in node:
                    if child.tag == _X14_FORMULA1:
                        formula = next((f.text for f in child if f.tag == _XM_F), None)
                    elif child.tag == _XM_SQREF:
                        sqref = child.text
                self.validations.append(_validation(node.get("type"), sqref, formula))

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
            if self.redaction is not None and value.format:
                value.format = self.redaction.text(value.format)

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
        if value.formula:
            if value.t == "blank" and value.issue is None:
                value.issue = "no_cached_value"
            value.cache = "absent" if value.issue == "no_cached_value" else self.cache
        return value

    @staticmethod
    def shared_index(node) -> int | None:
        """The shared string a ``t="s"`` cell's ``v`` names, as :meth:`value` reads it
        (``None`` for none or no number)."""
        v_text = None
        for child in node:
            if child.tag in _V:
                v_text = child.text or ""
        return _int(v_text)

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
        if looks_like_number(text, self.locale, grouping=self.grouping):
            return "number_as_text"
        if looks_like_date(text, self.locale, local_dates=self.local_dates):
            return "date_as_text"
        return None

    def _text(self, value: Value, text: str) -> None:
        if self.redaction is not None and text:
            text = self.redaction.text(text)
        value.v, value.t = text, "text"
        if len(text) >= TEXT_LIMIT:
            value.at_limit = True
        if not value.formula:
            value.issue = self._looks(text)

    def _number(self, value: Value, v_text: str | None, node) -> None:
        if not v_text:
            return
        text = v_text.strip()
        if text.isdigit() and text.isascii():   # most numbers: a whole one, no sign
            number = int(text)
        else:
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
        style = node.get("s")
        number_format = self._format_cache.get(style) or self._format(style)
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

    def note_shared(self, row: int, column: int, node) -> None:
        """What typing a cell would note of its formula for the cells after it: the first
        cell of a shared formula group (whose other cells name it)."""
        f_node = None
        for child in node:
            if child.tag in _F:
                f_node = child   # the last, as value() takes it
        if f_node is not None and f_node.get("t") == "shared" and f_node.text:
            group = f_node.get("si")
            if group is not None:
                anchor = f"{self.sheet}!{_letters(column)}{row}"
                self._shared[group] = (anchor, is_volatile(f_node.text),
                                       is_external(f_node.text, self.external_names))
                self.shared_groups[anchor] = (f_node.text, f_node.get("ref"))

    def _formula(self, value: Value, f_node) -> None:
        value.formula = True
        text = f_node.text
        kind = f_node.get("t")
        if text and self.redaction is not None and self.redaction.holds(text):
            self._formula_parts(value, f_node, text, kind)
            value.formula_text = None
            value.formula_withheld = True
            return
        self._formula_parts(value, f_node, text, kind)

    def _formula_parts(self, value: Value, f_node, text: str | None, kind: str | None) -> None:
        if kind == "shared":
            group = f_node.get("si")
            if text:
                value.formula_text = text
                value.volatile = is_volatile(text)
                value.external = is_external(text, self.external_names)
                if group is not None:
                    self._shared[group] = (value.a, value.volatile, value.external)
                    self.shared_groups[value.a] = (text, f_node.get("ref"))
            elif group in self._shared:
                value.shared_from, value.volatile, value.external = self._shared[group]
            return
        if kind == "array":
            value.array = f_node.get("ref") or value.ref
            block = _area(value.array)
            if block is not None:
                self.arrays.append(block)
        if text:
            value.formula_text = text
            value.volatile = is_volatile(text)
            value.external = is_external(text, self.external_names)


def scan(reader: SheetReader) -> SheetScan:
    """One full pass over ``reader``'s sheet: the used range, merges, ordering and the
    issues met (see :class:`SheetScan`)."""
    result = SheetScan(reader.sheet)
    reader.marks = []
    top = left = MAX_ROW + MAX_COLUMN
    bottom = right = 0
    for value in reader.values():
        if value.t == "blank" and not value.formula:
            if value.issue is not None:
                result._count(value.issue, value)
            continue
        result.cells += 1
        row, column = value.row, value.column
        if row < top:
            top = row
        if row > bottom:
            bottom = row
        if column < left:
            left = column
        if column > right:
            right = column
        if value.issue is not None:
            result._count(value.issue, value)
        if value.at_limit:
            result._count("at_limit", value)
        if value.volatile:
            result._count("volatile", value)
        if value.external:
            result._count("external_formula", value)
    if result.cells:
        result.used = Area(Cell(top, left), Cell(bottom, right))
    result.dimension = reader.dimension
    result.merges = list(reader.merges)
    result.validations = list(reader.validations)
    result.arrays = list(reader.arrays)
    result.hyperlinks = list(reader.hyperlinks)
    result.hidden_rows = len(reader.hidden_rows)
    result.ordered = reader.ordered
    if reader.ordered:
        result.index = RowIndex.build(reader.data, reader)
    reader.marks = None
    return result


def read_area(reader: SheetReader, area: Area, ordered: bool,
              index: RowIndex | None = None,
              columns: "set[int] | None" = None) -> Iterator[list[Value]]:
    """The rows of ``area``, each a list of values from its first column to its last,
    blanks filled in.  With ``ordered`` (the sheet's rows and cells in order, as a scan
    found) the pass starts at ``index``'s last checkpoint before the area, if any, and
    stops after the area's last row; otherwise the area's cells are gathered first.

    ``columns`` (sheet columns), for a pass that looks at some of the area's columns only:
    the cells of the others are not typed, and come as blanks.  The set is read as each
    row is reached, so the pass may change it between rows."""
    top, left = area.first.row, area.first.column
    bottom, right = area.last.row, area.last.column
    width = right - left + 1
    pending: dict[int, dict[int, Value]] = {}
    if not ordered:
        for value in reader.values():
            if top <= value.row <= bottom and left <= value.column <= right \
                    and (columns is None or value.column in columns):
                pending.setdefault(value.row, {})[value.column] = value
        for row in range(top, bottom + 1):
            yield _row(reader, row, left, width, pending.pop(row, None))
        return
    next_row = top
    current: dict[int, Value] | None = None
    current_row = None
    checkpoint = index.before(top) if index is not None else None
    start = (index, checkpoint) if checkpoint is not None else None
    for row, column, row_hidden, node in reader._walk(start):
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
            if columns is None or column in columns:
                current[column] = reader.value(row, column, row_hidden, node)
            else:
                reader.note_shared(row, column, node)
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


def _validation(kind: str | None, sqref: str | None, formula: str | None) -> Validation:
    areas = tuple(area for area in map(_area, (sqref or "").split()) if area is not None)
    return Validation((kind or "none").strip(), areas, formula)


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
