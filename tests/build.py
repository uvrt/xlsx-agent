"""A small SpreadsheetML writer, for the tests only: made-up workbooks built from Python.

The repository holds no workbook from anywhere else (ROADMAP.md, "Testing"); every fixture is
written here, with neutral content::

    from build import Error, Formula, Inline, Rich, Run, WorkbookBuilder

    book = WorkbookBuilder()
    stock = book.sheet("Stock")
    stock.rows("A1", [["SKU", "Price", "Added"],
                      ["SKU-0001", 12.5, date(2026, 1, 5)],
                      ["SKU-0002", Error("#N/A"), None]])
    stock.table("Products", "A1:C3", ["SKU", "Price", "Added"], write_header=False)
    stock.cell("E1", Formula("SUM(B2:B3)", cached=12.5))
    book.define_name("Prices", "Stock!$B$2:$B$3")
    data = book.to_bytes()

What is written is what ECMA-376 Part 1 asks for (the OPC package of Part 2: content types and
relationships; the workbook, worksheet, shared-string, style, table and comment parts of
§18), and nothing taken from what Excel happens to write: where Excel's own conventions
matter (a second fill, a comment's VML drawing, a default font), the Excel oracle decides
later, by measurement.  Options exist for the messy cases the reader must meet: hidden and
very hidden sheets, rows and cells without ``r``, a wrong or missing ``dimension``, formulas
without a cached value, both date systems.

Output is deterministic: the same calls give the same bytes.
"""

from __future__ import annotations

import io
import math
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, time
from typing import Any, Sequence

from lxml import etree

from ooxml_common.kinds import main_content_type

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_OFFICE_DOCUMENT = NS_R + "/officeDocument"
REL_WORKSHEET = NS_R + "/worksheet"
REL_STYLES = NS_R + "/styles"
REL_SHARED_STRINGS = NS_R + "/sharedStrings"
REL_TABLE = NS_R + "/table"
REL_COMMENTS = NS_R + "/comments"

_SML = "application/vnd.openxmlformats-officedocument.spreadsheetml."
CT_WORKSHEET = _SML + "worksheet+xml"
CT_STYLES = _SML + "styles+xml"
CT_SHARED_STRINGS = _SML + "sharedStrings+xml"
CT_TABLE = _SML + "table+xml"
CT_COMMENTS = _SML + "comments+xml"
CT_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"

#: The first id given to a format code of the workbook's own: clear of every built-in id
#: ECMA-376 Part 1 §18.8.30 lists.
FIRST_CUSTOM_FORMAT_ID = 164

#: ECMA-376's ST_TotalsRowFunction.
TOTALS_FUNCTIONS = ("none", "sum", "min", "max", "average", "count", "countNums", "stdDev", "var",
                    "custom")

#: ECMA-376's ST_SheetState.
SHEET_STATES = ("visible", "hidden", "veryHidden")

# A fixed entry time, so the same workbook is the same bytes (the earliest a ZIP can hold).
_ZIP_TIME = (1980, 1, 1, 0, 0, 0)


def _x(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


# ---------------------------------------------------------------------------------------------
# Cell values
# ---------------------------------------------------------------------------------------------

@dataclass(frozen=True)
class Error:
    """An error value (``t="e"``): ``Error("#N/A")``."""

    code: str


@dataclass(frozen=True)
class Run:
    """One run of a rich string, with the few run properties a test needs."""

    text: str
    bold: bool = False
    italic: bool = False
    size: float | None = None
    #: An ARGB hex colour, ``"FF0000FF"``.
    color: str | None = None


@dataclass(frozen=True)
class Rich:
    """A rich string: runs (``Run`` or plain ``str``), and phonetic runs as
    ``(start, end, text)`` over the base text's characters (``rPh``'s ``sb`` and ``eb``)."""

    runs: tuple = ()
    phonetic: tuple = ()

    def __post_init__(self) -> None:
        object.__setattr__(self, "runs", tuple(Run(r) if isinstance(r, str) else r for r in self.runs))
        object.__setattr__(self, "phonetic", tuple(tuple(p) for p in self.phonetic))

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)


@dataclass(frozen=True)
class Inline:
    """A string written in the cell (``t="inlineStr"``), not in the shared-string table."""

    value: "str | Rich"


@dataclass(frozen=True)
class IsoDate:
    """A date written as ISO 8601 text (``t="d"``, ST_CellType §18.18.11), not as a serial."""

    value: "date | datetime | time"


@dataclass(frozen=True)
class Formula:
    """A formula (text without the leading ``=``) and its cached value, if any.

    ``cached`` is a number, ``str`` (``t="str"``), ``bool``, :class:`Error`, or ``None`` for a
    formula with no cached value.  ``kind`` is ``"normal"``, ``"shared"`` (with ``si``; the
    group's master also gives ``ref`` and the text, a member only ``si``) or ``"array"`` (with
    ``ref``)."""

    text: str
    cached: Any = None
    kind: str = "normal"
    ref: str | None = None
    si: int | None = None


def date_serial(value: "date | datetime | time", date1904: bool = False) -> float:
    """A date, time or both as a serial in the workbook's date system (ECMA-376 Part 1
    §18.17.4.1): in the 1900 system serial 1 is 1900-01-01 and serial 60 the 29 February 1900
    the system counts, so later dates are one more than their distance from 1899-12-31; in
    the 1904 system serial 0 is 1904-01-01.  A ``time`` alone is the fraction of a day."""
    if isinstance(value, time):
        return (value.hour * 3600 + value.minute * 60 + value.second
                + value.microsecond / 1e6) / 86400
    day = value.date() if isinstance(value, datetime) else value
    if date1904:
        serial = (day - date(1904, 1, 1)).days
    else:
        serial = (day - date(1899, 12, 31)).days
        if day >= date(1900, 3, 1):
            serial += 1
    if serial < 0 or (not date1904 and serial < 1):
        raise ValueError(f"{value!r} is before the {'1904' if date1904 else '1900'} date system")
    if isinstance(value, datetime):
        return serial + date_serial(value.time())
    return float(serial)


# ---------------------------------------------------------------------------------------------
# References
# ---------------------------------------------------------------------------------------------

_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def column_number(letters: str) -> int:
    """``"A"`` -> 1, ``"AB"`` -> 28."""
    number = 0
    for char in letters.upper():
        number = number * 26 + ord(char) - 64
    return number


def column_name(number: int) -> str:
    """1 -> ``"A"``, 28 -> ``"AB"``."""
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def parse_cell(ref: str) -> tuple[int, int]:
    """``"B3"`` -> ``(3, 2)``: row and column, both from 1."""
    match = _CELL.match(ref)
    if not match:
        raise ValueError(f"not a cell reference: {ref!r}")
    return int(match.group(2)), column_number(match.group(1))


def parse_area(ref: str) -> tuple[tuple[int, int], tuple[int, int]]:
    """``"A1:C4"`` -> ``((1, 1), (4, 3))``; a single cell is a one-cell area."""
    first, _, last = ref.partition(":")
    return parse_cell(first), parse_cell(last or first)


def cell_name(row: int, column: int) -> str:
    return f"{column_name(column)}{row}"


# ---------------------------------------------------------------------------------------------
# Parts of a sheet
# ---------------------------------------------------------------------------------------------

@dataclass
class _Table:
    name: str
    ref: str
    columns: list[str]
    header_row: bool
    totals_row: bool
    totals_labels: dict[str, str]
    totals_functions: dict[str, str]
    style_name: str | None
    id: int = 0


@dataclass
class _Comment:
    ref: str
    text: "str | Rich"
    author: str


@dataclass
class _Validation:
    sqref: str
    type: str
    formula1: str | None
    formula2: str | None
    operator: str | None
    allow_blank: bool


@dataclass
class SheetBuilder:
    """One worksheet.  Made by :meth:`WorkbookBuilder.sheet`."""

    book: "WorkbookBuilder" = field(repr=False, compare=False)
    name: str
    sheet_id: int
    state: str = "visible"
    #: Write the sheet with ``sheetProtection sheet="1"``.
    protected: bool = False
    #: ``None``: the used range, computed from the cells (omitted when there are none);
    #: ``False``: no ``dimension``; a string: written as given, right or wrong.
    dimension: "str | bool | None" = None
    #: Leave out each row's and cell's ``r`` where its position follows from the one before
    #: (the next row, the next column), as a reader must then compute it.
    omit_refs: bool = False
    cells: dict[tuple[int, int], tuple[Any, int]] = field(default_factory=dict)
    hidden_rows: set[int] = field(default_factory=set)
    hidden_columns: set[int] = field(default_factory=set)
    merges: list[str] = field(default_factory=list)
    tables: list[_Table] = field(default_factory=list)
    comments: list[_Comment] = field(default_factory=list)
    validations: list[_Validation] = field(default_factory=list)

    # -- cells ----------------------------------------------------------------------------------

    def cell(self, ref: str, value: Any, *, style: int = 0) -> "SheetBuilder":
        """Set one cell.  ``value``: ``None`` (a cell with only its style), a number, ``bool``,
        ``str`` (a shared string), :class:`Rich`, :class:`Inline`, :class:`Error`,
        :class:`Formula`, :class:`IsoDate`, or a ``date``/``datetime``/``time`` (a serial in
        the workbook's date system, with a date style of the builder's own unless ``style``
        is given)."""
        if isinstance(value, (date, time)) and not style:
            style = self.book.style(num_fmt=_default_date_format(value))
        self.cells[parse_cell(ref)] = (value, style)
        return self

    def rows(self, top_left: str, rows: Sequence[Sequence[Any]], *, style: int = 0) -> "SheetBuilder":
        """Set a block of cells row by row from ``top_left``; ``None`` leaves a cell out."""
        top, left = parse_cell(top_left)
        for r, values in enumerate(rows):
            for c, value in enumerate(values):
                if value is not None:
                    self.cell(cell_name(top + r, left + c), value, style=style)
        return self

    def hide_row(self, row: int) -> "SheetBuilder":
        self.hidden_rows.add(row)
        return self

    def hide_column(self, column: str) -> "SheetBuilder":
        self.hidden_columns.add(column_number(column))
        return self

    # -- what lies over the cells ---------------------------------------------------------------

    def merge(self, ref: str) -> "SheetBuilder":
        parse_area(ref)
        self.merges.append(ref)
        return self

    def comment(self, ref: str, text: "str | Rich", *, author: str = "Reviewer") -> "SheetBuilder":
        parse_cell(ref)
        self.comments.append(_Comment(ref, text, author))
        return self

    def validation(self, sqref: str, type: str, formula1: str | None = None,
                   formula2: str | None = None, *, operator: str | None = None,
                   allow_blank: bool = False) -> "SheetBuilder":
        """A data-validation rule as ECMA-376 writes it: ``type`` is its
        ST_DataValidationType (``"list"``, ``"whole"``, ``"decimal"``...)."""
        self.validations.append(_Validation(sqref, type, formula1, formula2, operator, allow_blank))
        return self

    def validation_list(self, sqref: str, values: Sequence[str] | None = None, *,
                        source: str | None = None, allow_blank: bool = False) -> "SheetBuilder":
        """A drop-down list: inline ``values`` (written as one quoted, comma-separated
        formula) or a ``source`` range such as ``"$F$1:$F$4"``."""
        if (values is None) == (source is None):
            raise ValueError("give values or source, not both")
        if values is not None:
            if any("," in value for value in values):
                raise ValueError("an inline list cannot hold a comma")
            formula = '"' + ",".join(values).replace('"', '""') + '"'
        else:
            formula = source
        return self.validation(sqref, "list", formula, allow_blank=allow_blank)

    def table(self, name: str, ref: str, columns: Sequence[str], *, header_row: bool = True,
              totals_row: bool = False, totals_labels: dict[str, str] | None = None,
              totals_functions: dict[str, str] | None = None, write_header: bool = True,
              style_name: str | None = None) -> "SheetBuilder":
        """An Excel table (a ListObject, §18.5) over ``ref``, which includes its header and
        totals rows.  ``write_header`` writes the column names into the header row's cells
        (shared strings); the data and totals cells are the caller's."""
        (top, left), (bottom, right) = parse_area(ref)
        if len(columns) != right - left + 1:
            raise ValueError(f"{len(columns)} columns for a table {right - left + 1} wide")
        totals_labels = dict(totals_labels or {})
        totals_functions = dict(totals_functions or {})
        for function in totals_functions.values():
            if function not in TOTALS_FUNCTIONS:
                raise ValueError(f"not a totals-row function: {function!r}")
        for column in (*totals_labels, *totals_functions):
            if column not in columns:
                raise ValueError(f"no column {column!r}")
        totals_row = totals_row or bool(totals_labels or totals_functions)
        if bottom - top + 1 < int(header_row) + int(totals_row) + 1:
            raise ValueError("a table needs at least one data row")
        if header_row and write_header:
            for offset, column in enumerate(columns):
                self.cell(cell_name(top, left + offset), column)
        self.tables.append(_Table(name, ref, list(columns), header_row, totals_row,
                                  totals_labels, totals_functions, style_name))
        return self

    # -- writing --------------------------------------------------------------------------------

    def used_range(self) -> str | None:
        if not self.cells:
            return None
        rows = [row for row, _ in self.cells]
        columns = [column for _, column in self.cells]
        first, last = cell_name(min(rows), min(columns)), cell_name(max(rows), max(columns))
        return first if first == last else f"{first}:{last}"

    def _xml(self, strings: "_SharedStrings", rels: list[tuple[str, str, str]]) -> etree._Element:
        root = etree.Element(_x("worksheet"), nsmap={None: NS_MAIN, "r": NS_R})
        dimension = self.used_range() if self.dimension is None else self.dimension
        if dimension:
            etree.SubElement(root, _x("dimension"), ref=dimension)
        if self.hidden_columns:
            cols = etree.SubElement(root, _x("cols"))
            for column in sorted(self.hidden_columns):
                etree.SubElement(cols, _x("col"), min=str(column), max=str(column), hidden="1")
        data = etree.SubElement(root, _x("sheetData"))
        by_row: dict[int, list[int]] = {}
        for row, column in self.cells:
            by_row.setdefault(row, []).append(column)
        for row in self.hidden_rows:
            by_row.setdefault(row, [])
        previous_row = 0
        for row in sorted(by_row):
            node = etree.SubElement(data, _x("row"))
            if not (self.omit_refs and row == previous_row + 1):
                node.set("r", str(row))
            if row in self.hidden_rows:
                node.set("hidden", "1")
            previous_row, previous_column = row, 0
            for column in sorted(by_row[row]):
                value, style = self.cells[(row, column)]
                ref = None if self.omit_refs and column == previous_column + 1 else cell_name(row, column)
                node.append(self._cell(ref, value, style, strings))
                previous_column = column
        if self.protected:
            etree.SubElement(root, _x("sheetProtection"), sheet="1")
        if self.merges:
            merges = etree.SubElement(root, _x("mergeCells"), count=str(len(self.merges)))
            for ref in self.merges:
                etree.SubElement(merges, _x("mergeCell"), ref=ref)
        if self.validations:
            rules = etree.SubElement(root, _x("dataValidations"), count=str(len(self.validations)))
            for rule in self.validations:
                node = etree.SubElement(rules, _x("dataValidation"), type=rule.type)
                if rule.operator:
                    node.set("operator", rule.operator)
                if rule.allow_blank:
                    node.set("allowBlank", "1")
                node.set("sqref", rule.sqref)
                for tag, formula in (("formula1", rule.formula1), ("formula2", rule.formula2)):
                    if formula is not None:
                        etree.SubElement(node, _x(tag)).text = formula
        if self.tables:
            parts = etree.SubElement(root, _x("tableParts"), count=str(len(self.tables)))
            for table in self.tables:
                rel_id = f"rId{len(rels) + 1}"
                rels.append((rel_id, REL_TABLE, f"../tables/table{table.id}.xml"))
                etree.SubElement(parts, _x("tablePart")).set(f"{{{NS_R}}}id", rel_id)
        return root

    def _cell(self, ref: str | None, value: Any, style: int,
              strings: "_SharedStrings") -> etree._Element:
        node = etree.Element(_x("c"))
        if ref is not None:
            node.set("r", ref)
        if style:
            node.set("s", str(style))
        if value is None:
            return node
        if isinstance(value, Formula):
            f = etree.SubElement(node, _x("f"))
            if value.text:
                f.text = value.text
            if value.kind != "normal":
                if value.kind not in ("shared", "array"):
                    raise ValueError(f"not a formula kind: {value.kind!r}")
                f.set("t", value.kind)
            if value.ref is not None:
                f.set("ref", value.ref)
            if value.si is not None:
                f.set("si", str(value.si))
            if value.cached is not None:
                kind, text = _scalar(value.cached, self.book.date1904)
                if kind == "s":
                    kind = "str"
                if kind != "n":
                    node.set("t", kind)
                etree.SubElement(node, _x("v")).text = text
            return node
        if isinstance(value, Inline):
            node.set("t", "inlineStr")
            node.append(_string_item(_x("is"), value.value))
            return node
        if isinstance(value, IsoDate):
            node.set("t", "d")
            etree.SubElement(node, _x("v")).text = value.value.isoformat()
            return node
        if isinstance(value, (str, Rich)):
            node.set("t", "s")
            etree.SubElement(node, _x("v")).text = str(strings.index(value))
            return node
        kind, text = _scalar(value, self.book.date1904)
        if kind != "n":
            node.set("t", kind)
        etree.SubElement(node, _x("v")).text = text
        return node


def _scalar(value: Any, date1904: bool) -> tuple[str, str]:
    """A scalar's cell type and ``v`` text."""
    if isinstance(value, bool):
        return "b", "1" if value else "0"
    if isinstance(value, Error):
        return "e", value.code
    if isinstance(value, (date, time)):
        return "n", _number(date_serial(value, date1904))
    if isinstance(value, (int, float)):
        return "n", _number(value)
    if isinstance(value, str):
        return "s", value
    raise TypeError(f"not a cell value: {type(value).__name__}")


def _number(value: "int | float") -> str:
    if isinstance(value, int):
        return str(value)
    if not math.isfinite(value):
        raise ValueError("a cell holds finite numbers only")
    return str(int(value)) if value.is_integer() and abs(value) < 1e15 else repr(value)


def _default_date_format(value: "date | datetime | time") -> str:
    """The builder's own format codes for date cells given no style (ISO order; not Excel's
    built-in formats, which the oracle measures)."""
    if isinstance(value, datetime):
        return "yyyy-mm-dd hh:mm:ss"
    if isinstance(value, time):
        return "hh:mm:ss"
    return "yyyy-mm-dd"


def _text(parent: etree._Element, text: str) -> None:
    t = etree.SubElement(parent, _x("t"))
    t.text = text
    if text != text.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _string_item(tag: str, value: "str | Rich") -> etree._Element:
    """A CT_Rst element (``si``, ``is``, a comment's ``text``): plain text or rich runs."""
    node = etree.Element(tag)
    if isinstance(value, str):
        _text(node, value)
        return node
    for run in value.runs:
        r = etree.SubElement(node, _x("r"))
        properties = []
        if run.bold:
            properties.append(etree.Element(_x("b")))
        if run.italic:
            properties.append(etree.Element(_x("i")))
        if run.color:
            properties.append(etree.Element(_x("color"), rgb=run.color))
        if run.size is not None:
            properties.append(etree.Element(_x("sz"), val=_number(run.size)))
        if properties:
            etree.SubElement(r, _x("rPr")).extend(properties)
        _text(r, run.text)
    for start, end, text in value.phonetic:
        ph = etree.SubElement(node, _x("rPh"), sb=str(start), eb=str(end))
        _text(ph, text)
    if value.phonetic:
        etree.SubElement(node, _x("phoneticPr"), fontId="0")
    return node


class _SharedStrings:
    """The shared-string table: each distinct string once, in first-use order."""

    def __init__(self) -> None:
        self.items: list["str | Rich"] = []
        self._index: dict[Any, int] = {}
        self.count = 0

    def index(self, value: "str | Rich") -> int:
        self.count += 1
        if value not in self._index:
            self._index[value] = len(self.items)
            self.items.append(value)
        return self._index[value]

    def xml(self) -> etree._Element:
        root = etree.Element(_x("sst"), nsmap={None: NS_MAIN},
                             count=str(self.count), uniqueCount=str(len(self.items)))
        for item in self.items:
            root.append(_string_item(_x("si"), item))
        return root


# ---------------------------------------------------------------------------------------------
# The workbook
# ---------------------------------------------------------------------------------------------

@dataclass
class _Name:
    name: str
    refers_to: str
    sheet: str | None
    hidden: bool


class WorkbookBuilder:
    """A workbook to write.  ``kind`` is the package kind (``"xlsx"``, ``"xlsm"``, ``"xltx"``,
    ``"xltm"``) whose main content type the workbook part declares; ``calc`` the ``calcPr``
    attributes, if any (``{"fullCalcOnLoad": "1"}``, ``{"calcMode": "manual"}``)."""

    def __init__(self, *, date1904: bool = False, kind: str = "xlsx",
                 calc: dict[str, str] | None = None, protected: bool = False) -> None:
        self.date1904 = date1904
        self.content_type = main_content_type(kind)
        self.calc = dict(calc or {})
        self.protected = protected
        self.sheets: list[SheetBuilder] = []
        self.names: list[_Name] = []
        #: Style 0 is the default; each style is a number format id.
        self._styles: list[int] = [0]
        self._formats: dict[str, int] = {}

    def sheet(self, name: str, *, state: str = "visible", sheet_id: int | None = None) -> SheetBuilder:
        if state not in SHEET_STATES:
            raise ValueError(f"not a sheet state: {state!r}")
        if any(sheet.name == name for sheet in self.sheets):
            raise ValueError(f"a sheet {name!r} exists")
        if sheet_id is None:
            sheet_id = max((sheet.sheet_id for sheet in self.sheets), default=0) + 1
        if any(sheet.sheet_id == sheet_id for sheet in self.sheets):
            raise ValueError(f"sheet id {sheet_id} is taken")
        sheet = SheetBuilder(self, name, sheet_id, state)
        self.sheets.append(sheet)
        return sheet

    def style(self, *, num_fmt: str | None = None, num_fmt_id: int | None = None) -> int:
        """The index of a cell style (``cellXfs``) with a number format: a format code of
        the workbook's own (``num_fmt``) or a built-in id with no code in the file
        (``num_fmt_id``).  The same format gives the same index."""
        if (num_fmt is None) == (num_fmt_id is None):
            raise ValueError("give num_fmt or num_fmt_id")
        if num_fmt is not None:
            if num_fmt not in self._formats:
                self._formats[num_fmt] = FIRST_CUSTOM_FORMAT_ID + len(self._formats)
            num_fmt_id = self._formats[num_fmt]
        if num_fmt_id not in self._styles:
            self._styles.append(num_fmt_id)
        return self._styles.index(num_fmt_id)

    def define_name(self, name: str, refers_to: str, *, sheet: str | None = None,
                    hidden: bool = False) -> "WorkbookBuilder":
        """A defined name, workbook-wide or local to ``sheet``; ``refers_to`` is its formula
        text (``"Stock!$B$2:$B$9"``, or a constant or formula)."""
        if sheet is not None and not any(s.name == sheet for s in self.sheets):
            raise ValueError(f"no sheet {sheet!r}")
        self.names.append(_Name(name, refers_to, sheet, hidden))
        return self

    # -- writing --------------------------------------------------------------------------------

    def parts(self) -> dict[str, bytes]:
        """Every part by name, in the order they are written; ``[Content_Types].xml`` first."""
        strings = _SharedStrings()
        overrides: list[tuple[str, str]] = [("/xl/workbook.xml", self.content_type)]
        body: dict[str, bytes] = {}
        workbook_rels: list[tuple[str, str, str]] = []
        table_id = 0
        comment_id = 0
        for number, sheet in enumerate(self.sheets, 1):
            for table in sheet.tables:
                table_id += 1
                table.id = table_id
            rels: list[tuple[str, str, str]] = []
            path = f"xl/worksheets/sheet{number}.xml"
            body[path] = _serialize(sheet._xml(strings, rels))
            overrides.append(("/" + path, CT_WORKSHEET))
            workbook_rels.append((f"rId{number}", REL_WORKSHEET, f"worksheets/sheet{number}.xml"))
            for table in sheet.tables:
                table_path = f"xl/tables/table{table.id}.xml"
                body[table_path] = _serialize(_table_xml(table))
                overrides.append(("/" + table_path, CT_TABLE))
            if sheet.comments:
                comment_id += 1
                comments_path = f"xl/comments{comment_id}.xml"
                body[comments_path] = _serialize(_comments_xml(sheet.comments))
                overrides.append(("/" + comments_path, CT_COMMENTS))
                rels.append((f"rId{len(rels) + 1}", REL_COMMENTS, f"../comments{comment_id}.xml"))
            if rels:
                body[f"xl/worksheets/_rels/sheet{number}.xml.rels"] = _serialize(_rels_xml(rels))
        workbook_rels.append((f"rId{len(workbook_rels) + 1}", REL_STYLES, "styles.xml"))
        body["xl/styles.xml"] = _serialize(self._styles_xml())
        overrides.append(("/xl/styles.xml", CT_STYLES))
        if strings.items:
            workbook_rels.append((f"rId{len(workbook_rels) + 1}", REL_SHARED_STRINGS, "sharedStrings.xml"))
            body["xl/sharedStrings.xml"] = _serialize(strings.xml())
            overrides.append(("/xl/sharedStrings.xml", CT_SHARED_STRINGS))
        body["xl/workbook.xml"] = _serialize(self._workbook_xml(workbook_rels))
        body["xl/_rels/workbook.xml.rels"] = _serialize(_rels_xml(workbook_rels))

        parts = {"[Content_Types].xml": _serialize(_content_types_xml(overrides)),
                 "_rels/.rels": _serialize(_rels_xml([("rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml")]))}
        parts["xl/workbook.xml"] = body.pop("xl/workbook.xml")
        parts["xl/_rels/workbook.xml.rels"] = body.pop("xl/_rels/workbook.xml.rels")
        parts.update(body)
        return parts

    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            for name, data in self.parts().items():
                info = zipfile.ZipInfo(name, _ZIP_TIME)
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, data)
        return buffer.getvalue()

    def save(self, path) -> None:
        with open(path, "wb") as handle:
            handle.write(self.to_bytes())

    def _workbook_xml(self, rels: list[tuple[str, str, str]]) -> etree._Element:
        root = etree.Element(_x("workbook"), nsmap={None: NS_MAIN, "r": NS_R})
        if self.date1904:
            etree.SubElement(root, _x("workbookPr"), date1904="1")
        if self.protected:
            etree.SubElement(root, _x("workbookProtection"), lockStructure="1")
        sheets = etree.SubElement(root, _x("sheets"))
        for number, sheet in enumerate(self.sheets, 1):
            node = etree.SubElement(sheets, _x("sheet"), name=sheet.name, sheetId=str(sheet.sheet_id))
            if sheet.state != "visible":
                node.set("state", sheet.state)
            node.set(f"{{{NS_R}}}id", rels[number - 1][0])
        if self.names:
            names = etree.SubElement(root, _x("definedNames"))
            positions = {sheet.name: index for index, sheet in enumerate(self.sheets)}
            for name in self.names:
                node = etree.SubElement(names, _x("definedName"), name=name.name)
                if name.sheet is not None:
                    node.set("localSheetId", str(positions[name.sheet]))
                if name.hidden:
                    node.set("hidden", "1")
                node.text = name.refers_to
        if self.calc:
            etree.SubElement(root, _x("calcPr"), **self.calc)
        return root

    def _styles_xml(self) -> etree._Element:
        root = etree.Element(_x("styleSheet"), nsmap={None: NS_MAIN})
        if self._formats:
            formats = etree.SubElement(root, _x("numFmts"), count=str(len(self._formats)))
            for code, number in self._formats.items():
                etree.SubElement(formats, _x("numFmt"), numFmtId=str(number), formatCode=code)
        etree.SubElement(etree.SubElement(root, _x("fonts"), count="1"), _x("font"))
        fill = etree.SubElement(etree.SubElement(root, _x("fills"), count="1"), _x("fill"))
        etree.SubElement(fill, _x("patternFill"), patternType="none")
        etree.SubElement(etree.SubElement(root, _x("borders"), count="1"), _x("border"))
        base = etree.SubElement(root, _x("cellStyleXfs"), count="1")
        etree.SubElement(base, _x("xf"), numFmtId="0", fontId="0", fillId="0", borderId="0")
        xfs = etree.SubElement(root, _x("cellXfs"), count=str(len(self._styles)))
        for number in self._styles:
            xf = etree.SubElement(xfs, _x("xf"), numFmtId=str(number), fontId="0", fillId="0",
                                  borderId="0", xfId="0")
            if number:
                xf.set("applyNumberFormat", "1")
        styles = etree.SubElement(root, _x("cellStyles"), count="1")
        # builtinId 0 is the Normal style (ECMA-376 Part 1 §18.8.7).
        etree.SubElement(styles, _x("cellStyle"), name="Normal", xfId="0", builtinId="0")
        return root


def _table_xml(table: _Table) -> etree._Element:
    (top, left), (bottom, right) = parse_area(table.ref)
    root = etree.Element(_x("table"), nsmap={None: NS_MAIN}, id=str(table.id), name=table.name,
                         displayName=table.name, ref=table.ref)
    if not table.header_row:
        root.set("headerRowCount", "0")
    if table.totals_row:
        root.set("totalsRowCount", "1")
    if table.header_row:
        last = bottom - 1 if table.totals_row else bottom
        etree.SubElement(root, _x("autoFilter"), ref=f"{cell_name(top, left)}:{cell_name(last, right)}")
    columns = etree.SubElement(root, _x("tableColumns"), count=str(len(table.columns)))
    for index, name in enumerate(table.columns, 1):
        node = etree.SubElement(columns, _x("tableColumn"), id=str(index), name=name)
        if name in table.totals_labels:
            node.set("totalsRowLabel", table.totals_labels[name])
        if name in table.totals_functions:
            node.set("totalsRowFunction", table.totals_functions[name])
    if table.style_name:
        etree.SubElement(root, _x("tableStyleInfo"), name=table.style_name)
    return root


def _comments_xml(comments: list[_Comment]) -> etree._Element:
    root = etree.Element(_x("comments"), nsmap={None: NS_MAIN})
    authors: list[str] = []
    for comment in comments:
        if comment.author not in authors:
            authors.append(comment.author)
    names = etree.SubElement(root, _x("authors"))
    for author in authors:
        etree.SubElement(names, _x("author")).text = author
    entries = etree.SubElement(root, _x("commentList"))
    for comment in comments:
        node = etree.SubElement(entries, _x("comment"), ref=comment.ref,
                                authorId=str(authors.index(comment.author)))
        node.append(_string_item(_x("text"), comment.text))
    return root


def _rels_xml(rels: list[tuple[str, str, str]]) -> etree._Element:
    root = etree.Element(f"{{{NS_PKG_RELS}}}Relationships", nsmap={None: NS_PKG_RELS})
    for rel_id, rel_type, target in rels:
        etree.SubElement(root, f"{{{NS_PKG_RELS}}}Relationship", Id=rel_id, Type=rel_type, Target=target)
    return root


def _content_types_xml(overrides: list[tuple[str, str]]) -> etree._Element:
    root = etree.Element(f"{{{NS_CONTENT_TYPES}}}Types", nsmap={None: NS_CONTENT_TYPES})
    etree.SubElement(root, f"{{{NS_CONTENT_TYPES}}}Default", Extension="rels", ContentType=CT_RELATIONSHIPS)
    etree.SubElement(root, f"{{{NS_CONTENT_TYPES}}}Default", Extension="xml", ContentType="application/xml")
    for part, content_type in overrides:
        etree.SubElement(root, f"{{{NS_CONTENT_TYPES}}}Override", PartName=part, ContentType=content_type)
    return root


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
