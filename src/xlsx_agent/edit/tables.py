"""A block of rows written, a real Excel table made, rows appended (round 44: R54, E4).

::

    wb.write_table("Stock", "B2", ["Item", "Qty", "Added"],
                   [["Lamp", 3, date(2026, 1, 5)], ["Fan", 4, {"date": "2026-01-06"}]],
                   as_table="Stock")                 # -> TableWritten: Stock!B2:D4, table Stock
    wb.append_rows("t:Stock", [["Desk", 1, date(2026, 1, 7)]])     # the table grows
    wb.append_rows("Log", [["2026-11-08", "started"]])             # below the last value

**Values** are :meth:`set_cells <xlsx_agent.Workbook.set_cells>`' (text always literal,
numbers, booleans, dates, ``{"formula": ...}``, ``{"markdown": ...}``, ``None`` a blank);
``""`` is taken as a blank (an empty cell), as a row read from elsewhere has it.  A row is a
list in the columns' order -- shorter rows end in blanks, longer ones are refused -- or a
mapping by header (a table's column names, case aside).  A date written into a cell whose
format shows no date gets the built-in short date (``numFmtId`` 14), as a date typed into a
new cell does in Excel (measured, Excel for Mac 16.106, 2026-11-08: ``2026-01-05``,
``5-1-2026`` and ``5/1/2026`` typed each got ``numFmtId="14" applyNumberFormat="1"``);
``{"date": ..., "format": ...}`` gives another.  Every write passes set_cells' checks and
the host's guards (policy, protection, expectations); the whole call is one step, refused
whole.

**write_table(sheet, origin, headers, rows)** writes ``headers`` (text: a table's column
names) in the row at ``origin`` and the rows below, and returns the range written.
Refused, every problem listed together, before anything is kept: a cell of the block that
holds a value or a formula (``not_empty``; ``overwrite=True`` writes over it, formulas too),
a merged area or another table the block touches (``table_overlap``: never written over),
a block past the grid's edge.  ``as_table`` (``True``, or the table's name) makes it a real
Excel table (a ListObject, ECMA-376 Part 1 §18.5) as Excel makes one from a range (measured,
below), with its AutoFilter and the workbook's default table style banded by rows; a header
row is then required, its names unique (case aside), non-empty and on one line; a table with
no rows gets one empty data row, as Excel gives one.

*Table names* (measured, Excel for Mac 16.106, 2026-11-08, setting a table's name through
AppleScript): 1 to 255 characters; the first a letter (any script), ``_`` or ``\\``; the
rest letters, digits, ``_``, ``\\``, ``.``, ``?`` and characters beyond ASCII (``€``
taken; a space, ``-``, ``!``, ``$`` and ``:`` refused; other ASCII punctuation is refused
here, unmeasured); not a cell reference on the
grid (``A1``, ``XFD1`` refused; ``XFE1``, ``XFD1048577``, ``ABCD1`` taken) nor in R1C1 form
(``R``, ``C``, ``R1``, ``C1``, ``R1C1``, ``R5C`` refused; ``RC`` taken); not the name of
another table or a defined name, case aside.  ``TRUE`` and a sheet's name are taken.  A
refused name says why and offers one that would be taken.  The default is ``Table`` and
the lowest number free.

*The table part* is what Excel wrote for a table it made from a range (``make new list
object``, saved): ``id`` (the workbook's next), ``name`` and ``displayName`` the same,
``ref``, ``totalsRowShown="0"``, an ``autoFilter`` over the same ``ref``, a
``tableColumn`` per column (``id`` from 1, ``name`` its header's text) and
``tableStyleInfo`` with the style, ``showFirstColumn="0" showLastColumn="0"
showRowStripes="1" showColumnStripes="0"``; the header cells are plain text, without a
format of their own.  The style is the workbook's ``tableStyles defaultTableStyle``
(``TableStyleMedium2`` in a new workbook), or ``style``: a style the workbook defines, or a
built-in name.

**append_rows(target, rows)** -- where the rows go:

* *a table* (``"t:Orders"``, or a table's name): the rows after its last data row; the
  table grows (``ref``, and its ``autoFilter``, which leaves a totals row out), its
  calculated columns (``calculatedColumnFormula``) filled into each new row a row leaves
  blank, each new cell given the format of the cell above it when it has none of its own,
  and a totals row moved down below the new rows, its cells as they were.  Measured (Excel
  for Mac 16.106, 2026-11-08): a value typed in the row below a table grows it, gives every
  column's new cell the format above it and fills the calculated columns; a table with a
  totals row resized two rows larger keeps its totals row as the last, its
  ``SUBTOTAL(109, ...)`` formulas unchanged, and fills the calculated column into the rows
  between.  The rows the table grows into (below its last row, the totals row's included)
  must be empty in its columns, with no merged area or other table (``not_empty``,
  ``table_overlap``): Excel would shift cells there, which is inserting rows (not this).
  A totals row is moved only when nothing refers to its cells by address -- a totals
  formula's own references and every formula or defined name of the workbook are read
  (``totals_referenced``, naming each) -- since a moved cell keeps no reference to it.
  Structured references (``Orders[[#Totals],[Qty]]``) follow the table.
* *a sheet* (``"Log"``, ``"Log!"``): the row after the last row holding a value or a
  formula in the columns written -- column A to as many as the widest row -- or row 1.
  A cell holding only a format (a pre-formatted empty row) does not count; a formula
  that shows nothing does.
* *a region* (``"Log!B4:D200"``, ``"Log!B:D"``): the row after the last row holding a value
  in any of its columns, within its rows (or its first row), the rows written from its first
  column; rows wider than it are refused.  The rows may run past its last row.

Nothing is inserted and nothing below is moved but a table's totals row; a block that would
meet a merged area or a table (other than the one appended to) is refused.  Returns
:class:`Appended`: the range written, the table's range after, the totals row's new place,
the columns filled.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import Any, Mapping, Sequence

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell, Formula

from ..address import AddressError, RangeAddress, format_range, format_sheet, parse_address
from .cells import MAX_LISTED, Written

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_X = f"{{{NS_MAIN}}}"
REL_TABLE = NS_R + "/table"
REL_SHARED_STRINGS = NS_R + "/sharedStrings"
CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
CT_SHARED_STRINGS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"

#: The longest table name Excel takes (measured: 255 characters taken, 256 refused).
MAX_TABLE_NAME = 255
#: The table style a workbook without ``tableStyles defaultTableStyle`` gives a table: a
#: new workbook's (measured).
FALLBACK_TABLE_STYLE = "TableStyleMedium2"
#: The built-in short date (ECMA-376 Part 1 §18.8.30, id 14), what Excel gives a date typed
#: into a cell that shows none (measured, see the module).
SHORT_DATE = 14

_NAME_START = re.compile(r"[^\W\d]|[_\\]")
_NAME_REST = re.compile(r"[\w\\.?\u0080-\U0010ffff]*")
_A1 = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]+)")
_R1C1 = re.compile(r"[Rr][0-9]*[Cc][0-9]*|[Rr][0-9]*|[Cc][0-9]*")
_BUILTIN_STYLE = re.compile(r"TableStyle(Light|Medium|Dark)[0-9]{1,2}")


@dataclass(frozen=True)
class TableWritten:
    """What :meth:`Workbook.write_table <xlsx_agent.Workbook.write_table>` did: ``range``
    the block written (headers and rows; a table's empty data row too), ``table`` the
    table's name when one was made, ``header`` the header's range, ``rows`` the data rows
    written, ``written`` set_cells' :class:`~xlsx_agent.edit.cells.Written`.  Never a
    value."""

    range: str
    table: str | None
    header: str | None
    rows: int
    written: Written

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"range": self.range, "rows": self.rows}
        if self.header:
            out["header"] = self.header
        if self.table:
            out["table"] = self.table
        out.update({k: v for k, v in self.written.as_dict(limit).items()
                    if k in ("written", "changed", "fixed_height", "invalid", "unchecked",
                             "calculation")})
        return out


@dataclass(frozen=True)
class Appended:
    """What :meth:`Workbook.append_rows <xlsx_agent.Workbook.append_rows>` did: ``range``
    the rows written, ``rows`` how many, ``table`` the table appended to (its name) and
    ``table_range`` its range after, ``totals`` its totals row's range after it moved down
    (``None`` without one), ``filled`` the calculated columns filled into the new rows (by
    name), ``written`` set_cells' :class:`~xlsx_agent.edit.cells.Written`."""

    range: str
    rows: int
    table: str | None = None
    table_range: str | None = None
    totals: str | None = None
    filled: tuple[str, ...] = ()
    written: Written | None = None

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"range": self.range, "rows": self.rows}
        if self.table:
            out["table"] = self.table
            out["table_range"] = self.table_range
        if self.totals:
            out["totals"] = self.totals
        if self.filled:
            out["filled"] = list(self.filled)
        if self.written is not None:
            out.update({k: v for k, v in self.written.as_dict(limit).items()
                        if k in ("written", "changed", "invalid", "unchecked", "calculation")})
        return out


# -- names ---------------------------------------------------------------------------------------


def table_name_problem(name: Any, taken: "set[str] | frozenset[str]" = frozenset()) -> str | None:
    """Why Excel would refuse ``name`` for a table (see the module), or ``None``.
    ``taken``: the names already used, case folded (tables and defined names)."""
    if not isinstance(name, str) or not name:
        return "a table's name is text"
    if len(name) > MAX_TABLE_NAME:
        return f"longer than {MAX_TABLE_NAME} characters"
    if not _NAME_START.fullmatch(name[0]):
        return "it must start with a letter, _ or \\"
    if not _NAME_REST.fullmatch(name[1:]):
        return "only letters, digits, _, \\, . and ? (no spaces or other punctuation)"
    cell = _A1.fullmatch(name)
    if cell is not None:
        try:
            column = _column_number(cell.group(1))
        except ValueError:
            column = MAX_COLUMN + 1
        if column <= MAX_COLUMN and 1 <= int(cell.group(2)) <= MAX_ROW:
            return "it reads as a cell reference"
    if _R1C1.fullmatch(name) and name.upper() != "RC":
        return "it reads as an R1C1 reference"
    if name.casefold() in taken:
        return "another table or a defined name has it"
    return None


def suggest_name(name: Any, taken: "set[str] | frozenset[str]") -> str:
    """A table name close to ``name`` that :func:`table_name_problem` takes."""
    text = name if isinstance(name, str) else ""
    text = re.sub(r"[^\w\\.?]+", "_", text).strip("_")[:MAX_TABLE_NAME - 4] or "Table"
    if not _NAME_START.fullmatch(text[0]):
        text = "_" + text
    for candidate in (text, text + "_", "_" + text):
        if table_name_problem(candidate, taken) is None:
            return candidate
    stem = text if table_name_problem(text + "_1", set()) is None else "Table"
    if stem != "Table":
        stem += "_"
    number = 1
    while table_name_problem(f"{stem}{number}", taken) is not None:
        number += 1
    return f"{stem}{number}"


def _column_number(letters: str) -> int:
    number = 0
    for char in letters.upper():
        if not "A" <= char <= "Z":
            raise ValueError(letters)
        number = number * 26 + ord(char) - 64
    return number


def taken_names(workbook: Any) -> set[str]:
    """Every table's name and display name and every defined name, case folded."""
    out: set[str] = set()
    for table in workbook._table_parts():
        out.add(table.name.casefold())
        out.add(table.object_name.casefold())
    for name in workbook._defined_names():
        out.add(name.name.casefold())
    return out


# -- writing a table ----------------------------------------------------------------------------


def write_table(workbook: Any, sheet: Any, origin: Any, headers: "Sequence[Any] | None",
                rows: "Sequence[Any]", *, as_table: "bool | str" = False,
                name: str | None = None, style: str | None = None,
                overwrite: bool = False) -> TableWritten:
    """Write ``headers`` and ``rows`` at ``origin`` of ``sheet`` (see the module)."""
    from . import EditError

    problems: list[tuple[str, str, str | None]] = []
    target, cell = _origin(workbook, sheet, origin)
    label = format_sheet(target.name, target.id)
    if isinstance(as_table, str):
        if name is not None and name != as_table:
            raise EditError("bad_value", "as_table names the table and so does name: give "
                            "one", addresses=[label])
        name, as_table = as_table, True
    elif name is not None:
        as_table = True        # a name asks for a table (rule 1: the intent is clear)
    header = _headers(headers, label, bool(as_table))
    width = len(header) if header else max((_width(row) for row in rows), default=0)
    if width == 0:
        raise EditError("bad_value", "no columns: give headers or rows", addresses=[label])
    body = [_row(row, header, width, f"rows[{n}]") for n, row in enumerate(rows)]
    height = (1 if header else 0) + max(len(body), 1 if as_table else 0)
    if height == 0:
        raise EditError("bad_value", "no rows to write", addresses=[label])
    last = Cell(cell.row + height - 1, cell.column + width - 1)
    if last.row > MAX_ROW or last.column > MAX_COLUMN:
        raise EditError("bad_address", f"the block {height} rows by {width} columns from "
                        f"{cell.ref} runs past the sheet's last row or column",
                        addresses=[format_range(label, cell)])
    area = Area(cell, last)
    table_name = style_name = None
    if as_table:
        taken = taken_names(workbook)
        table_name = name if name is not None else _default_name(taken)
        problem = table_name_problem(table_name, taken)
        if problem is not None:
            problems.append(("bad_table_name", format_range(label, area),
                             f"{problem}; {suggest_name(table_name, taken)!r} would do"))
        style_name, problem = _style(workbook, style)
        if problem is not None:
            problems.append(("bad_table_style", format_range(label, area), problem))
        if _sheet_filter(workbook, target, area):
            problems.append(("table_overlap", format_range(label, area),
                             "the sheet's AutoFilter covers it: a table brings its own"))
    _block_problems(workbook, target, area, label, problems, empty=not overwrite)
    _raise(problems)
    items: list[tuple[str, Any]] = []
    for offset, text in enumerate(header):
        items.append((format_range(label, Cell(cell.row, cell.column + offset)), text))
    top = cell.row + (1 if header else 0)
    for r, values in enumerate(body):
        for c, value in enumerate(values):
            if value is None and not overwrite:
                continue
            items.append((format_range(label, Cell(top + r, cell.column + c)), value))
    editor = workbook.editor
    with editor.batch():
        if any(_holds_text(value) for _, value in items):
            ensure_shared_strings(workbook)
        _date_formats(workbook, target, items)
        written = workbook.set_cells(items, replace_formulas=overwrite) if items \
            else Written(cells=(), changed=0, calculation=workbook.calculation_status)
        if as_table:
            assert table_name is not None and style_name is not None
            _add_table(workbook, target, area, table_name, header, style_name)
            editor.touched(target, [])
            workbook._forget()
    return TableWritten(range=format_range(label, area), table=table_name,
                        header=format_range(label, Area(cell, Cell(cell.row, last.column)))
                        if header else None,
                        rows=len(body), written=written)


def _origin(workbook: Any, sheet: Any, origin: Any) -> tuple[Any, Cell]:
    """The sheet and first cell: ``sheet`` a name (or ``None`` when ``origin`` names one)
    and ``origin`` a cell, ``B2`` or ``Stock!B2`` (its sheet the same as ``sheet``)."""
    from . import EditError

    if not isinstance(origin, str) or not origin.strip():
        raise EditError("bad_address", "origin is a cell, as B2")
    text = origin.strip()
    named = None
    if "!" in text:
        try:
            parsed = parse_address(text)
        except (AddressError, ValueError):
            parsed = None
        if not isinstance(parsed, RangeAddress) or parsed.kind != "cell":
            raise EditError("bad_address", "origin is one cell, as B2 or Stock!B2",
                            addresses=[text[:80]])
        named = workbook.sheet_of(parsed)
        cell = parsed.area.first
    else:
        try:
            cell = Cell.parse(text.replace("$", ""))
        except ValueError:
            raise EditError("bad_address", "origin is one cell, as B2",
                            addresses=[text[:80]]) from None
    if sheet is None:
        if named is None:
            raise EditError("bad_address", "name the sheet, or give origin as Stock!B2")
        return named, cell
    try:
        target = workbook.sheet(sheet) if isinstance(sheet, str) else sheet
    except KeyError:
        raise EditError("bad_address", f"no sheet {sheet!r}; the sheets are "
                        f"{[s.name for s in workbook.sheets][:50]}") from None
    if named is not None and named.part != target.part:
        raise EditError("bad_address", f"origin names sheet {named.name!r}, not "
                        f"{target.name!r}", addresses=[text[:80]])
    return target, cell


def _headers(headers: Any, label: str, table: bool) -> list[str]:
    from . import EditError

    if headers is None or (isinstance(headers, (list, tuple)) and not headers):
        if table:
            raise EditError("bad_value", "a table needs its headers: its columns' names",
                            addresses=[label])
        return []
    if isinstance(headers, str) or not isinstance(headers, Sequence):
        raise EditError("bad_value", "headers are a list of texts", addresses=[label])
    out: list[str] = []
    seen: dict[str, int] = {}
    problems: list[tuple[str, str, str | None]] = []
    for n, text in enumerate(headers):
        where = f"headers[{n}]"
        if not isinstance(text, str) or not text.strip():
            problems.append(("bad_value", where, "a header is text that is not empty"))
            continue
        if table and ("\n" in text or "\r" in text):
            problems.append(("bad_value", where, "a table's column name is one line"))
            continue
        if table and text.casefold() in seen:
            problems.append(("bad_value", where, f"the same name as headers[{seen[text.casefold()]}]"
                             f" (case aside): a table's columns are named apart, e.g. "
                             f"{text + '2'!r}"))
            continue
        seen.setdefault(text.casefold(), n)
        out.append(text)
    _raise(problems)
    return out


def _width(row: Any) -> int:
    if isinstance(row, Mapping):
        return len(row)
    if isinstance(row, (str, bytes)) or not isinstance(row, Sequence):
        return 1
    return len(row)


def _row(row: Any, header: list[str], width: int, where: str) -> list[Any]:
    """``row`` as values in the columns' order, padded with blanks; ``""`` a blank."""
    from . import EditError

    if isinstance(row, Mapping) and not _is_value(row):
        if not header:
            raise EditError("bad_value", "a row by name needs headers", addresses=[where])
        index = {text.casefold(): n for n, text in enumerate(header)}
        values: list[Any] = [None] * width
        unknown = []
        for key, value in row.items():
            n = index.get(key.casefold()) if isinstance(key, str) else None
            if n is None:
                unknown.append(str(key)[:40])
                continue
            values[n] = value
        if unknown:
            raise EditError("bad_value", f"no column named {unknown[0]!r}; the columns are "
                            f"{header[:50]}", addresses=[where])
    elif isinstance(row, (str, bytes)) or not isinstance(row, Sequence):
        raise EditError("bad_value", "a row is a list of values, or a mapping by header",
                        addresses=[where])
    else:
        if len(row) > width:
            raise EditError("bad_value", f"{len(row)} values for {width} columns",
                            addresses=[where])
        values = list(row) + [None] * (width - len(row))
    return [None if isinstance(value, str) and value == "" else value for value in values]


def _is_value(given: Mapping) -> bool:
    """A mapping that is one cell's value (``{"date": ...}``), not a row by name."""
    return bool(set(given) & {"value", "formula", "date", "markdown"}) \
        and set(given) <= {"value", "formula", "date", "markdown", "format", "expect"}


def _holds_text(value: Any) -> bool:
    return isinstance(value, str) or (isinstance(value, Mapping) and "markdown" in value)


def _default_name(taken: set[str]) -> str:
    number = 1
    while f"table{number}" in taken:
        number += 1
    return f"Table{number}"


def _style(workbook: Any, style: Any) -> tuple[str, str | None]:
    """The table style to give: ``style`` checked, else the workbook's default."""
    root = _styles_root(workbook)
    node = root.find(_X + "tableStyles") if root is not None else None
    own = [s.get("name", "") for s in node.iterchildren(_X + "tableStyle")] \
        if node is not None else []
    if style is None:
        default = node.get("defaultTableStyle") if node is not None else None
        return (default or FALLBACK_TABLE_STYLE), None
    if isinstance(style, str) and (style in own or _BUILTIN_STYLE.fullmatch(style)):
        return style, None
    options = own + [FALLBACK_TABLE_STYLE, "TableStyleLight1", "TableStyleDark1"]
    return FALLBACK_TABLE_STYLE, (f"not a table style: one the workbook defines "
                                  f"({', '.join(own[:10]) or 'none'}) or a built-in, as "
                                  f"{', '.join(options[-3:])}")


def _styles_root(workbook: Any) -> Any:
    part = workbook._related("styles")
    return workbook.package.tree(part) if part else None


def _sheet_filter(workbook: Any, sheet: Any, area: Area) -> bool:
    root = workbook.package.tree(sheet.part)
    node = root.find(_X + "autoFilter") if root is not None else None
    if node is None:
        return False
    try:
        found = Area.parse(node.get("ref", ""), lenient=True)
    except ValueError:
        return False
    return _meets(found, area)


def _meets(a: Area, b: Area) -> bool:
    return not (a.last.row < b.first.row or b.last.row < a.first.row
                or a.last.column < b.first.column or b.last.column < a.first.column)


def _block_problems(workbook: Any, sheet: Any, area: Area, label: str,
                    problems: list, *, empty: bool, skip_table: str | None = None) -> None:
    """The cells of ``area`` holding a value or formula (when ``empty``), and the merged
    areas and tables (but ``skip_table``'s part) it meets."""
    for table in workbook._table_parts():
        if table.sheet == sheet.name and table.part != skip_table and table.area is not None \
                and _meets(table.area, area):
            problems.append(("table_overlap", format_range(label, table.area),
                             f"table {table.name} is there"))
    for merged in workbook.merges(sheet):
        if _meets(merged, area):
            problems.append(("merged_cell", format_range(label, merged),
                             "a merged area is there: unmerge it first"))
    if not empty:
        return
    held = [Cell(r, c) for r, c in _cells_with_content(workbook, sheet, area)]
    for found in held[:MAX_LISTED]:
        problems.append(("not_empty", format_range(label, found),
                         "it holds a value or a formula: overwrite=True writes over it"))


def _cells_with_content(workbook: Any, sheet: Any, area: Area) -> list[tuple[int, int]]:
    """The cells of ``area`` holding a value or a formula (a cell with only a format holds
    neither), as ``(row, column)`` in order: one streamed pass over the sheet."""
    out = []
    first, last = area.first, area.last
    for value in workbook.cells(sheet):
        if first.row <= value.row <= last.row and first.column <= value.column <= last.column \
                and (value.formula or value.t != "blank"):
            out.append((value.row, value.column))
    return sorted(out)


def _raise(problems: list) -> None:
    from . import EditError

    if not problems:
        return
    code, _, detail = problems[0]
    same = [address for c, address, _ in problems if c == code]
    others = len(problems) - len(same)
    if others:
        detail = (f"{detail}; " if detail else "") + f"{others} other problem(s)"
    raise EditError(code, detail, addresses=same,
                    problems=[(c, a) for c, a, _ in problems[:MAX_LISTED]])


def ensure_shared_strings(workbook: Any) -> bool:
    """Give the workbook a shared-string table when it has none (a new workbook's first
    text, as Excel adds one); whether one was added."""
    package = workbook.package
    if workbook._related("sharedStrings") is not None:
        return False
    part = "/xl/sharedStrings.xml"
    if package.has_part(part):
        part = package.unused_part_name("/xl/sharedStrings{n}.xml")
    package.add_part(part, (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                            f'<sst xmlns="{NS_MAIN}" count="0" uniqueCount="0"/>').encode(),
                     CT_SHARED_STRINGS, override=True)
    package.add_relationship(workbook.part, REL_SHARED_STRINGS, part)
    workbook._forget()
    return True


def _date_formats(workbook: Any, sheet: Any, items: list[tuple[str, Any]]) -> None:
    """Give each cell a date is written into the short date (:data:`SHORT_DATE`) when its
    format shows no date and the value names no format: as Excel formats a typed date."""
    styles = workbook.editor.cells.styles
    if styles is None:
        return
    for address, value in items:
        dated = isinstance(value, _dt.date) or (
            isinstance(value, Mapping) and "date" in value and value.get("format") is None)
        if not dated:
            continue
        parsed = parse_address(address)
        assert isinstance(parsed, RangeAddress)
        cell = parsed.area.first
        writer = workbook.editor.cells.sheet(sheet.stored_name)
        base = writer.style(cell, inherited=True)
        if base >= styles.count or workbook.styles.format(base).kind in (
                "date", "datetime", "time"):
            continue     # it shows a date; or a time, which set_cells says what to do with
        wanted = styles.with_number_format(base, SHORT_DATE)
        workbook.editor.write(sheet, cell, None, style=wanted)
        workbook._forget()


def _add_table(workbook: Any, sheet: Any, area: Area, name: str, header: list[str],
               style: str) -> str:
    """The table part over ``area`` (see the module), related to the sheet."""
    package = workbook.package
    ids = [0]
    for part in package.part_names:
        if package.content_type(part) == CT_TABLE:
            root = package.tree(part)
            raw = root.get("id", "") if root is not None else ""
            if raw.isdigit():
                ids.append(int(raw))
    part = package.unused_part_name("/xl/tables/table{n}.xml")
    root = etree.Element(_X + "table", nsmap={None: NS_MAIN})
    for key, value in (("id", str(max(ids) + 1)), ("name", name), ("displayName", name),
                       ("ref", area.ref), ("totalsRowShown", "0")):
        root.set(key, value)
    etree.SubElement(root, _X + "autoFilter").set("ref", area.ref)
    columns = etree.SubElement(root, _X + "tableColumns")
    columns.set("count", str(len(header)))
    for n, text in enumerate(header, 1):
        column = etree.SubElement(columns, _X + "tableColumn")
        column.set("id", str(n))
        column.set("name", text)
    info = etree.SubElement(root, _X + "tableStyleInfo")
    for key, value in (("name", style), ("showFirstColumn", "0"), ("showLastColumn", "0"),
                       ("showRowStripes", "1"), ("showColumnStripes", "0")):
        info.set(key, value)
    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    package.add_part(part, data, CT_TABLE, override=True)
    rel_id = package.add_relationship(sheet.part, REL_TABLE, part)
    sheet_root = package.tree(sheet.part)
    parts = sheet_root.find(_X + "tableParts")
    if parts is None:
        parts = etree.Element(_X + "tableParts")
        _insert_ordered(sheet_root, parts)
    etree.SubElement(parts, _X + "tablePart").set(f"{{{NS_R}}}id", rel_id)
    parts.set("count", str(len(parts)))
    package.mark_dirty(sheet.part)
    return part


def _insert_ordered(parent: Any, child: Any) -> None:
    """Put ``tableParts`` where CT_Worksheet has it (ECMA-376 Part 1 §18.3.1.99): the last
    child but ``extLst``."""
    ext = parent.find(_X + "extLst")
    if ext is not None:
        ext.addprevious(child)
    else:
        parent.append(child)


# -- appending rows -----------------------------------------------------------------------------


@dataclass
class _TableFacts:
    part: str
    root: Any
    name: str
    sheet: Any
    area: Area
    header_rows: int
    totals_rows: int
    columns: list[str]
    calculated: dict[int, str] = field(default_factory=dict)


def append_rows(workbook: Any, target: Any, rows: "Sequence[Any]") -> Appended:
    """Append ``rows`` to a table, a sheet or a region (see the module)."""
    from . import EditError

    if not isinstance(target, str) or not target.strip():
        raise EditError("bad_address", "the target is a table (t:Orders), a sheet (Log) or "
                        "a region (Log!B4:D200)")
    if isinstance(rows, (str, bytes)) or not isinstance(rows, Sequence) or not rows:
        raise EditError("bad_value", "rows is a list of rows, at least one",
                        addresses=[target.strip()[:80]])
    table, sheet, region = _target(workbook, target.strip())
    if table is not None:
        return _append_table(workbook, table, rows)
    return _append_sheet(workbook, sheet, region, rows)


def _target(workbook: Any, text: str) -> "tuple[_TableFacts | None, Any, Area | None]":
    from ..address import NameNotFound
    from . import EditError

    tables = {t.name.casefold(): t for t in workbook._table_parts()}
    if text.startswith("t:"):
        found = tables.get(text[2:].strip().casefold())
        if found is None:
            raise EditError("bad_address", f"no table {text[2:]!r}; the tables are "
                            f"{[t.name for t in workbook._table_parts()][:50]}",
                            addresses=[text[:80]])
        return _facts(workbook, found), None, None
    if "!" in text:
        try:
            parsed = parse_address(text)
        except (AddressError, ValueError):
            parsed = None
        if not isinstance(parsed, RangeAddress) or parsed.sheet is None:
            raise EditError("bad_address", "a region is a range with its sheet, as "
                            "Log!B4:D200 or Log!B:D", addresses=[text[:80]])
        try:
            sheet = workbook.sheet_of(parsed)
        except (NameNotFound, KeyError, ValueError):
            raise EditError("bad_address", f"no such sheet; the sheets are "
                            f"{[s.name for s in workbook.sheets][:50]}",
                            addresses=[text[:80]]) from None
        return None, sheet, None if parsed.span == "sheet" else parsed.area
    table = tables.get(text.casefold())
    try:
        sheet = workbook.sheet(text)
    except (NameNotFound, KeyError):
        sheet = None
    if table is not None and sheet is not None:
        raise EditError("bad_address", f"{text!r} is a table and a sheet: say which, "
                        f"t:{table.name} or {format_sheet(sheet.name)}!", addresses=[text[:80]])
    if table is not None:
        return _facts(workbook, table), None, None
    if sheet is None:
        raise EditError("bad_address", f"no table or sheet {text!r}; the tables are "
                        f"{[t.name for t in workbook._table_parts()][:25]}, the sheets "
                        f"{[s.name for s in workbook.sheets][:25]}", addresses=[text[:80]])
    return None, sheet, None


def _facts(workbook: Any, table: Any) -> _TableFacts:
    from . import EditError

    if table.area is None or table.kind != "worksheet":
        raise EditError("bad_address", f"table {table.name} cannot be grown here (its ref "
                        "is not a range, or it is a query or XML table)", addresses=[table.name])
    sheet = workbook.sheet(table.sheet)
    root = workbook.package.tree(table.part)
    calculated = {}
    container = root.find(_X + "tableColumns")
    for n, column in enumerate(container.findall(_X + "tableColumn") if container is not None
                               else []):
        formula = column.find(_X + "calculatedColumnFormula")
        if formula is not None and formula.text and formula.get("array") not in ("1", "true"):
            calculated[n] = formula.text
    return _TableFacts(part=table.part, root=root, name=table.name, sheet=sheet,
                       area=table.area, header_rows=table.header_rows,
                       totals_rows=table.totals_rows, columns=list(table.columns),
                       calculated=calculated)


def _append_table(workbook: Any, table: _TableFacts, rows: Sequence[Any]) -> Appended:
    from ..read.reach import references
    from . import EditError

    sheet, area = table.sheet, table.area
    label = format_sheet(sheet.name, sheet.id)
    width = area.columns
    body = [_row(row, table.columns, width, f"rows[{n}]") for n, row in enumerate(rows)]
    count = len(body)
    if area.last.row + count > MAX_ROW:
        raise EditError("bad_address", "the table would grow past the sheet's last row",
                        addresses=[format_range(label, area)])
    below = Area(Cell(area.last.row + 1, area.first.column),
                 Cell(area.last.row + count, area.last.column))
    problems: list[tuple[str, str, str | None]] = []
    _block_problems(workbook, sheet, below, label, problems, empty=True, skip_table=table.part)
    first_new = area.last.row - table.totals_rows + 1
    writer = workbook.editor.cells.sheet(sheet.stored_name)
    totals = None
    if table.totals_rows:
        totals = Area(Cell(area.last.row - table.totals_rows + 1, area.first.column), area.last)
        if sheet.protected:
            problems.append(("protected", format_range(label, totals), "the totals row moves "
                             "down, and the sheet is protected"))
        allowed = workbook._allowed()
        for r in range(totals.first.row, totals.last.row + 1):
            for c in range(totals.first.column, totals.last.column + 1):
                cell = Cell(r, c)
                formula = writer.content(cell).formula
                if formula and any(ref.kind == "cells" for ref in references(formula)):
                    problems.append(("totals_referenced", format_range(label, cell),
                                     "a totals cell's formula refers to cells by address, "
                                     "which would not take in the new rows"))
                if allowed is not None and not (allowed.allows(sheet.part, cell) and
                                                allowed.allows(sheet.part, Cell(r + count, c))):
                    problems.append(("not_writable", format_range(label, cell),
                                     "the totals row moves down"))
        for address in _totals_referenced(workbook, sheet, totals):
            problems.append(("totals_referenced", address, "it refers to the totals row by "
                             "address, which would not follow the row down"))
    _raise(problems)
    new_area = Area(area.first, Cell(area.last.row + count, area.last.column))
    filled: list[str] = []
    editor = workbook.editor
    with editor.batch():
        if totals is not None:
            moved = []
            for r in range(totals.last.row, totals.first.row - 1, -1):
                for c in range(totals.first.column, totals.last.column + 1):
                    source, to = Cell(r, c), Cell(r + count, c)
                    writer.move(source, to)
                    moved += [source, to]
            editor.cells.drop_calc_chain()
            editor.cells.request_full_calculation()
            editor.touched(sheet, moved)
            workbook._forget()
        # the format above each new cell, where the new cell has none
        last_data = first_new - 1
        styled = []
        for r in range(count):
            for c in range(width):
                at = Cell(first_new + r, area.first.column + c)
                above = writer.style(Cell(last_data, at.column)) if last_data > area.first.row \
                    + table.header_rows - 1 else 0
                if above and not writer.style(at):
                    styled.append((at, above))
        for at, above in styled:
            editor.write(sheet, at, None, style=above)
        if styled:
            workbook._forget()
        items: list[tuple[str, Any]] = []
        for r, values in enumerate(body):
            for c, value in enumerate(values):
                at = Cell(first_new + r, area.first.column + c)
                if value is None and c in table.calculated:
                    value = Formula(table.calculated[c])
                    if table.columns[c] not in filled:
                        filled.append(table.columns[c])
                if value is None:
                    continue
                items.append((format_range(label, at), value))
        if any(_holds_text(value) for _, value in items):
            ensure_shared_strings(workbook)
        _date_formats(workbook, sheet, items)
        written = workbook.set_cells(items) if items \
            else Written(cells=(), changed=0, calculation=workbook.calculation_status)
        _grow(workbook, table, new_area)
        editor.touched(sheet, [])
        workbook._forget()
    rows_area = Area(Cell(first_new, area.first.column),
                     Cell(first_new + count - 1, area.last.column))
    moved_totals = None
    if totals is not None:
        moved_totals = format_range(label, Area(Cell(totals.first.row + count, area.first.column),
                                                Cell(totals.last.row + count, area.last.column)))
    return Appended(range=format_range(label, rows_area), rows=count, table=table.name,
                    table_range=format_range(label, new_area), totals=moved_totals,
                    filled=tuple(filled), written=written)


def _grow(workbook: Any, table: _TableFacts, area: Area) -> None:
    """The table part's ``ref`` set to ``area``, its AutoFilter's to the same less the
    totals rows."""
    root = table.root
    root.set("ref", area.ref)
    auto = root.find(_X + "autoFilter")
    if auto is not None:
        last = Cell(area.last.row - table.totals_rows, area.last.column)
        auto.set("ref", Area(area.first, last).ref)
    workbook.package.mark_dirty(table.part)


def _totals_referenced(workbook: Any, sheet: Any, totals: Area) -> list[str]:
    """The cells (and defined names) whose formula refers to a cell of ``totals`` by
    address, a totals cell's own formula's references to its row too; conservative: a
    shared formula's references are taken across its whole group."""
    from ..read.reach import references

    out: list[str] = []
    for each in workbook._every_sheet:
        if each.part is None or each.type != "worksheet":
            continue
        root = workbook.package.tree(each.part)
        if root is None:
            continue
        own = each.part == sheet.part
        label = format_sheet(each.name, each.id)
        for node in root.iter(_X + "f"):
            text = node.text
            if not text:
                continue
            cell_node = node.getparent()
            ref = cell_node.get("r") or ""
            if own and ref:
                try:
                    if totals.contains(Cell.parse(ref)):
                        continue        # a totals cell's own formula: read on its own
                except ValueError:
                    pass
            spread = (0, 0)
            if node.get("t") == "shared" and node.get("ref"):
                try:
                    group = Area.parse(node.get("ref"), lenient=True)
                    spread = (group.rows - 1, group.columns - 1)
                except ValueError:
                    spread = (MAX_ROW, MAX_COLUMN)
            if _reaches(references(text), own, sheet.name, totals, spread):
                out.append(f"{label}!{ref}" if ref else label)
                if len(out) >= MAX_LISTED:
                    return out
    for name in workbook._defined_names():
        if name.formula and _reaches(references(name.formula), False, sheet.name, totals,
                                     (0, 0)):
            out.append(f"name {name.name}")
    return out


def _reaches(found: list, own: bool, sheet_name: str, totals: Area,
             spread: tuple[int, int]) -> bool:
    for reference in found:
        if reference.kind != "cells":
            continue
        sheets = reference.sheets
        if sheets is None:
            if not own:
                continue
        elif not any(s.casefold() == sheet_name.casefold() for s in sheets):
            continue
        first, last = reference.first, reference.last or reference.first
        if first is None or first.row is None and last.row is None:
            continue        # whole columns: a cell moved down its column stays in them
        low_row = min(_v(first.row, 1), _v(last.row, 1))
        high_row = max(_v(first.row, MAX_ROW), _v(last.row, MAX_ROW))
        low_col = min(_v(first.column, 1), _v(last.column, 1))
        high_col = max(_v(first.column, MAX_COLUMN), _v(last.column, MAX_COLUMN))
        if first.row_moves or last.row_moves:
            high_row += spread[0]
        if first.column_moves or last.column_moves:
            high_col += spread[1]
        if not (high_row < totals.first.row or low_row > totals.last.row
                or high_col < totals.first.column or low_col > totals.last.column):
            return True
    return False


def _v(value: int | None, default: int) -> int:
    return default if value is None else value


def _append_sheet(workbook: Any, sheet: Any, region: Area | None,
                  rows: Sequence[Any]) -> Appended:
    from . import EditError

    if sheet.type != "worksheet" or sheet.part is None:
        raise EditError("not_a_worksheet", addresses=[sheet.name])
    label = format_sheet(sheet.name, sheet.id)
    width = max(_width(row) for row in rows)
    first_column = region.first.column if region is not None else 1
    if region is not None and width > region.columns:
        raise EditError("bad_value", f"rows of {width} values for a region {region.columns} "
                        "columns wide", addresses=[format_range(label, region)])
    if first_column + width - 1 > MAX_COLUMN:
        raise EditError("bad_address", "the rows run past the sheet's last column",
                        addresses=[label])
    body = [_row(row, [], width, f"rows[{n}]") for n, row in enumerate(rows)]
    columns = (first_column, first_column + width - 1)
    low = region.first.row if region is not None else 1
    high = region.last.row if region is not None else MAX_ROW
    read = (first_column, region.last.column) if region is not None else columns
    scan = Area(Cell(low, read[0]), Cell(high, read[1]))
    held = _cells_with_content(workbook, sheet, scan)
    start = (max(r for r, _ in held) + 1) if held else low
    count = len(body)
    if start + count - 1 > MAX_ROW:
        raise EditError("bad_address", "the rows would run past the sheet's last row",
                        addresses=[label])
    block = Area(Cell(start, columns[0]), Cell(start + count - 1, columns[1]))
    problems: list[tuple[str, str, str | None]] = []
    _block_problems(workbook, sheet, block, label, problems, empty=True)
    _raise(problems)
    items = [(format_range(label, Cell(start + r, first_column + c)), value)
             for r, values in enumerate(body) for c, value in enumerate(values)
             if value is not None]
    editor = workbook.editor
    with editor.batch():
        if any(_holds_text(value) for _, value in items):
            ensure_shared_strings(workbook)
        _date_formats(workbook, sheet, items)
        written = workbook.set_cells(items) if items \
            else Written(cells=(), changed=0, calculation=workbook.calculation_status)
    return Appended(range=format_range(label, block), rows=count, written=written)


__all__ = ["Appended", "FALLBACK_TABLE_STYLE", "MAX_TABLE_NAME", "SHORT_DATE", "TableWritten",
           "append_rows", "ensure_shared_strings", "suggest_name", "table_name_problem",
           "taken_names", "write_table"]
