"""Tables declared from the caller's configuration, before ``describe``.

A workbook made from a template is known to its application: which sheets hold which
table, on which row its header is, which column is its key, what each column means, and
which columns a model should not see.  The application says so once, as plain data (its
own configuration, mapped to this), and the tables are then addressed, described, read,
queried and derived from like any other (``t:checks``)::

    wb.declare_tables({"tables": [
        {"name": "checks", "sheets": ["Area *"], "header_row": 3, "key": "Ref",
         "description": "One row per check, every area's sheet",
         "fields": [
             {"header": "Rating (1 to 5, in halves)", "id": "rating",
              "description": "The reviewer's rating"},
             {"header": "Reviewer notes", "hidden": True}]},
        {"name": "run", "type": "key_value", "sheets": "Run info"}]})
    wb.read("t:checks/C-07/rating")      # the value, with its own cell: 'Area 2'!D5
    wb.read("t:run/report_date/value")

Declarations are the application's configuration, never a model's: no tool takes them.
The tool layer takes them when it opens a workbook (``xlsx_agent.tools.format(tables=...)``
for every workbook the format opens, or ``xlsx_agent.tools.declare(session, ...)`` for one
session's), and :meth:`Workbook.declare_tables
<xlsx_agent.workbook.Workbook.declare_tables>` in Python.

**A declaration** is a mapping (a list of them, or ``{"tables": [...]}``):

* ``name`` (required): the table's name, ``t:<name>``; a name as ECMA-376 defines one
  (:func:`~xlsx_agent.address.is_name`), no Excel table's or workbook-wide defined
  name's.  Declaring a name again replaces the table (or a region declared with it).
* ``sheets`` (required): a sheet name, or a list of them; each an exact name (matched
  case aside, as Excel matches them) or a glob (``*``, ``?``, ``[...]``, which no sheet
  name can hold: ``Area *``).  The sheets are taken in the workbook's order (the tabs'),
  each once.  An exact name that is no sheet, or a sheet that holds no cells (a chart
  sheet), is refused; a glob skips sheets that hold no cells, and matching none is
  refused.
* ``type``: ``"rows"`` (the default: a row per item) or ``"key_value"`` (see below).
* ``header_row``: the header's row on every sheet (1 by default; 0 for none, the default
  of a key-value table).  The header is one row; its text names the columns.  Every sheet's
  header must be the same text, column by column, as the first sheet's: one that differs
  is refused, naming both cells.
* ``first_row``: the first data row (the row under the header; without a header, the first
  row of the sheet's used range).
* ``last_row``: the last data row on every sheet.  Without it, the data rows run down to the
  row before the first row blank across the columns (or the used range's end), as a
  detected region's do; when a sheet holds more below that blank row, within the columns,
  the table says so (``rows_below``, at the first such cell), and ``last_row`` takes it in.
* ``columns``: the sheet columns, ``"A:F"`` (by default, the first sheet's header row from
  its first cell holding something to its last; without a header, the used range's
  columns).
* ``key``: the key column (by header or id): ``t:checks/C-07/rating``.  Every data row,
  on every sheet, must have a key (text, or a whole number) that no other row has: a key
  two rows share, on one sheet or two, is refused naming both cells (never the key); a row
  without one, naming its cell.
* ``description``: what the table holds, shown by ``describe``.
* ``fields``: what is said of some columns, each a mapping with ``header`` (its header
  text, matched case aside) or ``column`` (its letter, ``"D"``), and any of ``id`` (the
  name the column is addressed and shown by: ``t:checks[rating]``; its header still
  finds it), ``description`` (shown by ``describe`` and ``xlsx_profile``) and ``hidden``
  (``true``: the column is hidden from the model).  Columns not listed keep their
  header's name.
* ``optional``: ``true`` when a workbook without the sheets is no error: the table is then
  not declared.

**A table over several sheets** reads as one: its data rows are each sheet's, sheet by
sheet in the workbook's order, its header the first sheet's.  It is one rectangle of its
own (:class:`~xlsx_agent.read.regions.Span`), and every value read from it -- by
``xlsx_read``, ``xlsx_query``, ``describe``'s samples, a derive's sources and left-out
cells, a dataset's sources, a comparison, a check -- carries its own cell's address
(``'Area 2'!D5``); a row by position counts across the sheets (``t:checks/#12``), and a
row by key finds its sheet.  Its range is each sheet's range, joined by ``,``.  **The
sheet is not a column**: a value is a cell, and a sheet's name is no cell, so there is no
``_sheet`` column to group or filter by.  Every address names its sheet; for one sheet's
figures, declare that sheet's table as well (``{"name": "checks_north", "sheets": "Area
North", ...}``).

**Hidden from the model**: a hidden column is left out of the table -- ``describe``,
``xlsx_read``, ``xlsx_query``, ``xlsx_derive``, ``xlsx_profile``, ``xlsx_compare``,
``xlsx_validate_data`` and a deck's or document's ``from`` never show it, and an address
that names it is refused as naming no column (its name is never offered).  Its cells are
withheld from a read of the sheet's cells (``xlsx_read`` of ``Area 1!A1:H40`` gives them as
``withheld``), from ``find_text``, and from a region detected or declared over them.  A
column Excel hides (a hidden column of the sheet) stays a column like any other unless
it is declared hidden; its values say ``hidden``, as always.

**A key-value table** (``"type": "key_value"``) is a field column and a value column, and
any more (a confidence, a source...): ``t:run/report_date/value`` is the value of the row
whose field is ``report_date``.  Its first two columns are called ``field`` and ``value``
(an ``id`` renames them; a header's text finds them too), the rest by their header or
letter; its key is its first column.  ``describe`` shows it compactly, as its fields and
values.  A block of two columns with no header, whose first column is a different text
on every row (two rows at least) and whose first row holds something other than text, is
detected as one (``t:'Run info'!A1/report_date/value``); a wider one is declared, since a
first row of numbers over more columns is as likely a header of years.

**Validation**: a declaration that is not one of these, or that the workbook does not
fit, is refused with :class:`DeclarationError` -- its ``code``, ``path``
(``tables[0].fields[2].header``) and ``valid_options`` -- before any table is kept, so a
list of declarations is kept whole or not at all.  Never guessed past, and never with a
cell's content in the message.
"""

from __future__ import annotations

import fnmatch
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Mapping, Sequence

from ooxml_edit.charts.workbook import Area, Cell, column_index, column_letters

from .address import NameNotFound, RangeAddress, SheetRef, is_name, nearest, quote_sheet
from .read.regions import KEY_VALUE_COLUMNS, Part, Region, Span, _header_names
from .read.types import Problem

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: What a declared table is.
TYPES = ("rows", "key_value")

#: The keys a table's declaration takes.
TABLE_KEYS = ("name", "sheets", "type", "header_row", "first_row", "last_row", "columns",
              "key", "description", "fields", "optional")

#: The keys a field takes.
FIELD_KEYS = ("header", "column", "id", "description", "hidden")

#: The longest description (characters) a table or column takes: ``describe`` shows it.
MAX_DESCRIPTION = 300

#: The longest column id (characters).
MAX_ID = 64

#: The characters that make a sheet name a glob: no sheet name holds them.
GLOB = "*?["


class DeclarationError(ValueError):
    """A declaration refused: ``code`` says why (``bad_declaration``, ``unknown_sheet``,
    ``no_sheets``, ``not_cells``, ``empty_sheet``, ``name_taken``, ``bad_name``,
    ``duplicate_name``, ``header_mismatch``, ``no_header``, ``bad_columns``, ``bad_rows``,
    ``no_rows``, ``unknown_column``, ``duplicate_field``, ``bad_id``, ``id_taken``,
    ``no_columns``, ``hidden_key``, ``no_key``, ``duplicate_key``), ``path`` where in the
    declarations (``tables[0].fields[2].header``), ``valid_options`` what would do instead.
    The message names sheets, cells, paths and what the declaration gave, never a cell's
    content."""

    def __init__(self, code: str, path: str, message: str, *,
                 valid_options: Sequence[str] = ()) -> None:
        super().__init__(f"declaration {path}: {message}" if path else f"declaration: {message}")
        self.code = code
        self.path = path
        self.message = str(self)
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class FieldSpec:
    """One field of a declaration, checked for its form."""

    path: str
    header: str | None = None
    column: int | None = None
    id: str | None = None
    description: str | None = None
    hidden: bool = False


@dataclass(frozen=True)
class TableSpec:
    """One table's declaration, checked for its form (not yet against a workbook)."""

    path: str
    name: str
    sheets: tuple[str, ...]
    type: str = "rows"
    header_row: int = 1
    first_row: int | None = None
    last_row: int | None = None
    columns: tuple[int, int] | None = None
    key: str | None = None
    description: str | None = None
    fields: tuple[FieldSpec, ...] = ()
    optional: bool = False


# ------------------------------------------------------------------------------------------
# The form
# ------------------------------------------------------------------------------------------


def parse_declarations(declarations: Any) -> list[TableSpec]:
    """``declarations`` (a mapping, a list of them, or ``{"tables": [...]}``) as
    :class:`TableSpec`, checked for their form; :class:`DeclarationError` otherwise."""
    if isinstance(declarations, Mapping) and "tables" in declarations:
        extra = sorted(set(declarations) - {"tables"})
        if extra:
            raise DeclarationError("bad_declaration", extra[0], "unknown key; the "
                                   "declarations are {\"tables\": [...]}",
                                   valid_options=["tables"])
        items, base = declarations["tables"], "tables"
    elif isinstance(declarations, Mapping):
        items, base = [declarations], "tables"
    else:
        items, base = declarations, "tables"
    if isinstance(items, (str, bytes)) or not isinstance(items, Sequence):
        raise DeclarationError("bad_declaration", base, "a list of table declarations")
    specs = [_table(item, f"{base}[{n}]") for n, item in enumerate(items)]
    seen: dict[str, str] = {}
    for spec in specs:
        folded = spec.name.casefold()
        if folded in seen:
            raise DeclarationError("duplicate_name", f"{spec.path}.name", f"t:{spec.name} is "
                                   f"declared twice (first at {seen[folded]})")
        seen[folded] = spec.path
    return specs


def _table(item: Any, path: str) -> TableSpec:
    if not isinstance(item, Mapping):
        raise DeclarationError("bad_declaration", path, "a table's declaration is a mapping "
                               "with name, sheets and more", valid_options=list(TABLE_KEYS))
    for key in item:
        if key not in TABLE_KEYS:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "unknown key",
                                   valid_options=nearest(str(key), TABLE_KEYS))
    for key in ("name", "sheets"):
        if key not in item:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "is required")
    name = item["name"]
    if not is_name(name):
        raise DeclarationError("bad_name", f"{path}.name", "a table's name is a name as "
                               "ECMA-376 defines one: a letter, _ or \\ first, then letters, "
                               "digits, _ and .; not TRUE, FALSE or a cell reference",
                               valid_options=["checks", "run_info"])
    sheets = item["sheets"]
    if isinstance(sheets, str):
        sheets = [sheets]
    if isinstance(sheets, (bytes, Mapping)) or not isinstance(sheets, Sequence) or not sheets \
            or not all(isinstance(s, str) and s for s in sheets):
        raise DeclarationError("bad_declaration", f"{path}.sheets", "a sheet name or glob "
                               "(Area *), or a list of them")
    kind = item.get("type", "rows")
    if kind not in TYPES:
        raise DeclarationError("bad_declaration", f"{path}.type", "unknown table type",
                               valid_options=list(TYPES))
    header_row = _row(item, "header_row", path, minimum=0,
                      default=1 if kind == "rows" else 0)
    first_row = _row(item, "first_row", path)
    last_row = _row(item, "last_row", path)
    if first_row is not None and header_row and first_row <= header_row:
        raise DeclarationError("bad_rows", f"{path}.first_row", "the first data row is "
                               "below the header row")
    top = first_row or header_row + 1
    if last_row is not None and last_row < top - 1:
        raise DeclarationError("bad_rows", f"{path}.last_row", "the last data row is at or "
                               "below the first (one above it: no data rows)")
    columns = None
    if "columns" in item:
        columns = _columns(item["columns"], f"{path}.columns")
    key = item.get("key")
    if key is not None and not (isinstance(key, str) and key):
        raise DeclarationError("bad_declaration", f"{path}.key", "a column's header or id")
    description = _description(item, path)
    fields = item.get("fields", ())
    if isinstance(fields, (str, bytes, Mapping)) or not isinstance(fields, Sequence):
        raise DeclarationError("bad_declaration", f"{path}.fields", "a list of fields")
    optional = item.get("optional", False)
    if not isinstance(optional, bool):
        raise DeclarationError("bad_declaration", f"{path}.optional", "true or false")
    return TableSpec(path=path, name=name, sheets=tuple(sheets), type=kind,
                     header_row=header_row, first_row=first_row, last_row=last_row,
                     columns=columns, key=key, description=description,
                     fields=tuple(_field(f, f"{path}.fields[{n}]") for n, f in enumerate(fields)),
                     optional=optional)


def _field(item: Any, path: str) -> FieldSpec:
    if not isinstance(item, Mapping):
        raise DeclarationError("bad_declaration", path, "a field is a mapping with header "
                               "or column, and id, description or hidden",
                               valid_options=list(FIELD_KEYS))
    for key in item:
        if key not in FIELD_KEYS:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "unknown key",
                                   valid_options=nearest(str(key), FIELD_KEYS))
    if ("header" in item) == ("column" in item):
        raise DeclarationError("bad_declaration", path, "a field names its column by header "
                               "or by column (its letter), one of the two",
                               valid_options=["header", "column"])
    header = item.get("header")
    if header is not None and not (isinstance(header, str) and header):
        raise DeclarationError("bad_declaration", f"{path}.header", "the header's text")
    column = None
    if "column" in item:
        column = _letter(item["column"], f"{path}.column")
    ident = item.get("id")
    if ident is not None and not (isinstance(ident, str) and ident and len(ident) <= MAX_ID
                                  and ident == ident.strip() and "\n" not in ident):
        raise DeclarationError("bad_id", f"{path}.id", f"a column id is a name of 1 to {MAX_ID} "
                               "characters, no space at its ends", valid_options=["rating"])
    hidden = item.get("hidden", False)
    if not isinstance(hidden, bool):
        raise DeclarationError("bad_declaration", f"{path}.hidden", "true or false")
    return FieldSpec(path=path, header=header, column=column, id=ident,
                     description=_description(item, path), hidden=hidden)


def _description(item: Mapping, path: str) -> str | None:
    text = item.get("description")
    if text is not None and not (isinstance(text, str) and text.strip()
                                 and len(text) <= MAX_DESCRIPTION):
        raise DeclarationError("bad_declaration", f"{path}.description", "a description is "
                               f"text of 1 to {MAX_DESCRIPTION} characters")
    return text


def _row(item: Mapping, key: str, path: str, *, minimum: int = 1,
         default: int | None = None) -> int | None:
    value = item.get(key, default)
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= 1048576:
        raise DeclarationError("bad_rows", f"{path}.{key}", f"a row number from {minimum}")
    return value


def _letter(value: Any, path: str) -> int:
    try:
        if isinstance(value, str) and value.isalpha() and value.isascii():
            number = column_index(value.upper())
            if 1 <= number <= 16384:
                return number
    except (ValueError, KeyError):
        pass
    raise DeclarationError("bad_columns", path, "a column's letters, A to XFD",
                           valid_options=["D"])


def _columns(value: Any, path: str) -> tuple[int, int]:
    if not isinstance(value, str) or value.count(":") != 1:
        raise DeclarationError("bad_columns", path, "the columns as letters, A:F",
                               valid_options=["A:F"])
    first, last = (_letter(part.strip(), path) for part in value.split(":"))
    if first > last:
        raise DeclarationError("bad_columns", path, "the first column comes first, A:F",
                               valid_options=["A:F"])
    return first, last


# ------------------------------------------------------------------------------------------
# Against the workbook
# ------------------------------------------------------------------------------------------


def declare_tables(workbook: "Workbook", declarations: Any) -> list[Region]:
    """Declare the tables of ``declarations`` on ``workbook`` (see the module's notes):
    every one checked against the workbook first, then all kept; the tables declared, in
    order (an optional one whose sheets are not there left out)."""
    specs = declarations if _specs(declarations) else parse_declarations(declarations)
    built = [build(workbook, spec) for spec in specs]
    return [workbook.regions.keep_table(region) for region in built if region is not None]


def _specs(value: Any) -> bool:
    return isinstance(value, (list, tuple)) and bool(value) \
        and all(isinstance(v, TableSpec) for v in value)


def build(workbook: "Workbook", spec: TableSpec) -> Region | None:
    """The table ``spec`` declares, checked against ``workbook``; ``None`` for an optional
    one whose sheets are not there."""
    regions = workbook.regions
    path = spec.path
    folded = spec.name.casefold()
    taken = [t.address.text for t in regions.tables if t.name.casefold() == folded]
    taken += [n.address.text for n in regions.names if n.address is not None
              and n.scope is None and n.name.casefold() == folded]
    if taken:
        raise DeclarationError("name_taken", f"{path}.name", f"{taken[0]} is already a table "
                               "or defined name; a declared table takes a name of its own")
    sheets = _sheets(workbook, spec)
    if not sheets:
        return None
    first_sheet = sheets[0]
    used = workbook.used_range(first_sheet)
    left, right = _extent(workbook, spec, first_sheet, used)
    parts: list[Part] = []
    flags: list[Problem] = []
    belows: list[dict[int, tuple[int, str]]] = []
    names: list[str] | None = None
    texts: list[str] | None = None
    for sheet in sheets:
        used = workbook.used_range(sheet)
        if used is None:
            raise DeclarationError("empty_sheet", f"{path}.sheets", f"sheet "
                                   f"{quote_sheet(sheet.name)} holds no cells")
        header = spec.header_row or None
        written = quote_sheet(sheet.name)
        if header is not None:
            if header > used.last.row:
                raise DeclarationError("no_header", f"{path}.header_row", f"row {header} of "
                                       f"sheet {written} is below its last cell")
            row = _row_values(workbook, sheet, header, left, right)
            found = [v.v if v.t == "text" and isinstance(v.v, str) else "" for v in row]
            if texts is None:
                texts = found
                names = _header_names(row, left, flags, written)
            elif found != texts:
                at = next(i for i, (a, b) in enumerate(zip(found, texts)) if a != b)
                cell = f"{column_letters(left + at)}{header}"
                raise DeclarationError("header_mismatch", f"{path}.sheets", f"the header of "
                                       f"sheet {written} differs from sheet "
                                       f"{quote_sheet(first_sheet.name)}'s at {cell}: a table "
                                       "over several sheets has one header",
                                       valid_options=[f"{written}!{cell}",
                                                      f"{quote_sheet(first_sheet.name)}!{cell}"])
        first = spec.first_row or (header + 1 if header is not None else used.first.row)
        last, below = _last_row(workbook, spec, sheet, used, first, left, right)
        belows.append(below)
        parts.append(Part(sheet.name, sheet.id, header, first, last))
    if names is None:
        names = [f"Column {column_letters(c)}" for c in range(left, right + 1)]
    if spec.type == "key_value":
        if right - left < 1:
            raise DeclarationError("bad_columns", f"{path}.columns", "a key-value table has "
                                   "two columns at least: a field and a value")
        if spec.header_row == 0:
            names[:2] = list(KEY_VALUE_COLUMNS)
    return _region(workbook, spec, parts, left, right, names, flags, belows)


def _sheets(workbook: "Workbook", spec: TableSpec) -> "list[Sheet]":
    path = f"{spec.path}.sheets"
    every = workbook.sheets
    chosen: set[int] = set()
    cells = ("worksheet", "macrosheet")
    for n, wanted in enumerate(spec.sheets):
        where = path if len(spec.sheets) == 1 else f"{path}[{n}]"
        if any(ch in wanted for ch in GLOB):
            pattern = wanted.casefold()
            found = [i for i, s in enumerate(every) if s.part is not None and s.type in cells
                     and fnmatch.fnmatchcase(s.name.casefold(), pattern)]
            if not found:
                if spec.optional:
                    continue
                raise DeclarationError("no_sheets", where, "the glob matches no sheet that "
                                       "holds cells", valid_options=[s.name for s in every][:50])
            chosen.update(found)
            continue
        try:
            sheet = workbook.sheet(wanted)
        except NameNotFound as error:
            if spec.optional:
                continue
            raise DeclarationError("unknown_sheet", where, f"no sheet {wanted!r}",
                                   valid_options=error.valid_options) from None
        if sheet.part is None or sheet.type not in cells:
            raise DeclarationError("not_cells", where, f"sheet {quote_sheet(sheet.name)} is "
                                   f"a {sheet.type or 'sheet of no known type'}: it holds no cells")
        chosen.add(next(i for i, s in enumerate(every) if s is sheet or s == sheet))
    return [every[i] for i in sorted(chosen)]


def _extent(workbook: "Workbook", spec: TableSpec, sheet: "Sheet",
            used: Area | None) -> tuple[int, int]:
    """The table's sheet columns: declared, or the first sheet's header's (its first cell
    holding something to its last), or its used range's."""
    if used is None:
        raise DeclarationError("empty_sheet", f"{spec.path}.sheets", f"sheet "
                               f"{quote_sheet(sheet.name)} holds no cells")
    if spec.columns is not None:
        return spec.columns
    if spec.header_row and spec.header_row <= used.last.row:
        row = _row_values(workbook, sheet, spec.header_row, used.first.column, used.last.column)
        filled = [v.column for v in row if v.t != "blank"]
        if not filled:
            raise DeclarationError("no_header", f"{spec.path}.header_row", f"row "
                                   f"{spec.header_row} of sheet {quote_sheet(sheet.name)} is "
                                   "blank; give the row of the header, or columns")
        return min(filled), max(filled)
    return used.first.column, used.last.column


def _row_values(workbook: "Workbook", sheet: "Sheet", row: int, left: int, right: int) -> list:
    return next(iter(workbook.read(RangeAddress(SheetRef(name=sheet.name),
                                                Area(Cell(row, left), Cell(row, right))))))


def _last_row(workbook: "Workbook", spec: TableSpec, sheet: "Sheet", used: Area, first: int,
              left: int, right: int) -> tuple[int, dict[int, tuple[int, str]]]:
    """The last data row of ``sheet``, and in each column the first cell holding something
    below the blank row that ended the rows, by sheet column: its row and address (none:
    nothing there, or ``last_row`` declared)."""
    if spec.last_row is not None:
        return min(spec.last_row, max(used.last.row, first - 1)), {}
    if first > used.last.row:
        return first - 1, {}
    last = first - 1
    rows = workbook.read(RangeAddress(SheetRef(name=sheet.name),
                                      Area(Cell(first, left), Cell(used.last.row, right))))
    for values in rows:
        if all(v.t == "blank" for v in values):
            break
        last = values[0].row
    below: dict[int, tuple[int, str]] = {}
    for values in rows:
        for value in values:
            if value.t != "blank" and value.column not in below:
                below[value.column] = (value.row, value.a)
    return last, below


def _region(workbook: "Workbook", spec: TableSpec, parts: list[Part], left: int, right: int,
            names: list[str], flags: list[Problem],
            belows: "list[dict[int, tuple[int, str]]]") -> Region:
    path = spec.path
    width = right - left + 1
    ids: list[str | None] = [None] * width
    notes: list[str | None] = [None] * width
    hidden = [False] * width
    if spec.type == "key_value" and spec.header_row:
        ids[:2] = list(KEY_VALUE_COLUMNS)
    seen: dict[int, str] = {}
    for field in spec.fields:
        if field.column is not None:
            offset = field.column - left
            if not 0 <= offset < width:
                raise DeclarationError("unknown_column", f"{field.path}.column", f"column "
                                       f"{column_letters(field.column)} is outside the "
                                       f"table's, {column_letters(left)}:{column_letters(right)}")
        else:
            offset = next((i for i, name in enumerate(names)
                           if name.casefold() == field.header.casefold()), None)
            if offset is None:
                shown = [n for i, n in enumerate(names) if not _hidden_by(spec, i, left, names)]
                raise DeclarationError("unknown_column", f"{field.path}.header",
                                       f"no column has the header {field.header!r}",
                                       valid_options=nearest(field.header, shown))
        if offset in seen:
            raise DeclarationError("duplicate_field", field.path, f"column "
                                   f"{column_letters(left + offset)} is declared twice "
                                   f"(first at {seen[offset]})")
        seen[offset] = field.path
        if field.id is not None:
            ids[offset] = field.id
        notes[offset] = field.description
        hidden[offset] = field.hidden
    columns = [ids[i] or names[i] for i in range(width)]
    visible = [i for i in range(width) if not hidden[i]]
    if not visible:
        raise DeclarationError("no_columns", f"{path}.fields", "every column is hidden")
    taken: dict[str, int] = {}
    for i in visible:
        for name in dict.fromkeys([columns[i].casefold(), names[i].casefold()]):
            if name in taken and taken[name] != i:
                which = spec.fields[_field_of(spec, i, left, names)].path \
                    if _field_of(spec, i, left, names) is not None else f"{path}.fields"
                raise DeclarationError("id_taken", f"{which}.id", f"{columns[i]!r} is another "
                                       "column's name or header: a column is found by "
                                       "either, so each names one column")
            taken[name] = i
    if spec.type == "key_value" and (hidden[0] or hidden[1]):
        raise DeclarationError("hidden_key", f"{path}.fields", "a key-value table's field "
                               "and value columns are never hidden")
    key = None
    if spec.key is not None:
        wanted = spec.key.casefold()
        found = next((i for i in range(width)
                      if wanted in (columns[i].casefold(), names[i].casefold())), None)
        if found is None or hidden[found]:
            if found is not None:
                raise DeclarationError("hidden_key", f"{path}.key", "the key column is hidden "
                                       "from the model; keys are in addresses")
            raise DeclarationError("unknown_column", f"{path}.key", f"no column "
                                   f"{spec.key!r}", valid_options=nearest(
                                       spec.key, [columns[i] for i in visible]))
        key = columns[found]
    elif spec.type == "key_value":
        key = columns[0]
    sheet_columns = tuple(left + i for i in visible)
    # What the table says of its cells names no cell of a column hidden from the model:
    # a header flag there is dropped, and more below is flagged at its first visible cell.
    flags = [f for f in flags if f.address is None
             or _column_of(f.address) in (None, *sheet_columns)]
    for below in belows:
        shown = sorted((row, column, address) for column, (row, address) in below.items()
                       if column in sheet_columns)
        if shown:
            flags.append(Problem("rows_below", shown[0][2]))
    header_rows = 1 if spec.header_row else 0
    common = dict(kind="declared", name=spec.name, columns=tuple(columns[i] for i in visible),
                  header_rows=header_rows, key=key, flags=tuple(dict.fromkeys(flags)),
                  headers=tuple(names[i] for i in visible),
                  notes=tuple(notes[i] for i in visible), about=spec.description,
                  shape=spec.type, workbook=workbook)
    hidden_columns = tuple(left + i for i in range(width) if hidden[i])
    if len(parts) == 1 and not hidden_columns:
        part = parts[0]
        top = part.header if part.header is not None else part.first
        if part.last < top:
            raise DeclarationError("no_rows", path, f"sheet {quote_sheet(part.sheet)} has no "
                                   "data rows, and the table no header row")
        region = Region(sheet=part.sheet, sheet_id=part.sheet_id,
                        area=Area(Cell(top, left), Cell(max(part.last, top), right)), **common)
    else:
        if not header_rows and not any(p.rows for p in parts):
            raise DeclarationError("no_rows", path, "the sheets have no data rows, and the "
                                   "table no header row")
        span = Span(name=f"t:{spec.name}", parts=tuple(parts), columns=sheet_columns,
                    header_rows=header_rows, left=left, right=right, hidden=hidden_columns)
        region = Region(sheet=span.name, area=span.area, span=span, **common)
    if key is not None:
        if region.span is not None:
            with workbook.regions.reading(region.span):
                _check_keys(workbook, region, path)
        else:
            _check_keys(workbook, region, path)
    return region


def _column_of(address: str) -> int | None:
    """The sheet column of a cell's address (``'Area 1'!E3``), ``None`` for another form."""
    letters = address.rpartition("!")[2].rstrip("0123456789").lstrip("$")
    try:
        return column_index(letters) if letters.isalpha() else None
    except (ValueError, KeyError):
        return None


def _hidden_by(spec: TableSpec, offset: int, left: int, names: list[str]) -> bool:
    found = _field_of(spec, offset, left, names)
    return found is not None and spec.fields[found].hidden


def _field_of(spec: TableSpec, offset: int, left: int, names: list[str]) -> int | None:
    for n, field in enumerate(spec.fields):
        if field.column is not None and field.column - left == offset:
            return n
        if field.header is not None and field.header.casefold() == names[offset].casefold():
            return n
    return None


def _check_keys(workbook: "Workbook", region: Region, path: str) -> None:
    """Every data row has a key no other row has: refused otherwise, naming the cells."""
    index = workbook.regions.key_index(region)
    offset = region.column_index(region.key)

    def cell(position: int) -> str:
        row = region.data.first.row + position - 1
        column = region.area.first.column + offset
        return region.located(Area(Cell(row, column), Cell(row, column)))

    missing = [p for p, k in enumerate(index.keys, 1) if k is None]
    if missing:
        raise DeclarationError("no_key", f"{path}.key", f"{len(missing)} rows have no key in "
                               f"column {region.key!r} (blank, or not text or a whole "
                               f"number), the first at {cell(missing[0])}",
                               valid_options=[cell(p) for p in missing[:50]])
    if index.duplicates:
        rows = next(iter(index.duplicates.values()))
        count = len(index.duplicates)
        which = "a key" if count == 1 else f"{count} keys"
        are = "is" if count == 1 else "are"
        raise DeclarationError("duplicate_key", f"{path}.key", f"{which} of column "
                               f"{region.key!r} {are} on more than one row: "
                               f"{cell(rows[0])} and {cell(rows[1])} hold the same",
                               valid_options=[cell(p) for p in rows[:50]])


__all__ = ["DeclarationError", "FieldSpec", "MAX_DESCRIPTION", "MAX_ID", "TableSpec", "TYPES",
           "build", "declare_tables", "parse_declarations"]
