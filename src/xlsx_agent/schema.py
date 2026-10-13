"""Checking a workbook against a declared schema (ROADMAP.md, "Tools": ``xlsx_validate_data``)::

    report = wb.validate_data({
        "sheets": ["Stock"],
        "tables": [{"table": "t:Products", "key": ["SKU"],
                    "columns": [{"name": "SKU", "type": "text", "pattern": "SKU-\\\\d{4}"},
                                {"name": "Price", "type": "number", "min": 0},
                                {"name": "Score", "type": "number", "min": 1, "max": 5,
                                 "step": 0.5},
                                {"name": "Category", "enum": ["Hardware", "Garden"]},
                                {"name": "Status", "validation_list": True}]}]})
    report.ok            # False
    report.counts        # {"wrong_type": 2, "not_in_enum": 1}: cells, by code
    list(report)         # [DataProblem(code='wrong_type', address='Stock!C5:C6', ...)]

**The schema** is plain data (JSON writes it), a mapping with two keys, both optional:

* ``sheets``: names of sheets that must be there (matched case aside, as Excel matches
  them).  One that is not is ``missing_sheet``.
* ``tables``: the tables and regions to check, each a mapping with

  - ``table``: a ``t:`` address of an Excel table, a defined name, a declared region or a
    detected one (``t:Products``, ``t:Stock!Prices``, ``t:Stock!A1``); **or** a declared
    region spec instead, as :meth:`Workbook.declare_region
    <xlsx_agent.workbook.Workbook.declare_region>` takes one: ``range``
    (``Returns!A3:F200``), ``name`` (``Returns``), ``header_rows`` (default 1) and
    ``headers`` (one name per column; required with two or more header rows).  The spec is
    used for this check only, never kept in the workbook.  A table that is not there, or
    that the workbook cannot give (a defined name that is a formula, a header that is not
    detected...), is ``missing_table`` with the ``reason`` (``not_found``,
    ``sheet_not_found``, ``no_cells``, or the :class:`~xlsx_agent.read.regions.RegionError`
    code);
  - ``key``: a column name, or several: the row's key, never blank (``required_blank``)
    and unique together (``duplicate_key``);
  - ``columns``: the columns to check, each a mapping with ``name`` (matched case aside;
    a column that is not there is ``missing_column``) and any of the checks below.

**The checks** of a column, on each of its data cells (header and totals rows aside):

* ``type``: one of the reader's types (:data:`~xlsx_agent.read.types.TYPES`, ``blank``
  aside: ``number``, ``percent``, ``date``, ``time``, ``datetime``, ``text``, ``bool``,
  ``error``).  A cell of another type is ``wrong_type``, with the type found and the
  reader's issue (``number_as_text``...); nothing is coerced, so a percent is not a
  number and a number stored as text is not one either (ROADMAP.md, "Design rules", 4).
  A cell of the wrong type gets no other check.
* ``required``: a blank cell is ``required_blank``.  A formula without a cached value is
  ``no_cached_value`` in any column checked, required or not: its value is unknown.
* ``min`` and ``max``: the least and the most a value may be, written as a filter's value
  is (:mod:`xlsx_agent.query`: ``5``, ``0.25`` for 25 %, ``2026-01-05``, ``10:30``; in
  Python a number too); for a ``number``, ``percent``, ``date``, ``time`` or ``datetime``
  column (``type`` is then required), compared as a query compares
  (``out_of_range``).
* ``step``: values are ``min`` (0 without one) plus a whole number of steps -- ``1`` for
  whole numbers, ``0.5`` for halves, as a score of 1 to 5 in halves -- for a ``number`` or
  ``percent`` column; compared as decimals, each number as its shortest decimal form
  (``off_step``).
* ``enum``: the values allowed (at most :data:`MAX_ENUM`, written as ``min`` is); each cell
  must equal one, as its own type reads it (``not_in_enum``).  ``case`` says how text is
  compared: ``exact`` (the default: case counts, nothing trimmed, as a query's ``eq``)
  or ``ignore`` (:meth:`str.casefold` on both sides).  When a text cell would match but
  for case or for space at its ends, the problem says so (``hint``: ``case_differs``,
  ``space_differs``), without the value.
* ``validation_list``: the cell must be one of the values of the workbook's own
  data-validation lists over the column (each list that covers any of its cells: an
  inline list's values as written, or its source range's, as the profile reads them,
  :mod:`xlsx_agent.profile`), compared as ``enum`` is, ``case`` too (``not_in_list``).  A
  column with no such list, or whose list is not readable (a source that is a formula;
  more than :data:`~xlsx_agent.profile.ALLOWED_CAP` values), is ``no_validation_list``
  with the reason, and its cells are not checked against one.  Not with ``enum``.
* ``pattern``: a Python regular expression a text cell must match whole
  (:func:`re.fullmatch`; ``pattern_mismatch``); other types are left to ``type``.
* ``unique``: no two cells hold the same value (``duplicate_value``).  Values are the same
  as keys are (:func:`~xlsx_agent.read.regions.key_text`: text exactly, a whole number as
  its digits, so ``42`` and ``"42"`` are the same), others by type and value.

A schema that is not one of these is refused (:class:`SchemaError`) before anything is
read, saying where (``tables[0].columns[2].min``), why, and what is valid instead -- never
echoing a value it was given (it may be a cell's).

**Problems** (:class:`DataProblem`) are given by code (:data:`CODES`) and address, in
order: sheets, then each table as the schema lists them -- the table's own problems
(``missing_table``, ``missing_column``, ``no_validation_list``), then its cells by row and
column.  Cells next to each other in a column with the same problem become one problem
over a range (``Stock!C5:C9``, ``cells`` 5).  Each has the table and column, what was
expected (the type, the range, the step, how many values the enum or list holds and
its case rule), and for a repeat the address of the first cell holding it (``first``);
:attr:`DataProblem.message` says it in a line for a person fixing the file.  **No problem
holds a cell's content** -- not the value found, not a key, not a header the schema did
not name -- only addresses, names, codes, types and counts (ROADMAP.md, "Design rules", 7).

**Bounded.**  One streamed pass per table over the columns checked (:meth:`Workbook.read
<xlsx_agent.workbook.Workbook.read>`), keeping the open runs of problems, a value set per
``unique`` column and the key (long text as a 16-byte digest), and at most
:data:`MAX_PROBLEMS` problems; past that, cells are counted (:attr:`ValidationReport.counts`
is always every cell) and the rest left out (:attr:`ValidationReport.cut`).
"""

from __future__ import annotations

import hashlib
import math
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING, Any, Callable, Iterator, Mapping, Sequence

from ooxml_edit.charts.workbook import Area, Cell, column_letters

from .address import (AddressError, NameNotFound, RangeAddress, TableAddress, nearest,
                      parse_address, quote_sheet)
from .profile import ALLOWED_CAP, _allowed, _covered, _list_rules
from .query import MAX_IN, _Bad, _key, _operand
from .read.regions import Region, RegionError, key_text
from .read.types import TYPES, Value

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Workbook

#: Every problem code, in the order a report lists a table's problems.
CODES = ("missing_sheet", "missing_table", "missing_column", "no_validation_list",
         "no_cached_value", "required_blank", "wrong_type", "out_of_range", "off_step",
         "not_in_enum", "not_in_list", "pattern_mismatch", "duplicate_key", "duplicate_value")

#: The types a column may be declared as: the reader's, blank aside.
COLUMN_TYPES = tuple(t for t in TYPES if t != "blank")

#: The types ``min`` and ``max`` are for, and ``step``.
ORDERED_TYPES = ("number", "percent", "date", "time", "datetime")
STEP_TYPES = ("number", "percent")

#: How ``enum`` and ``validation_list`` compare text.
CASES = ("exact", "ignore")

#: What a schema, a table and a column take.
SCHEMA_KEYS = ("sheets", "tables")
TABLE_KEYS = ("table", "range", "name", "header_rows", "headers", "key", "columns")
COLUMN_KEYS = ("name", "type", "required", "unique", "min", "max", "step", "enum", "case",
               "validation_list", "pattern")

#: The most of each a schema takes.
MAX_SHEETS = 255
MAX_TABLES = 50
MAX_COLUMNS = 500
MAX_KEY = 8
MAX_ENUM = MAX_IN
MAX_PATTERN = 1000

#: The most problems a report keeps; past it, cells are only counted.
MAX_PROBLEMS = 10_000

#: How many nearest names a missing sheet, table or column offers.
NEAREST = 5

_KEY_TEXT = 64


class SchemaError(ValueError):
    """A schema the check refuses, before anything is read: ``code`` is ``bad_schema``;
    ``path`` where in it (``tables[0].columns[2].min``); ``valid_options`` what is valid
    there instead.  The message never holds a value the schema gave."""

    def __init__(self, path: str, message: str, *, valid_options: Sequence[Any] = ()) -> None:
        super().__init__(f"schema {path}: {message}" if path else f"schema: {message}")
        self.code = "bad_schema"
        self.path = path
        self.message = str(self)
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class DataProblem:
    """One problem: ``code`` (:data:`CODES`), ``address`` (a cell, a range of cells next to
    each other in one column, a table's address, or a sheet as ``Stock!``), ``cells`` how
    many cells it covers, ``table`` and ``column`` (as the workbook spells them), and
    ``details``: what was expected (``expected``, ``found``, ``issue``, ``min``, ``max``,
    ``step``, ``base``, ``enum_size``, ``list_size``, ``case``, ``hint``, ``first``,
    ``key``, ``reason``, ``nearest``).  Never a cell's content."""

    code: str
    address: str
    cells: int = 1
    table: str | None = None
    column: str | None = None
    details: tuple[tuple[str, Any], ...] = ()

    @property
    def detail(self) -> dict[str, Any]:
        return {k: list(v) if isinstance(v, tuple) else v for k, v in self.details}

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"code": self.code, "at": self.address}
        if self.cells != 1:
            out["cells"] = self.cells
        if self.table is not None:
            out["table"] = self.table
        if self.column is not None:
            out["column"] = self.column
        out.update(self.detail)
        out["message"] = self.message
        return out

    @property
    def message(self) -> str:
        """The problem in a line, for a person fixing the file: addresses, names, codes and
        counts only."""
        d = self.detail
        where = self.address
        if self.column is not None and self.code not in ("missing_column", "no_validation_list"):
            where += f" ({self.column})"
        if self.cells > 1:
            where += f", {self.cells} cells"
        return f"{where}: {_SAY[self.code](d, self)}"


def _enum_text(d: dict, what: str) -> str:
    case = "case counts" if d.get("case", "exact") == "exact" else "case ignored"
    text = f"not one of the {what} ({case})"
    hint = d.get("hint")
    if hint == "case_differs":
        text += "; one matches but for case"
    elif hint == "space_differs":
        text += "; one matches but for space at the ends"
    return text


def _range_text(d: dict) -> str:
    if "min" in d and "max" in d:
        return f"outside {d['min']} to {d['max']}"
    if "min" in d:
        return f"less than {d['min']}"
    return f"more than {d['max']}"


def _nearest(d: dict) -> str:
    return f"; nearest: {', '.join(d['nearest'])}" if d.get("nearest") else ""


_SAY: dict[str, Callable[[dict, DataProblem], str]] = {
    "missing_sheet": lambda d, p: "no such sheet" + _nearest(d),
    "missing_table": lambda d, p: f"not found ({d.get('reason', 'not_found')})" + _nearest(d),
    "missing_column": lambda d, p: f"no column {p.column}" + _nearest(d),
    "no_validation_list": lambda d, p: f"column {p.column} has no readable data-validation "
                                       f"list ({d.get('reason', 'none')})",
    "no_cached_value": lambda d, p: "a formula without a cached value: its value is unknown",
    "required_blank": lambda d, p: "blank; a value is required",
    "wrong_type": lambda d, p: f"not {d['expected']} (found {d['found']}"
                               + (f", {d['issue']}" if d.get("issue") else "") + ")",
    "out_of_range": lambda d, p: _range_text(d),
    "off_step": lambda d, p: f"not {d['base']} plus a whole number of steps of {d['step']}",
    "not_in_enum": lambda d, p: _enum_text(d, f"{d['enum_size']} allowed values"),
    "not_in_list": lambda d, p: _enum_text(d, f"{d['list_size']} values of the "
                                              "data-validation list"),
    "pattern_mismatch": lambda d, p: "does not match the pattern",
    "duplicate_key": lambda d, p: f"the key repeats that of the row at {d['first']}",
    "duplicate_value": lambda d, p: f"repeats the value at {d['first']}",
}


@dataclass(frozen=True)
class CheckedTable:
    """A table or region the schema named and the check found: ``table`` its address,
    ``range`` where it lies, ``rows`` its data rows, ``columns`` how many it checked."""

    table: str
    range: str
    rows: int
    columns: int

    def as_dict(self) -> dict[str, Any]:
        return {"table": self.table, "range": self.range, "rows": self.rows,
                "columns": self.columns}


@dataclass(frozen=True)
class ValidationReport:
    """What :func:`validate` found: ``problems`` in order (iterating the report gives
    them); ``counts`` every problem cell by code, always complete (a table's or sheet's
    problem counts one); ``tables`` the tables checked; ``cut`` how many problem cells
    were past :data:`MAX_PROBLEMS` and are counted only."""

    problems: tuple[DataProblem, ...]
    counts: dict[str, int]
    tables: tuple[CheckedTable, ...]
    cut: int = 0

    @property
    def ok(self) -> bool:
        """No problem at all."""
        return not self.counts

    def __iter__(self) -> Iterator[DataProblem]:
        return iter(self.problems)

    def __len__(self) -> int:
        return len(self.problems)


# ------------------------------------------------------------------------------------------
# The schema, checked
# ------------------------------------------------------------------------------------------


@dataclass
class _Column:
    name: str
    path: str
    type: str | None = None
    required: bool = False
    unique: bool = False
    low: Any = None
    high: Any = None
    shown: dict[str, Any] = field(default_factory=dict)    # min and max as given
    step: Decimal | None = None
    base: Decimal = Decimal(0)
    enum: tuple | None = None
    case: str = "exact"
    validation_list: bool = False
    pattern: re.Pattern | None = None


@dataclass
class _Table:
    path: str
    table: TableAddress | None
    range: RangeAddress | None
    name: str | None
    header_rows: int
    headers: list[str] | None
    key: tuple[str, ...]
    columns: list[_Column]

    @property
    def label(self) -> str:
        return self.table.text if self.table is not None else f"t:{self.name}"


def _fail(path: str, message: str, options: Sequence[Any] = ()) -> SchemaError:
    return SchemaError(path, message, valid_options=options)


def _mapping(raw: Any, path: str, keys: Sequence[str]) -> Mapping:
    if not isinstance(raw, Mapping):
        raise _fail(path, "is an object", keys)
    unknown = [k for k in raw if k not in keys]
    if unknown:
        raise _fail(path, f"has keys it does not take ({len(unknown)}); it takes "
                    f"{', '.join(keys)}", keys)
    return {k: v for k, v in raw.items() if v is not None}


def _names(raw: Any, path: str, most: int, what: str) -> list[str]:
    if isinstance(raw, str):
        raw = [raw]
    if isinstance(raw, (bytes, Mapping)) or not isinstance(raw, Sequence) or not raw \
            or len(raw) > most:
        raise _fail(path, f"is a list of 1 to {most} {what}")
    for index, item in enumerate(raw):
        if not isinstance(item, str) or not item:
            raise _fail(f"{path}[{index}]", f"is a non-empty string, one of the {what}")
    return list(raw)


def _flag(raw: Mapping, name: str, path: str) -> bool:
    value = raw.get(name, False)
    if not isinstance(value, bool):
        raise _fail(f"{path}.{name}", "is true or false", [True, False])
    return value


def parse_schema(schema: Any, date1904: bool = False) -> tuple[list[str], list[_Table]]:
    """The schema checked, as :func:`validate` uses it; :class:`SchemaError` for one that
    is not as the module's notes say.  ``date1904``: the workbook's date system, which an
    ISO date's ``min`` or ``max`` is read in."""
    raw = _mapping(schema, "", SCHEMA_KEYS)
    sheets = _names(raw["sheets"], "sheets", MAX_SHEETS, "sheet names") \
        if "sheets" in raw else []
    tables: list[_Table] = []
    if "tables" in raw:
        items = raw["tables"]
        if isinstance(items, (str, bytes, Mapping)) or not isinstance(items, Sequence) \
                or len(items) > MAX_TABLES:
            raise _fail("tables", f"is a list of at most {MAX_TABLES} tables")
        tables = [_table(item, f"tables[{i}]", date1904) for i, item in enumerate(items)]
    if not sheets and not tables:
        raise _fail("", "names some sheets or tables to check", list(SCHEMA_KEYS))
    return sheets, tables


def _table(raw: Any, path: str, date1904: bool) -> _Table:
    item = _mapping(raw, path, TABLE_KEYS)
    table = where = name = headers = None
    header_rows = 1
    if ("table" in item) == ("range" in item):
        raise _fail(path, "names a table (t:Products) or declares a region (range, name, "
                    "header_rows, headers), one of them", ["table", "range"])
    if "table" in item:
        for spec in ("name", "header_rows", "headers"):
            if spec in item:
                raise _fail(f"{path}.{spec}", "is for a declared region (with range), not "
                            "with table", ["range"])
        table = _address(item["table"], f"{path}.table")
    else:
        where = _range(item["range"], f"{path}.range")
        name = item.get("name")
        if not isinstance(name, str) or not name:
            raise _fail(f"{path}.name", "names the declared region (Returns): a region "
                        "spec needs one", ["Returns"])
        header_rows = item.get("header_rows", 1)
        if isinstance(header_rows, bool) or not isinstance(header_rows, int) \
                or header_rows < 0:
            raise _fail(f"{path}.header_rows", "is a whole number from 0")
        if "headers" in item:
            headers = _names(item["headers"], f"{path}.headers", 16384, "header names")
    key = tuple(_names(item["key"], f"{path}.key", MAX_KEY, "key columns")) \
        if "key" in item else ()
    if len({k.casefold() for k in key}) != len(key):
        raise _fail(f"{path}.key", "names each key column once")
    columns: list[_Column] = []
    if "columns" in item:
        items = item["columns"]
        if isinstance(items, (str, bytes, Mapping)) or not isinstance(items, Sequence) \
                or len(items) > MAX_COLUMNS:
            raise _fail(f"{path}.columns", f"is a list of at most {MAX_COLUMNS} columns")
        columns = [_column(c, f"{path}.columns[{i}]", date1904) for i, c in enumerate(items)]
        folded = [c.name.casefold() for c in columns]
        if len(set(folded)) != len(folded):
            raise _fail(f"{path}.columns", "names each column once, case aside")
    if not key and not columns:
        raise _fail(path, "has a key or columns to check", ["key", "columns"])
    return _Table(path, table, where, name, header_rows, headers, key, columns)


def _address(raw: Any, path: str) -> TableAddress:
    try:
        address = parse_address(raw) if isinstance(raw, str) else None
    except AddressError as error:
        raise _fail(path, f"is not an address ({error.message})") from None
    if not isinstance(address, TableAddress) or not address.is_whole:
        raise _fail(path, "is a table or region address, as t:Products or t:Stock!A1",
                    ["t:Products"])
    return address


def _range(raw: Any, path: str) -> RangeAddress:
    try:
        address = parse_address(raw) if isinstance(raw, str) else None
    except AddressError as error:
        raise _fail(path, f"is not an address ({error.message})") from None
    if not isinstance(address, RangeAddress):
        raise _fail(path, "is a range, as Returns!A3:F200", ["Returns!A3:F200"])
    return address


def _column(raw: Any, path: str, date1904: bool) -> _Column:
    item = _mapping(raw, path, COLUMN_KEYS)
    name = item.get("name")
    if not isinstance(name, str) or not name:
        raise _fail(f"{path}.name", "names the column, as its header does")
    column = _Column(name, path)
    t = item.get("type")
    if t is not None and t not in COLUMN_TYPES:
        raise _fail(f"{path}.type", "is one of the reader's types", COLUMN_TYPES)
    column.type = t
    column.required = _flag(item, "required", path)
    column.unique = _flag(item, "unique", path)
    column.validation_list = _flag(item, "validation_list", path)
    for bound in ("min", "max"):
        if bound not in item:
            continue
        if t not in ORDERED_TYPES:
            raise _fail(f"{path}.{bound}", "needs the column's type, one that is ordered",
                        ORDERED_TYPES)
        value = _value(item[bound], t, date1904, f"{path}.{bound}")
        setattr(column, "low" if bound == "min" else "high", value)
        column.shown[bound] = item[bound]
    if column.low is not None and column.high is not None and column.low > column.high:
        raise _fail(f"{path}.max", "is less than min")
    if "step" in item:
        if t not in STEP_TYPES:
            raise _fail(f"{path}.step", "needs the column's type, a number or percent",
                        STEP_TYPES)
        column.step = _decimal(item["step"], f"{path}.step")
        if column.step <= 0:
            raise _fail(f"{path}.step", "is more than 0")
        if column.low is not None:
            column.base = _decimal(column.low, f"{path}.min")
    if "enum" in item:
        if column.validation_list:
            raise _fail(f"{path}.enum", "is not with validation_list: one or the other",
                        ["enum", "validation_list"])
        values = item["enum"]
        if isinstance(values, (str, bytes, Mapping)) or not isinstance(values, Sequence) \
                or not values or len(values) > MAX_ENUM:
            raise _fail(f"{path}.enum", f"is a list of 1 to {MAX_ENUM} values")
        for index, value in enumerate(values):
            if not isinstance(value, (str, int, float)):
                raise _fail(f"{path}.enum[{index}]", "is a value, written as a filter's is")
            if t is not None:
                _value(value, t, date1904, f"{path}.enum[{index}]")
        column.enum = tuple(values)
    if "case" in item:
        if item["case"] not in CASES:
            raise _fail(f"{path}.case", "is how text is compared", CASES)
        if column.enum is None and not column.validation_list:
            raise _fail(f"{path}.case", "is for enum or validation_list")
        column.case = item["case"]
    if "pattern" in item:
        pattern = item["pattern"]
        if t not in (None, "text"):
            raise _fail(f"{path}.pattern", "is for a text column", ["text"])
        if not isinstance(pattern, str) or not pattern or len(pattern) > MAX_PATTERN:
            raise _fail(f"{path}.pattern", f"is a regular expression of 1 to {MAX_PATTERN} "
                        "characters")
        try:
            column.pattern = re.compile(pattern)
        except re.error as error:
            # where only: some of re.error's messages quote a piece of the pattern
            raise _fail(f"{path}.pattern", "is not a Python regular expression (at "
                        f"character {(error.pos or 0) + 1})") from None
    if not (t or column.required or column.unique or column.shown or column.step
            or column.enum or column.validation_list or column.pattern):
        raise _fail(path, "has nothing to check; give type, required, unique, min, max, "
                    "step, enum, validation_list or pattern", COLUMN_KEYS[1:])
    return column


def _value(raw: Any, t: str, date1904: bool, path: str) -> Any:
    try:
        return _operand(raw, t, date1904)
    except _Bad:
        raise _fail(path, f"is not a {t} value, written as a filter's value is") from None


def _decimal(raw: Any, path: str) -> Decimal:
    if isinstance(raw, bool) or not isinstance(raw, (int, float, str)):
        raise _fail(path, "is a number")
    try:
        number = Decimal(repr(raw) if isinstance(raw, float) else str(raw).strip())
    except InvalidOperation:
        raise _fail(path, "is a number") from None
    if not number.is_finite():
        raise _fail(path, "is a finite number")
    return number


# ------------------------------------------------------------------------------------------
# The check
# ------------------------------------------------------------------------------------------


def validate(workbook: "Workbook", schema: Any) -> ValidationReport:
    """``workbook`` checked against ``schema`` (see the module's notes):
    :class:`SchemaError` for a schema that is not one, else the report."""
    sheets, tables = parse_schema(schema, workbook.date1904)
    out = _Problems()
    for name in sheets:
        try:
            workbook.sheet(name)
        except NameNotFound as error:
            out.add(-1, 0, (0, 0), "missing_sheet", f"{quote_sheet(name)}!", None, None,
                    {"nearest": error.valid_options[:NEAREST]})
    checked: list[CheckedTable] = []
    for index, spec in enumerate(tables):
        region = _region(workbook, spec, index, out)
        if region is not None:
            checked.append(_check(workbook, region, spec, index, out))
    problems, counts, cut = out.done()
    return ValidationReport(problems, counts, tuple(checked), cut)


def _region(workbook: "Workbook", spec: _Table, index: int, out: "_Problems") -> Region | None:
    label = spec.label
    try:
        if spec.table is not None:
            return workbook.region(spec.table)
        return workbook.regions.declare(spec.name, spec.range, header_rows=spec.header_rows,
                                        headers=spec.headers, keep=False)
    except NameNotFound as error:
        reason = "sheet_not_found" if error.what in ("sheet", "sheet id") else "not_found"
        options = error.valid_options[:NEAREST] if reason == "not_found" else []
        out.add(index, 0, (0, 0), "missing_table", label, label, None,
                {"reason": reason, "nearest": options})
    except RegionError as error:
        if error.code in ("bad_name", "name_taken", "bad_headers", "headers_required"):
            raise _fail(spec.path, error.message, error.valid_options) from None
        reason = "no_cells" if error.code == "no_header" and spec.table is None else error.code
        out.add(index, 0, (0, 0), "missing_table", label, label, None, {"reason": reason})
    return None


class _Check:
    """One column's checks, cell by cell."""

    def __init__(self, column: _Column, name: str, offset: int, required: bool,
                 allowed: tuple | None, date1904: bool) -> None:
        self.spec = column
        self.name = name
        self.offset = offset
        self.required = required
        self.date1904 = date1904
        self.allowed = allowed                    # validation_list's values
        self.seen: dict | None = {} if column.unique else None
        self._sets: dict[tuple[str, str], tuple] = {}

    def run(self, value: Value, emit: Callable[[str, Value, dict], None]) -> None:
        spec = self.spec
        if value.issue == "no_cached_value":
            emit("no_cached_value", value, {})
            return
        t = value.t
        if t == "blank":
            if self.required:
                emit("required_blank", value, {})
            return
        if spec.type is not None and t != spec.type:
            detail = {"expected": spec.type, "found": t}
            if value.issue is not None:
                detail["issue"] = value.issue
            emit("wrong_type", value, detail)
            return
        if spec.low is not None or spec.high is not None:
            key = _key(value, t, self.date1904)
            if key is not None and ((spec.low is not None and key < spec.low)
                                    or (spec.high is not None and key > spec.high)):
                emit("out_of_range", value, dict(spec.shown))
        if spec.step is not None and isinstance(value.v, (int, float)) \
                and not isinstance(value.v, bool) and math.isfinite(value.v):
            if (Decimal(repr(value.v)) - spec.base) % spec.step != 0:
                emit("off_step", value, {"step": _number(spec.step),
                                         "base": _number(spec.base)})
        if spec.enum is not None:
            hint = self._member(value, spec.enum, "enum")
            if hint is not True:
                emit("not_in_enum", value, {"enum_size": len(spec.enum), "case": spec.case,
                                            **({"hint": hint} if hint else {})})
        if self.allowed is not None:
            hint = self._member(value, self.allowed, "list")
            if hint is not True:
                emit("not_in_list", value, {"list_size": len(self.allowed), "case": spec.case,
                                            **({"hint": hint} if hint else {})})
        if spec.pattern is not None and t == "text" and not spec.pattern.fullmatch(value.v):
            emit("pattern_mismatch", value, {})
        if self.seen is not None:
            identity = _identity(value)
            first = self.seen.get(identity)
            if first is None:
                self.seen[identity] = value.a
            else:
                emit("duplicate_value", value, {"first": first})

    def _member(self, value: Value, allowed: tuple, which: str) -> "bool | str | None":
        """``True`` when the cell is one of ``allowed``; else the hint (``case_differs``,
        ``space_differs``) or ``None``."""
        t = value.t
        sets = self._sets.get((which, t))
        if sets is None:
            parsed = set()
            for raw in allowed:
                try:
                    parsed.add(_operand(raw, t, self.date1904))
                except _Bad:
                    continue
            if t == "text":
                folded = {p.casefold() for p in parsed}
                stripped = {p.strip().casefold() for p in parsed}
                if self.spec.case == "ignore":
                    parsed = folded
                sets = (frozenset(parsed), frozenset(folded), frozenset(stripped))
            else:
                sets = (frozenset(parsed), frozenset(), frozenset())
            self._sets[(which, t)] = sets
        own, folded, stripped = sets
        key = _key(value, t, self.date1904)
        if t == "text":
            text = value.v
            if (text.casefold() if self.spec.case == "ignore" else text) in own:
                return True
            if self.spec.case == "exact" and text.casefold() in folded:
                return "case_differs"
            if text.strip().casefold() in stripped:
                return "space_differs"
            return None
        return True if key is not None and key in own else None


def _identity(value: Value) -> Any:
    """A cell as :func:`~xlsx_agent.read.regions.key_text` keys it, else by type and value;
    long text as a digest."""
    text = key_text(value)
    v: Any = text if text is not None else value.v
    if isinstance(v, str) and len(v) > _KEY_TEXT:
        v = hashlib.blake2b(v.encode("utf-8", "surrogatepass"), digest_size=16).digest()
    return v if text is not None else (value.t, v)


def _number(value: Decimal) -> "int | float":
    return int(value) if value == value.to_integral_value() else float(value)


def _check(workbook: "Workbook", region: Region, spec: _Table, index: int,
           out: "_Problems") -> CheckedTable:
    label = spec.label
    folded = {c.casefold(): i for i, c in enumerate(region.columns)}
    date1904 = workbook.date1904
    checks: list[_Check] = []
    key_offsets: list[int] = []

    def missing(name: str) -> None:
        out.add(index, 0, (0, 0), "missing_column", label, label, name,
                {"nearest": nearest(name, region.columns)[:NEAREST]})

    for name in spec.key:
        if name.casefold() not in folded:
            missing(name)
        else:
            key_offsets.append(folded[name.casefold()])
    key_names = {region.columns[o].casefold() for o in key_offsets}
    rules = None
    for column in spec.columns:
        offset = folded.get(column.name.casefold())
        if offset is None:
            if column.name.casefold() not in {k.casefold() for k in spec.key}:
                missing(column.name)
            continue
        name = region.columns[offset]
        allowed = None
        if column.validation_list:
            if rules is None:
                rules = _list_rules(workbook, region)
            allowed, reason = _allowed_values(workbook, region, rules, name)
            if allowed is None:
                out.add(index, 0, (0, offset), "no_validation_list", label, label, name,
                        {"reason": reason})
        checks.append(_Check(column, name, offset, column.required
                             or name.casefold() in key_names, allowed, date1904))
    for offset in key_offsets:   # a key column not otherwise checked is still never blank
        if not any(c.offset == offset for c in checks):
            name = region.columns[offset]
            checks.append(_Check(_Column(name, spec.path, required=True), name, offset, True,
                                 None, date1904))

    data = region.data
    rows = region.rows
    if data is not None and checks:
        first = region.area.first.column
        offsets = [c.offset for c in checks]
        left, right = first + min(offsets), first + max(offsets)
        area = Area(Cell(data.first.row, left), Cell(data.last.row, right))
        positions = [first + c.offset - left for c in checks]
        key_positions = [first + o - left for o in key_offsets]
        keys: dict = {}
        reading = {first + o for o in offsets}
        for row in workbook._read(RangeAddress(region.sheet_ref, area), columns=reading):
            for check, position in zip(checks, positions):
                value = row[position]
                check.run(value, lambda code, v, detail, check=check: out.add(
                    index, 1, (v.row, check.offset), code, None, label, check.name, detail,
                    cell=v))
            if key_positions:
                parts = [row[p] for p in key_positions]
                if any(p.t == "blank" for p in parts):
                    continue        # reported as required_blank (or no_cached_value)
                identity = tuple(_identity(p) for p in parts)
                lead = parts[0]
                found = keys.get(identity)
                if found is None:
                    keys[identity] = lead.a
                else:
                    detail: dict[str, Any] = {"first": found}
                    if len(key_positions) > 1:
                        detail["key"] = [region.columns[o] for o in key_offsets]
                    out.add(index, 1, (lead.row, key_offsets[0]), "duplicate_key", None,
                            label, region.columns[key_offsets[0]], detail, cell=lead)
    return CheckedTable(label, region.range.text, rows, len(checks))


def _allowed_values(workbook: "Workbook", region: Region, rules: list,
                    name: str) -> tuple[tuple | None, str | None]:
    """The values of the data-validation lists over a column (each list that covers any of
    its cells), or ``None`` and why there are none."""
    area = region.column_area(name)
    if area is None:
        return None, "no_rows"
    lists = [_allowed(workbook, region, rule, area) for rule in rules
             if _covered(rule.areas, area)]
    if not lists:
        return None, "no_list"
    values: list = []
    for found in lists:
        if found.values is None:
            return None, found.reason or "formula"
        if found.cut:
            return None, f"over_{ALLOWED_CAP}_values"
        values.extend(found.values)
    return tuple(dict.fromkeys(values)), None


class _Problems:
    """The problems as they are met: runs of cells next to each other in a column with the
    same problem kept open until they end; at most :data:`MAX_PROBLEMS` kept."""

    def __init__(self) -> None:
        self.kept: list[tuple[tuple, DataProblem]] = []
        self.open: dict[tuple, list] = {}       # run key -> [order, first, last row, cells]
        self.counts: dict[str, int] = {}
        self.cut = 0
        self.runs = 0

    def add(self, table: int, phase: int, place: tuple[int, int], code: str,
            address: str | None, label: str | None, column: str | None, detail: dict,
            *, cell: Value | None = None) -> None:
        self.counts[code] = self.counts.get(code, 0) + 1
        details = tuple((k, tuple(v) if isinstance(v, list) else v)
                        for k, v in detail.items() if v not in (None, [], ()))
        if cell is None:
            self._keep((table, phase, place, CODES.index(code)),
                       DataProblem(code, address, 1, label, column, details))
            return
        run_key = (table, cell.sheet, cell.column, code, details)
        run = self.open.get(run_key)
        if run is not None and run[2] == cell.row - 1:
            run[2] = cell.row
            run[3] += 1
            return
        if run is not None:
            self._close(run_key, run)
        if self.runs >= MAX_PROBLEMS:
            self.cut += 1
            return
        self.runs += 1
        self.open[run_key] = [(table, phase, place, CODES.index(code)), cell, cell.row, 1,
                              label, column]

    def _close(self, run_key: tuple, run: list) -> None:
        order, cell, last, cells, label, column = run
        address = cell.a if last == cell.row else \
            f"{cell.a}:{column_letters(cell.column)}{last}"
        self.kept.append((order, DataProblem(run_key[3], address, cells, label, column,
                                             run_key[4])))
        del self.open[run_key]

    def _keep(self, order: tuple, problem: DataProblem) -> None:
        self.kept.append((order, problem))

    def done(self) -> tuple[tuple[DataProblem, ...], dict[str, int], int]:
        for run_key, run in list(self.open.items()):
            self._close(run_key, run)
        self.kept.sort(key=lambda item: item[0])
        counts = {code: self.counts[code] for code in CODES if code in self.counts}
        return tuple(p for _, p in self.kept), counts, self.cut


__all__ = ["CASES", "CODES", "COLUMN_TYPES", "MAX_PROBLEMS", "CheckedTable", "DataProblem",
           "SchemaError", "ValidationReport", "parse_schema", "validate"]
