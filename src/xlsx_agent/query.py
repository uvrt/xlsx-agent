"""Queries: the rows of a table, name or region that meet some filters, sorted, with the
columns asked for, each cell with its address (ROADMAP.md, "Tools": ``xlsx_query``)::

    products = wb.table("t:Products")
    found = products.query(["SKU", "Price"], where={"Category": {"eq": "Hardware"},
                                                    "Price": {"gte": 5, "lte": 50}},
                           sort=[("Price", "desc"), "SKU"], limit=20)
    found.matches        # how many rows match, all of them
    found.rows           # [(Value(a='Stock!A7', ...), Value(a='Stock!C7', ...)), ...]

**Filters.**  ``where`` maps a column name to its tests, ``{"Price": {"gte": 5}}``, or is
a list of ``{"column": ..., "op": ..., "value": ...}`` (``values`` for ``in``), as the tool
gives it; a row matches when it passes every test (they combine with *and*).  The tests:
``eq`` (equal to the value), ``in`` (equal to one of the values), ``gte`` and ``lte`` (at
least, at most), ``contains`` (text holding the value), ``words_gt`` (text of more words
than the value, a whole number: words are the runs of characters between white space, as
:meth:`str.split` splits them; a test for text too long to place), ``blank`` and
``not_blank`` (the reader's ``blank``: no value, a formula without a cached value among
them; ``True`` in the mapping form, nothing in the list form).

**Types.**  Each filtered column's type is its profile's (:mod:`xlsx_agent.profile`: the
type most of its values have, :func:`~xlsx_agent.profile.majority_type`), and a test is
checked against it before any row is given: ``gte``/``lte`` take a number, percent, date,
time or datetime column; ``contains`` and ``words_gt`` a text column; ``eq`` and ``in``
any column with values; ``blank`` and ``not_blank`` any column.  Any other pairing, and a
value that is not of the column's type, is refused (:class:`QueryError`, ``wrong_type``)
naming the column's type, with the tests it takes as ``valid_options``.  A cell of another type than its
column's never matches a test that compares values: nothing is coerced (ROADMAP.md, "Design
rules", 4); such cells are counted per column, with their first addresses
(:attr:`QueryResult.not_compared`).  A value may be given as Python's (``5``, ``True``) or
as text in its JSON form (``"5"``, ``"2.5e3"``, ``"true"``), as a tool call gives it.

* **number** and **percent**: compared as stored; a percent is the stored fraction (25 %
  is ``0.25``).
* **date**, **time** and **datetime**: compared as the reader shows them, whole
  milliseconds, in the workbook's date system: a value is ISO 8601 text as the reader
  gives ``v`` (``2026-01-05``, ``2026-01-05T10:30:00``, ``10:30``, an elapsed time
  ``PT36H0M0S``) or a serial as it gives ``serial``.  A date column compares days (a date
  cell's time of day is not shown, and not compared), a time column the time of day (an
  elapsed time its length), a datetime column the moment; an ISO date on a datetime column
  is its midnight.  A cell written as ISO text (``t="d"``) is read in the same system.
* **text**: ``eq`` and ``in`` match exactly: case counts, nothing trimmed or normalised,
  as keys are matched (:mod:`xlsx_agent.read.regions`).  ``contains`` ignores case
  (:meth:`str.casefold` on both sides), as a search does.  Neither is Excel's own
  comparison, which ignores case in both and whose collation is not measured here; text
  has no ``gte``/``lte`` for that reason.
* **bool**: ``True``/``False``; **error**: the code as the reader gives it (``#N/A``),
  exactly.
* **list** (a declared column whose text is split, :mod:`xlsx_agent.clean`): ``eq``,
  ``in`` and ``contains`` test each item as text is tested, and a cell passes when any
  item does (``{"owners": {"eq": "Kim"}}``: the rows Kim is among); a list sorts by its
  items, as text does, item by item.

**Sorting.**  ``sort`` is a list of keys, the first deciding first: a column name
(ascending) or ``(name, "asc" | "desc")``.  Within a key, the cells of the column's type
come first, in order (text by :meth:`str.casefold`, then as written: code-point order, not
Excel's collation); then cells of other types, then blanks, each as they were (Excel too
puts blanks last either way).  Rows that tie keep the table's order.

**Bounded.**  :func:`matches` is one streamed pass over the filtered and sorted columns
that keeps, per matching row, its row number and sort keys, and the cells asked for while
they are at most :data:`KEEP_CELLS`; :func:`read_rows` gives a page of rows from those, or
reads them again once they are known (:func:`query` does both, ``offset`` and ``limit``
giving the page).  ``xlsx_query`` pages through them as ``xlsx_read`` pages through cells.

**A dataset** (``d:by_region``, :func:`dataset_matches`; round 10, O5) is queried as a
table is: the same filters, types and sort, over its values as it reports them, its rows
named by position.

**Content** (ROADMAP.md, "Design rules", 7).  A query's rows are a result.  An error says
which column, which test and which type, never a filter's value (a value may be a cell's,
as a key may).
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Iterable, Iterator, Mapping, Sequence, cast

from ooxml_edit.sheets import Area, Cell

from .address import RangeAddress
from .profile import FIRST_ADDRESSES, Count, majority_type
from .read.regions import Region
from .read.stream import CHECKPOINT_ROWS
from .read.types import Value, day_serial

if TYPE_CHECKING:  # pragma: no cover
    from .dataset import DatasetTable
    from .workbook import Workbook

#: Every test a filter can make.
OPS = ("eq", "in", "gte", "lte", "contains", "words_gt", "blank", "not_blank")

_ORDERED = ("eq", "in", "gte", "lte", "blank", "not_blank")
_ANY = ("eq", "in", "blank", "not_blank")

#: The tests each column type takes.
OPS_BY_TYPE: dict[str, tuple[str, ...]] = {
    "number": _ORDERED, "percent": _ORDERED, "date": _ORDERED, "time": _ORDERED,
    "datetime": _ORDERED, "text": ("eq", "in", "contains", "words_gt", "blank", "not_blank"),
    "bool": _ANY, "error": _ANY, "blank": ("blank", "not_blank"),
    "list": ("eq", "in", "contains", "blank", "not_blank")}

#: The most values an ``in`` test takes.
MAX_IN = 1000

#: The most cells :func:`matches` keeps of the columns asked for, so that the pages of a
#: query need no second pass; past it they are read again, a page at a time.
KEEP_CELLS = 50_000

#: What a value of each type is written as, for an error message.
_FORMS = {"number": "a number", "percent": "a number (the stored fraction: 25 % is 0.25)",
          "date": "an ISO 8601 date (2026-01-05) or a serial",
          "time": "an ISO 8601 time (10:30:00), an elapsed time (PT36H0M0S) or a serial",
          "datetime": "an ISO 8601 date and time (2026-01-05T10:30:00) or a serial",
          "text": "text", "bool": "true or false", "error": "an error code (#N/A)"}

_MOMENTS = ("date", "time", "datetime")
_MS_PER_DAY = 86_400_000
_NUMBER = re.compile(r"-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
_ZONE = r"(?:Z|[+-]\d{2}:\d{2})?"
_CLOCK = r"(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?"
_ISO_DAY = re.compile(rf"(\d{{4}})-(\d{{2}})-(\d{{2}})(?:[T ]{_CLOCK})?{_ZONE}")
_ISO_TIME = re.compile(rf"{_CLOCK}{_ZONE}")
_DURATION = re.compile(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)(?:\.(\d+))?S)?")


class QueryError(ValueError):
    """A query the library refuses: ``code`` says why.

    ``wrong_type`` (a test the column's type does not take, or a value not of its type:
    ``column_type`` names the type, ``valid_options`` the tests it takes), ``bad_filter``
    (an unknown test, a missing value, a value where none is taken) and ``bad_sort``.
    ``field`` is the argument at fault (``where``, ``sort``, ``columns``).  The message
    names columns, tests and types, never a value."""

    def __init__(self, code: str, message: str, *, field: str = "where",
                 column: str | None = None, column_type: str | None = None,
                 valid_options: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.field = field
        self.column = column
        self.column_type = column_type
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class Matches:
    """What :func:`matches` found: ``rows`` the sheet rows of the matching data rows, in
    the result's order (of a dataset, :func:`dataset_matches`, their positions); ``types`` each filtered or sorted column's type;
    ``not_compared`` per filtered column, the cells a comparing test met of another type
    than the column's (never matched).  ``values``: the cells of the columns asked for, by
    sheet row, when the pass kept them (:data:`KEEP_CELLS`); ``None`` when it did not."""

    region: "Region | DatasetTable"
    rows: tuple[int, ...]
    types: dict[str, str]
    not_compared: dict[str, Count] = field(default_factory=dict)
    values: dict[int, list[Value]] | None = field(default=None, repr=False, compare=False)


@dataclass(frozen=True)
class QueryResult:
    """A query's page: ``address`` the region's (``t:Products``); ``columns`` the names
    given, in order; ``rows`` the matching rows from ``offset``, each a value per column
    with its address; ``positions`` each row's data-row position (from 1:
    ``t:Products/#5``); ``matches`` how many rows match in all; ``types`` and
    ``not_compared`` as :class:`Matches` gives them."""

    address: str
    columns: tuple[str, ...]
    rows: tuple[tuple[Value, ...], ...]
    positions: tuple[int, ...]
    matches: int
    types: dict[str, str]
    not_compared: dict[str, Count] = field(default_factory=dict)

    def as_dict(self) -> dict:
        """The page as plain data: rows of cells as :meth:`Value.as_dict
        <xlsx_agent.read.types.Value.as_dict>` writes them."""
        out: dict = {"address": self.address, "columns": list(self.columns),
                     "rows": [[v.as_dict() for v in row] for row in self.rows],
                     "positions": list(self.positions), "matches": self.matches,
                     "types": dict(self.types)}
        if self.not_compared:
            out["not_compared"] = {n: c.as_dict() for n, c in self.not_compared.items()}
        return out


def query(workbook: "Workbook", region: Region, columns: Sequence[str] | None = None, *,
          where: "Mapping[str, Mapping[str, Any]] | Sequence[Mapping[str, Any]] | None" = None,
          sort: "Sequence[str | tuple[str, str]] | None" = None, limit: int | None = None,
          offset: int = 0) -> QueryResult:
    """The rows of ``region`` that pass ``where``, ordered by ``sort``, from ``offset`` and
    at most ``limit`` of them, with ``columns`` (names, matched case aside; every column
    when ``None`` or empty); see the module's notes."""
    if limit is not None and (isinstance(limit, bool) or not isinstance(limit, int) or limit < 0):
        raise ValueError("limit is a whole number of rows, 0 or more")
    if isinstance(offset, bool) or not isinstance(offset, int) or offset < 0:
        raise ValueError("offset is a whole number of rows, 0 or more")
    names = select(region, columns)
    found = matches(workbook, region, where=where, sort=sort, columns=names)
    page = found.rows[offset:] if limit is None else found.rows[offset:offset + limit]
    top = region.area.first.row + region.header_rows
    rows = read_rows(workbook, region, names, page, kept=found.values)
    return QueryResult(address=region.address.text, columns=names,
                       rows=tuple(tuple(row) for row in rows),
                       positions=tuple(row - top + 1 for row in page),
                       matches=len(found.rows), types=found.types,
                       not_compared=found.not_compared)


def select(region: "Region | DatasetTable", columns: Sequence[str] | None
           ) -> tuple[str, ...]:
    """The columns named, as the region spells them, each once in the order first named;
    every column when ``None`` or empty."""
    if isinstance(columns, str):
        columns = [columns]
    if not columns:
        return tuple(region.columns)
    return tuple(dict.fromkeys(region.columns[region.column_index(name)] for name in columns))


def matches(workbook: "Workbook", region: Region, *,
            where: "Mapping[str, Mapping[str, Any]] | Sequence[Mapping[str, Any]] | None" = None,
            sort: "Sequence[str | tuple[str, str]] | None" = None,
            columns: Sequence[str] = ()) -> Matches:
    """The matching rows of ``region`` in order, in one streamed pass over the columns
    filtered and sorted; :class:`QueryError` for a filter its column's type does not take
    (checked once the pass has the types, before any row is given).  The cells of
    ``columns`` (names as the region spells them) in the matching rows are kept while they
    are few (:attr:`Matches.values`)."""
    tests = _tests(region, where)
    keys = _sort_keys(region, sort)
    data = region.data
    if data is None:
        return Matches(region, (), {})
    if not tests and not keys:
        return Matches(region, tuple(range(data.first.row, data.last.row + 1)), {})
    date1904 = workbook.date1904
    first = region.area.first.column
    offsets = sorted({t.offset for t in tests} | {k.offset for k in keys})
    wanted_offsets = [region.column_index(name) for name in columns]
    span = offsets + wanted_offsets
    left, right = first + min(span), first + max(span)
    kept: dict[int, list[Value]] | None = {} if wanted_offsets else None
    room = KEEP_CELLS
    counted: dict[int, dict] = {o: {} for o in offsets}  # offset -> type -> [count, firsts]
    typed = [t for t in tests if t.compares]
    found: list[tuple] = []                   # (row, types per typed test, keys per sort key)
    area = Area(Cell(data.first.row, left), Cell(data.last.row, right))
    # The cells typed: the filtered and sorted columns', and the kept ones' while they are.
    reading = {first + o for o in offsets} | {first + o for o in wanted_offsets}
    for row in workbook._read(RangeAddress(region.sheet_ref, area), columns=reading):
        cells = {o: row[first + o - left] for o in offsets}
        for o, value in cells.items():
            if value.t != "blank":
                entry = counted[o].setdefault(value.t, [0, []])
                entry[0] += 1
                if len(entry[1]) < FIRST_ADDRESSES:
                    entry[1].append(value)
        if not all(t.passes(cells[t.offset], date1904) for t in tests):
            continue
        number = row[0].row
        found.append((number, tuple(cells[t.offset].t for t in typed),
                      tuple(k.key(cells[k.offset], date1904) for k in keys)))
        if kept is not None:
            room -= len(wanted_offsets)
            if room < 0:
                kept = None
                reading.difference_update(first + o for o in wanted_offsets)
                reading.update(first + o for o in offsets)
            else:
                kept[number] = [row[first + o - left] for o in wanted_offsets]

    rows, types, not_compared = _settled(region, tests, keys, offsets, counted, found, date1904)
    return Matches(region, rows, types, not_compared, kept)


def dataset_matches(table: "DatasetTable", *,
                    where: "Mapping[str, Mapping[str, Any]] | Sequence[Mapping[str, Any]] | None" = None,
                    sort: "Sequence[str | tuple[str, str]] | None" = None) -> Matches:
    """:func:`matches` of a dataset (:class:`~xlsx_agent.dataset.DatasetTable`, round 10,
    O5): its rows that pass ``where``, ordered by ``sort``, by position (from 1).  The
    values tested are the dataset's as it reports them -- what ``xlsx_read`` shows of it,
    rounded where its derive rounded -- each typed as a dataset's values are (a column's
    label first); the filters, types and sort are a table's."""
    tests = _tests(table, where)
    keys = _sort_keys(table, sort)
    if not tests and not keys:
        return Matches(table, tuple(range(1, table.rows + 1)), {})
    date1904 = table.date1904
    offsets = sorted({t.offset for t in tests} | {k.offset for k in keys})
    counted: dict[int, dict] = {o: {} for o in offsets}
    typed = [t for t in tests if t.compares]
    found: list[tuple] = []
    for position in range(1, table.rows + 1):
        cells = {o: table.value(position, o) for o in offsets}
        for o, value in cells.items():
            if value.t != "blank":
                entry = counted[o].setdefault(value.t, [0, []])
                entry[0] += 1
                if len(entry[1]) < FIRST_ADDRESSES:
                    entry[1].append(value)
        if all(t.passes(cells[t.offset], date1904) for t in tests):
            found.append((position, tuple(cells[t.offset].t for t in typed),
                          tuple(k.key(cells[k.offset], date1904) for k in keys)))
    rows, types, not_compared = _settled(table, tests, keys, offsets, counted, found, date1904)
    return Matches(table, rows, types, not_compared)


def _settled(region: Any, tests: list, keys: list, offsets: list[int], counted: dict,
             found: list[tuple], date1904: bool) -> tuple[tuple[int, ...], dict, dict]:
    """The pass's end, for a sheet's region and a dataset alike: each column's type, the
    tests checked against it, the rows of that type that passed, in order, and the cells
    of other types the comparing tests met."""
    typed = [t for t in tests if t.compares]
    types = {region.columns[o]: majority_type({t: n for t, (n, _) in counted[o].items()})
             for o in offsets}
    for test in tests:
        test.check(types[test.column], date1904)
    wanted = tuple(types[t.column] for t in typed)
    rows = [entry for entry in found if entry[1] == wanted]
    for index in reversed(range(len(keys))):
        rows = _ordered(rows, index, types[keys[index].column], keys[index].desc)
    not_compared = {}
    for test in typed:
        if test.column in not_compared:
            continue
        others = [(n, values) for t, (n, values) in counted[test.offset].items()
                  if t != types[test.column]]
        if others:
            values = sorted((v for _, vs in others for v in vs), key=lambda v: v.row)
            not_compared[test.column] = Count(sum(n for n, _ in others),
                                              tuple(v.a for v in values[:FIRST_ADDRESSES]))
    return tuple(entry[0] for entry in rows), types, not_compared


def read_rows(workbook: "Workbook", region: Region, columns: Sequence[str],
              rows: Sequence[int], *, kept: "dict[int, list[Value]] | None" = None
              ) -> Iterator[list[Value]]:
    """The values of ``columns`` (names as the region spells them) in the sheet rows
    ``rows``, in that order: from ``kept`` (:attr:`Matches.values`, of the same columns)
    when it has them, else in one streamed pass from the first of them to the last."""
    if not rows or not columns:
        return
    if kept is not None and all(number in kept for number in rows):
        yield from (kept[number] for number in rows)
        return
    first = region.area.first.column
    offsets = [region.column_index(name) for name in columns]
    left, right = first + min(offsets), first + max(offsets)
    wanted = set(rows)
    reading = {first + o for o in offsets}
    read: dict[int, list[Value]] = {}
    for top, bottom in _runs(sorted(wanted), workbook._rows_apart(region.range)):
        area = Area(Cell(top, left), Cell(bottom, right))
        for row in workbook._read(RangeAddress(region.sheet_ref, area), columns=reading):
            number = row[0].row
            if number in wanted:
                read[number] = [row[first + o - left] for o in offsets]
    for number in rows:
        yield read[number]


def _runs(rows: list[int], apart: bool) -> list[tuple[int, int]]:
    """Sorted sheet rows as the ranges of rows to read them by: one, from the first to the
    last; or, where rows far apart read alike apart (``apart``:
    :meth:`Workbook._rows_apart <xlsx_agent.workbook.Workbook._rows_apart>`), a range per
    run of rows at most a checkpoint's rows apart."""
    if not apart:
        return [(rows[0], rows[-1])]
    out = [[rows[0], rows[0]]]
    for row in rows[1:]:
        if row - out[-1][1] <= CHECKPOINT_ROWS:
            out[-1][1] = row
        else:
            out.append([row, row])
    return [(top, bottom) for top, bottom in out]


# -- filters -----------------------------------------------------------------------------------


class _Test:
    """One test of a filter: ``column`` (as the region spells it) at ``offset``, ``op``,
    and the value or values as given."""

    def __init__(self, column: str, offset: int, op: str, operand: Any) -> None:
        self.column = column
        self.offset = offset
        self.op = op
        self.operand = operand
        self.compares = op not in ("blank", "not_blank")
        self._parsed: dict[str, Any] = {}

    def _value(self, t: str, date1904: bool) -> Any:
        """The operand read as a value of type ``t`` (a set for ``in``); ``_BAD`` when it is
        not one."""
        if t not in self._parsed:
            read_as = "text" if t == "list" else t      # a list's items are text
            try:
                if self.op == "in":
                    parsed: Any = frozenset(_operand(v, read_as, date1904)
                                            for v in self.operand)
                else:
                    parsed = _words(self.operand) if self.op == "words_gt" else \
                        _operand(self.operand, read_as, date1904)
                    if self.op == "contains":
                        parsed = parsed.casefold()
            except _Bad:
                parsed = _BAD
            self._parsed[t] = parsed
        return self._parsed[t]

    def passes(self, value: Value, date1904: bool) -> bool:
        """Whether ``value`` passes, were its type the column's."""
        if self.op == "blank":
            return value.t == "blank"
        if self.op == "not_blank":
            return value.t != "blank"
        t = value.t
        if t == "blank" or self.op not in OPS_BY_TYPE[t]:
            return False
        wanted = self._value(t, date1904)
        if wanted is _BAD:
            return False
        if t == "list":                     # any of its items
            items = cast("list[str]", value.v)
            if self.op == "eq":
                return wanted in items
            if self.op == "in":
                return any(item in wanted for item in items)
            return any(wanted in item.casefold() for item in items)
        key = _key(value, t, date1904)
        if key is None:
            return False
        if self.op == "eq":
            return key == wanted
        if self.op == "in":
            return key in wanted
        if self.op == "gte":
            return key >= wanted
        if self.op == "lte":
            return key <= wanted
        text = cast(str, value.v)                       # words_gt and contains: on text
        if self.op == "words_gt":
            return len(text.split()) > wanted
        return wanted in text.casefold()             # contains, on text


    def check(self, column_type: str, date1904: bool) -> None:
        """Refuse the test unless the column's type takes it, with a value of that type."""
        allowed = OPS_BY_TYPE[column_type]
        if self.op not in allowed:
            if column_type == "blank":
                reason = "holds no values"
            elif self.op in ("gte", "lte"):
                reason = "is not ordered here: gte and lte compare numbers, dates and times"
            else:
                reason = f"has no text to search: {self.op} is for text columns"
            raise QueryError("wrong_type", f"{self.column} is a {column_type} column and "
                             f"{reason}; it takes {', '.join(allowed)}", column=self.column,
                             column_type=column_type, valid_options=allowed)
        if self.compares and self._value(column_type, date1904) is _BAD:
            what = "each value of in" if self.op == "in" else f"the value of {self.op}"
            form = "a whole number of words (150)" if self.op == "words_gt" \
                else _FORMS[column_type]
            raise QueryError("wrong_type", f"{self.column} is a {column_type} column: {what} "
                             f"is to be {form}", column=self.column,
                             column_type=column_type, valid_options=allowed)


class _Bad(Exception):
    """A value that is not of the type asked for."""


_BAD = object()


def _tests(region: "Region | DatasetTable", where: Any) -> list[_Test]:
    if where is None:
        return []
    out: list[_Test] = []
    if isinstance(where, Mapping):
        for column, ops in where.items():
            if not isinstance(column, str) or not isinstance(ops, Mapping) or not ops:
                raise QueryError("bad_filter", "where maps a column name to its tests, as "
                                 "{\"Price\": {\"gte\": 5}}")
            name = region.columns[region.column_index(column)]
            for op, operand in ops.items():
                out.append(_test(region, name, op, operand, mapping=True))
        return out
    if isinstance(where, (str, bytes)) or not isinstance(where, Iterable):
        raise QueryError("bad_filter", "where is a mapping of columns to tests, or a list of "
                         "{column, op, value} filters")
    for item in where:
        if not isinstance(item, Mapping) or not isinstance(item.get("column"), str) \
                or set(item) - {"column", "op", "value", "values"}:
            raise QueryError("bad_filter", "each filter is {column, op, value} (values for in)")
        op = item.get("op")
        name = region.columns[region.column_index(item["column"])]
        if op == "in":
            if "value" in item:
                raise QueryError("bad_filter", f"in on {name} takes values, a list",
                                 column=name)
            operand = item.get("values")
        elif op in ("blank", "not_blank"):
            if "value" in item or "values" in item:
                raise QueryError("bad_filter", f"{op} on {name} takes no value", column=name)
            operand = True
        else:
            if "values" in item:
                raise QueryError("bad_filter", f"{op} on {name} takes one value, not values",
                                 column=name, valid_options=OPS if op not in OPS else ())
            operand = item.get("value")
        out.append(_test(region, name, op, operand, mapping=False))
    return out


def _test(region: "Region | DatasetTable", name: str, op: Any, operand: Any, *,
          mapping: bool) -> _Test:
    if op not in OPS:
        raise QueryError("bad_filter", f"a filter on {name} has an unknown test; the tests are "
                         f"{', '.join(OPS)}", column=name, valid_options=OPS)
    if op in ("blank", "not_blank"):
        if operand is not True:
            raise QueryError("bad_filter", f"{op} on {name} is written {op}: true; for the "
                             "other, use " + ("not_blank" if op == "blank" else "blank"),
                             column=name)
    elif op == "in":
        if isinstance(operand, (str, bytes, Mapping)) or not isinstance(operand, Iterable):
            raise QueryError("bad_filter", f"in on {name} takes a list of values", column=name)
        operand = list(operand)
        if not operand or len(operand) > MAX_IN:
            raise QueryError("bad_filter", f"in on {name} takes 1 to {MAX_IN} values",
                             column=name)
    elif operand is None:
        raise QueryError("bad_filter", f"{op} on {name} needs a value", column=name)
    elif op == "contains" and operand == "":
        raise QueryError("bad_filter", f"contains on {name} needs some text", column=name)
    return _Test(name, region.column_index(name), op, operand)


def _words(raw: Any) -> int:
    """``words_gt``'s value: a whole number of words, 0 or more (``150`` or ``"150"``);
    :class:`_Bad` when it is not one."""
    if isinstance(raw, int) and not isinstance(raw, bool):
        number = raw
    elif isinstance(raw, str) and re.fullmatch(r"[0-9]+", raw.strip()):
        number = int(raw.strip())
    else:
        raise _Bad
    if number < 0:
        raise _Bad
    return number


def _operand(raw: Any, t: str, date1904: bool) -> Any:
    """A filter's value as a key of type ``t`` (:func:`_key`'s); :class:`_Bad` when it is
    not one."""
    if t in ("number", "percent"):
        if isinstance(raw, bool):
            raise _Bad
        if isinstance(raw, (int, float)):
            number = raw
        elif isinstance(raw, str) and _NUMBER.fullmatch(raw.strip()):
            number = float(raw)
        else:
            raise _Bad
        if not math.isfinite(number):
            raise _Bad
        return number
    if t in _MOMENTS:
        if isinstance(raw, bool):
            raise _Bad
        if isinstance(raw, (int, float)) and math.isfinite(raw) and raw >= 0:
            return _serial_key(float(raw), t, raw >= 1 and t == "time")
        if isinstance(raw, str):
            key = _iso_key(raw.strip(), t, date1904)
            if key is not None:
                return key
        raise _Bad
    if t == "bool":
        if isinstance(raw, bool):
            return raw
        if isinstance(raw, str) and raw.strip().casefold() in ("true", "false"):
            return raw.strip().casefold() == "true"
        raise _Bad
    if not isinstance(raw, str):        # text and error
        raise _Bad
    return raw


def _key(value: Value, t: str, date1904: bool) -> Any:
    """What a value of type ``t`` is compared and sorted by; ``None`` when it has none."""
    if t in _MOMENTS:
        if value.serial is not None:
            elapsed = isinstance(value.v, str) and value.v.startswith("PT")
            return _serial_key(value.serial, t, elapsed)
        return _iso_key(value.v, t, date1904) if isinstance(value.v, str) else None
    return value.v


def _serial_key(serial: float, t: str, elapsed: bool) -> int:
    """A serial as the reader shows it: the day (``date``), the millisecond of the day
    (``time``; of the whole length when ``elapsed``), the millisecond (``datetime``)."""
    total = round(serial * _MS_PER_DAY)
    if t == "date":
        return total // _MS_PER_DAY
    if t == "time" and not elapsed:
        return total % _MS_PER_DAY
    return total


def _clock_ms(hours: str, minutes: str, seconds: str | None, fraction: str | None) -> int | None:
    h, m, s = int(hours), int(minutes), int(seconds or 0)
    if h > 23 or m > 59 or s > 59:
        return None
    ms = round(float(f"0.{fraction}") * 1000) if fraction else 0
    return ((h * 60 + m) * 60 + s) * 1000 + ms


def _iso_key(text: str, t: str, date1904: bool) -> int | None:
    """ISO 8601 text as :func:`_serial_key` keys a serial, in the workbook's date system;
    ``None`` when it is not a value of type ``t`` (a time on a date column, say)."""
    if t == "time":
        match = _ISO_TIME.fullmatch(text)
        if match:
            return _clock_ms(*match.groups())
        match = _DURATION.fullmatch(text)
        if match and text != "PT":
            h, m, s, fraction = match.groups()
            ms = round(float(f"0.{fraction}") * 1000) if fraction else 0
            return ((int(h or 0) * 60 + int(m or 0)) * 60 + int(s or 0)) * 1000 + ms
        return None
    match = _ISO_DAY.fullmatch(text)
    if not match:
        return None
    day = day_serial(int(match.group(1)), int(match.group(2)), int(match.group(3)), date1904)
    if day is None:
        return None
    if t == "date":
        return day
    clock = _clock_ms(*match.groups()[3:]) if match.group(4) else 0
    return None if clock is None else day * _MS_PER_DAY + clock


# -- sorting -----------------------------------------------------------------------------------


class _SortKey:
    def __init__(self, column: str, offset: int, desc: bool) -> None:
        self.column = column
        self.offset = offset
        self.desc = desc

    def key(self, value: Value, date1904: bool) -> tuple[str, Any]:
        """The value's type and its key were that its column's type."""
        t = value.t
        if t == "blank":
            return t, None
        key = _key(value, t, date1904)
        if t == "text":
            key = (key.casefold(), key)
        elif t == "list":
            key = tuple((item.casefold(), item) for item in key)
        return t, key


def _sort_keys(region: "Region | DatasetTable", sort: Any) -> list[_SortKey]:
    if not sort:
        return []
    if isinstance(sort, (str, tuple)):
        sort = [sort]
    out = []
    for item in sort:
        if isinstance(item, str):
            column, order = item, "asc"
        elif isinstance(item, Mapping) and isinstance(item.get("column"), str) \
                and not set(item) - {"column", "desc"}:
            column, order = item["column"], "desc" if item.get("desc") else "asc"
        elif isinstance(item, (tuple, list)) and len(item) == 2 and isinstance(item[0], str):
            column, order = item
        else:
            raise QueryError("bad_sort", "each sort key is a column name or (name, \"asc\" or "
                             "\"desc\")", field="sort")
        if order not in ("asc", "desc"):
            raise QueryError("bad_sort", "a sort order is asc or desc", field="sort",
                             valid_options=["asc", "desc"])
        name = region.columns[region.column_index(column)]
        out.append(_SortKey(name, region.column_index(name), order == "desc"))
    return out


def _ordered(rows: list[tuple], index: int, column_type: str, desc: bool) -> list[tuple]:
    """``rows`` stably sorted by sort key ``index``: the cells of the column's type first, in
    order; then those of other types; then blanks."""
    own, other, blank = [], [], []
    for entry in rows:
        t, key = entry[2][index]
        if t == "blank":
            blank.append(entry)
        elif t == column_type and key is not None:
            own.append(entry)
        else:
            other.append(entry)
    own.sort(key=lambda entry: entry[2][index][1], reverse=desc)
    return own + other + blank

