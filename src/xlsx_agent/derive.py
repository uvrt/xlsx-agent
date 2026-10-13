"""Derived figures: counts, sums, means, shares, groups, distributions and pivots of a
table, name or region, computed here and each traceable to its cells (ROADMAP.md, "Tools":
``xlsx_derive``; "Design rules", 1 to 4)::

    ds = wb.derive(from_="t:Products", where={"Status": {"in": ["Active"]}},
                   group_by=["Category"],
                   measures=[{"fn": "mean", "column": "Price", "as": "average"},
                             {"fn": "count", "as": "items"}],
                   sort=[("average", "desc")], round=2, id="by_category")
    ds.rows               # [{"Category": "Hardware", "average": 41.25, "items": 12}, ...]
    ds.unrounded          # the same rows before rounding
    ds.sources["Hardware"]["average"]   # [RangeAddress(Stock!F3:F14), ...]: the cells used
    ds.excluded           # {"average": {"blank": Count(2, ("Stock!F9", ...))}}

The model never does arithmetic: every figure an agent reports comes from here or from a
cell, and each one says which cells it came from.

**Rows.**  ``from_`` is a table, defined name or region (``t:`` address, or a
:class:`~xlsx_agent.read.regions.Region`); its data rows are read (header and totals rows
excluded).  Or a dataset (R11: a :class:`Derived`, an ooxml-edit
:class:`~ooxml_edit.tools.Dataset` or a :class:`~xlsx_agent.dataset.DatasetTable`; in a
tool session, ``d:<id>``): its rows, the unrounded values of a :class:`Derived`, with
every value's sources carried through to the original cells
(:mod:`xlsx_agent.dataset`).
``where`` is :mod:`xlsx_agent.query`'s filters, with the same validation and
refusals (:class:`~xlsx_agent.query.QueryError`), checked against the same column types
(the type most of a column's values have, :func:`~xlsx_agent.profile.majority_type`, over
all its data rows); a cell of another type than its column's never passes a test that
compares values, and is counted in :attr:`Derived.not_compared`.  Rows from hidden rows,
columns and sheets are used like any other and counted in :attr:`Derived.hidden`.

**Groups.**  ``group_by`` names columns; a row's group is its cells in them, by type and
value (``1`` and ``"1"`` are two groups; ``3`` and ``3.0`` one; text exactly, case
counting, as ``eq`` matches; a date by its day, a datetime by its millisecond; a blank
cell is a group of its own, shown as ``None``).  A row whose group cell is a formula
without a cached value has no group: it is left out and counted
(:attr:`Derived.left_out`), never guessed.  Groups come in the order first met in the
table.  Without ``group_by`` the whole table is one row.

**Measures**, each ``{"fn": ..., "column": ..., "as": ...}`` (``as`` defaults to ``fn``,
or ``"fn column"``):

=========== ==================================== =======================================
fn          column types                         value
=========== ==================================== =======================================
count       any; no column: the rows             how many values (rows)
distinct    any                                  how many different values
sum         number, percent                      the exact sum, rounded once
mean        number, percent, date, time,         the exact sum over the count, rounded
            datetime                             once
median      as mean                              the middle value; of an even count the
                                                 exact mean of the two middle ones
min, max    as mean                              the least, the greatest value
share       number, percent; no column: the rows this group's sum (row count) as a
                                                 fraction of the total over every row
                                                 used (:attr:`Derived.totals`)
=========== ==================================== =======================================

A measure on a column whose type is not listed is refused before anything is computed
(:class:`DeriveError`, ``wrong_type``, naming the column, its type and the functions that
take it): text has no ``min``/``max`` (Excel's collation is not measured), dates no
``sum``.  A measure takes the cells of its column's type only; nothing is coerced
silently (ROADMAP.md, "Design rules", 4).  Every other cell of the group's rows is left
out of that measure and counted by reason: ``blank``, ``no_cached_value``, ``error``,
``number_as_text`` (text that reads as a number in the workbook's locale), ``wrong_type``
(any other).  ``coerce=["number_as_text"]`` opts in to reading such text in a ``number``
column as its number (:func:`~xlsx_agent.read.types.text_number`); the result says so,
with the cells (:attr:`Derived.coerced`).  A percent column's text stays out: its stored
values are fractions, and text such as ``25`` is not one.

Dates, times and datetimes are measured as :mod:`xlsx_agent.query` compares them: a date
by its day, a time by its millisecond of the day (an elapsed time, ``PT36H0M0S``, by its
length), a datetime by its millisecond, in the workbook's date system; a mean or median
is shown as the reader shows a value of that type (ISO 8601; a date's mean as its day).

**Exact arithmetic.**  A sum is :func:`math.fsum`'s (kept as Shewchuk's partials while
streaming, which hold the sum exactly); a mean, a median of an even count and a share are
computed from those exact sums with :class:`fractions.Fraction` and rounded once, to the
nearest double.  So every value is the exact result of its cells' stored doubles, rounded
once (where ``sum(x) / n`` would round twice).  No measure of no values has a value
(``None``), except ``count`` and ``distinct`` (``0``); a share of a total of 0 is ``None``.

**Rounding** (ROADMAP.md, "Design rules", 3).  ``round`` is a number of decimals (``2``;
``-1`` to tens; at most :data:`MAX_DIGITS` either way), or ``{"as": digits}`` per
measure (each of a pivoted measure's columns alike).  It is **half away from zero** (``2.5`` to ``3``, ``-2.5`` to ``-3``), applied to
the decimal Python shows for the double (:func:`repr`, the shortest that reads back as
it), so ``2.675`` (stored as 2.67499999999999982236431605997495353221893310546875) rounds
to ``2.68``, as it reads.  ROADMAP.md names the rule after Excel's ``ROUND``; whether
Excel rounds every double the same way is **not measured**: no parity is claimed.
It applies to number and percent values (a share is a fraction: ``round=2`` gives
``0.33``); counts are whole and dates are not rounded.  It changes the reported value only:
:attr:`Derived.rows` is rounded, :attr:`Derived.unrounded` keeps the values computed.

**Provenance** (ROADMAP.md, "Design rules", 2).  Every value has its sources
(:attr:`Derived.provenance`): the cells it was computed from, exactly -- the measure's
cells in the group's rows that were used -- as compact ranges, a run of consecutive rows
being one range (``Stock!F3:F14``); how many; and how many cells of the group's rows were
left out, by reason.  A group's key value has the key column's cells of its rows; a
``count`` or ``share`` of rows has the rows' cells in the first group column (else the
bins' column, the pivot's, or the table's first).  So a check can read each value's
sources and recompute it, and ``count + left out`` is the group's rows: nothing is dropped
silently.  ``min`` and ``max`` list every cell compared, not only the extreme one.

**Bins** (distributions): ``bins={"column": "Price", "edges": [0, 10, 50], "as": ...}``
on a number, percent, date, time or datetime column; edges as filter values are written
(:mod:`xlsx_agent.query`: numbers, ISO 8601 text or serials), strictly increasing.  Each
bin holds its lower edge and not its upper one (``[0, 10)``), the last both (``[10,
50]``).  A row whose cell lies outside every bin, or is not of the column's type, is left
out and counted (``outside_bins``, or the reason as for a measure).  The bin is a key
column after ``group_by``'s, shown as its interval (``"[0, 10)"``, edges as numbers are
written, ``number_text``, or ISO 8601); every bin is listed for every group, an empty one
with ``count`` 0 and no other value.

**Pivot**: ``pivot="Region"``: each value of that column becomes columns, one per measure
(named by the value, ``North``, with one measure; ``North sales`` with more), in the order
first met; a group without rows for a value has ``count`` 0 and no other value.  A
value's name is its text, or a number as written (``number_text``), a date's ISO 8601,
``true``/``false``, an error's code, ``(blank)``; names that would clash get the type
(``1 (text)``) or a number.  :attr:`Derived.pivot_columns` maps each column to its value.

**Sort and top.**  ``sort`` is a list of output columns, each ``"name"`` (ascending) or
``("name", "asc" | "desc")``, the first deciding first; values without a value (``None``,
a blank group) come last either way; a key column of mixed types sorts by type first (in
the order of :data:`~xlsx_agent.read.types.TYPES`), text by :meth:`str.casefold` then as
written; a bin by its order; a date by its day.  Ties keep the groups' order.  The
unrounded values are sorted.  ``top`` keeps the first that many rows after sorting;
:attr:`Derived.groups` counts them all.

**Bounded.**  Two streamed passes over the columns concerned (the first counts their
types, so that every refusal comes before anything is computed; the second computes).
What is held is per group, never per row: the measures' running state (exact sums as
partials, extremes, counts), the group's rows as runs (a run of consecutive rows is two
numbers), and only for ``median`` its numbers and for ``distinct`` its keys.  So the
memory is bounded by the groups when a group's rows are consecutive; when they
interleave, the runs grow with the rows (16 bytes a run), up to :data:`MAX_RUNS`.  The
limits refuse with :class:`DeriveError`: :data:`MAX_GROUPS` groups (``too_many_groups``),
:data:`MAX_PIVOT_VALUES` pivot values (``too_many_columns``), :data:`MAX_ROWS` rows after
``top`` (``too_many_rows``) and :data:`MAX_CELLS` values (``too_many_cells``),
:data:`MAX_RUNS` runs of sources (``too_many_sources``), :data:`MAX_KEPT` numbers kept for
medians and keys for distinct counts (``too_many_values``).

**Content** (ROADMAP.md, "Design rules", 7).  Group keys, pivot values and the values
computed are content: they appear in results only.  Errors and their messages name
columns, functions, types, codes and counts, never a cell's value, a key or a filter's
value.
"""

from __future__ import annotations

import bisect
import hashlib
import math
from array import array
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal, localcontext
from fractions import Fraction
from typing import TYPE_CHECKING, Any, Iterable, Mapping, Sequence

from ooxml_edit.charts.workbook import Area, Cell, number_text
from ooxml_edit.tools import Dataset

from .address import DatasetAddress, RangeAddress, SheetRef, TableAddress, parse_address
from .dataset import DatasetTable, as_dataset
from .profile import FIRST_ADDRESSES, Count, _add_partial, majority_type
from .query import QueryError, _Bad, _key, _operand, _tests
from .read.regions import Region
from .read.types import TYPES, Value, serial_text, text_number

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Workbook

#: Every measure function.
FUNCTIONS = ("count", "sum", "mean", "median", "min", "max", "share", "distinct")

_NUMBERS = ("number", "percent")
_MOMENTS = ("date", "time", "datetime")
_ORDERED = _NUMBERS + _MOMENTS

#: The column types each function takes.
TYPES_BY_FUNCTION: dict[str, tuple[str, ...]] = {
    "count": TYPES, "distinct": TYPES, "sum": _NUMBERS, "share": _NUMBERS,
    "mean": _ORDERED, "median": _ORDERED, "min": _ORDERED, "max": _ORDERED}

#: The functions that take no column (they count rows).
ROW_FUNCTIONS = ("count", "share")

#: What a caller may opt in to coerce.
COERCIONS = ("number_as_text",)

#: Why a cell is left out of a measure, or a row out of the groups.
REASONS = ("blank", "no_cached_value", "error", "number_as_text", "wrong_type",
           "outside_bins")

#: The rounding rule (ROADMAP.md, "Design rules", 3).
ROUNDING = "half_away_from_zero"

#: The most groups (output rows before ``top``; with bins, groups times bins).
MAX_GROUPS = 10_000
#: The most result rows after ``top`` (ROADMAP.md, "Tools": ``xlsx_derive``).
MAX_ROWS = 1_000
#: The most values in a result (rows times columns).
MAX_CELLS = 50_000
#: The most values a pivot column holds (each a column per measure).
MAX_PIVOT_VALUES = 100
#: The most measures in one derive.
MAX_MEASURES = 50
#: The most bin edges.
MAX_EDGES = 1_001
#: The most runs of source rows held, over every value.
MAX_RUNS = 1_000_000
#: The most numbers (medians) and keys (distinct counts) held, over every value.
MAX_KEPT = 1_000_000
#: The most decimals ``round`` takes, either way.
MAX_DIGITS = 15

_MS_PER_DAY = 86_400_000
# Text longer than this is kept as a digest in a distinct set.
_KEY_TEXT = 64


class DeriveError(QueryError):
    """A derive the library refuses: ``code`` says why (``wrong_type``, ``bad_source``,
    ``bad_measure``, ``bad_group``, ``bad_bins``, ``bad_pivot``, ``bad_sort``, ``bad_top``,
    ``bad_round``, ``bad_coerce``, ``duplicate_name``, and the limits: ``too_many_groups``,
    ``too_many_columns``, ``too_many_rows``, ``too_many_cells``, ``too_many_sources``,
    ``too_many_values``); ``field`` the argument at fault.  A :class:`QueryError`, as a
    ``where`` refusal is: the message names columns, functions and types, never a value."""


class _Sourced:
    """Cells of one column as runs of rows, kept compact (two numbers a run) and given as
    range addresses when asked for."""

    __slots__ = ()

    @property
    def sources(self) -> "tuple[RangeAddress | str, ...]":
        """The cells, as compact ranges: a run of consecutive rows is one range
        (``Stock!F3:F14``), in row order.  Derived from a dataset (``via``), the cells its
        values came from (:meth:`DatasetTable.sources_of
        <xlsx_agent.dataset.DatasetTable.sources_of>`)."""
        if self.via is not None:
            return self.via.sources_of(self.column, self.runs)
        b, sheet, column = self.runs, self.sheet, self.column
        return tuple(RangeAddress(sheet, Area(Cell(b[i], column), Cell(b[i + 1], column)))
                     for i in range(0, len(b), 2))

    @property
    def ranges(self) -> int:
        """How many ranges :attr:`sources` gives."""
        if self.via is not None:
            return len(self.sources)
        return len(self.runs) // 2


@dataclass(frozen=True)
class Provenance(_Sourced):
    """Where one value came from: :attr:`sources` the cells it was computed from, as
    compact ranges; ``count`` how many cells those are; ``excluded`` how many cells of the
    group's rows were left out of it, by reason.  ``sheet``, ``column`` and ``runs`` (first
    and last rows, flat) are what :attr:`sources` is made from; ``via`` the dataset it was
    derived from, if any (then ``column`` is the dataset's, from 0, ``runs`` its positions,
    and ``count`` how many of its values were used)."""

    sheet: SheetRef | None
    column: int
    runs: array = field(repr=False)
    count: int
    excluded: dict[str, int] = field(default_factory=dict)
    via: "DatasetTable | None" = field(default=None, repr=False, compare=False)

    def as_dict(self) -> dict:
        out: dict = {"sources": [str(s) for s in self.sources], "count": self.count}
        if self.excluded:
            out["excluded"] = dict(self.excluded)
        return out


@dataclass(frozen=True)
class Total(_Sourced):
    """The total a ``share`` is a fraction of: ``value`` (the exact sum, rounded once; a
    count of rows), with its cells as :class:`Provenance` gives them."""

    value: float | int
    sheet: SheetRef | None
    column: int
    runs: array = field(repr=False)
    count: int
    via: "DatasetTable | None" = field(default=None, repr=False, compare=False)

    def as_dict(self) -> dict:
        return {"value": self.value, "sources": [str(s) for s in self.sources],
                "count": self.count}


@dataclass(frozen=True)
class Derived:
    """A derive's result (see the module's notes).

    ``address`` the table's (``t:Products``); ``id`` as given; ``columns`` the output
    columns in order (key columns first: ``group_by``'s, then the bins'); ``key_columns``
    how many of them are keys; ``types`` each column's type (a key column's is its
    column's; a bin's ``text``; ``count`` and ``distinct`` ``number``; ``share``
    ``percent``; the others their column's); ``rows`` the values, rounded, one mapping per
    row; ``unrounded`` the same before rounding; ``provenance`` per row, per column, its
    :class:`Provenance`; ``rounding`` the rule and the decimals per column.

    ``matched`` how many data rows passed ``where``; ``used`` how many of those were
    grouped; ``groups`` how many result rows there were before ``top``; ``left_out`` per
    column, the rows not grouped by reason (a group or pivot cell without a cached value; a
    bin cell outside the bins or of the wrong type); ``excluded`` per measure (by ``as``),
    the cells left out by reason, each with its first addresses; ``coerced`` per measure
    (and the bins, by their ``as``), the cells read as numbers by ``coerce``;
    ``not_compared`` per filtered column, as :class:`~xlsx_agent.query.QueryResult` gives
    it; ``hidden`` the rows used that hold a
    hidden cell among the columns read; ``totals`` per ``share`` measure, its total;
    ``pivot_columns`` per pivoted column, its measure and pivot value; ``date1904`` the
    date system its dates are in (its source's).

    :meth:`dataset` gives it as an ooxml-edit :class:`~ooxml_edit.tools.Dataset`, what a
    tool session keeps under ``d:<id>`` (:mod:`xlsx_agent.dataset`); ``wb.derive(ds, ...)``
    derives from it."""

    address: str
    id: str | None
    columns: tuple[str, ...]
    key_columns: int
    types: dict[str, str]
    rows: list[dict[str, Any]]
    unrounded: list[dict[str, Any]]
    provenance: list[dict[str, Provenance]] = field(repr=False)
    rounding: dict[str, Any] = field(default_factory=dict)
    matched: int = 0
    used: int = 0
    groups: int = 0
    left_out: dict[str, dict[str, Count]] = field(default_factory=dict)
    excluded: dict[str, dict[str, Count]] = field(default_factory=dict)
    coerced: dict[str, Count] = field(default_factory=dict)
    not_compared: dict[str, Count] = field(default_factory=dict)
    hidden: Count = field(default_factory=lambda: Count(0))
    totals: dict[str, Total] = field(default_factory=dict)
    pivot_columns: dict[str, dict[str, Any]] = field(default_factory=dict)
    date1904: bool = False

    def dataset(self) -> "Dataset":
        """This result as an ooxml-edit :class:`~ooxml_edit.tools.Dataset`: the rows as
        reported (rounded), each value's sources as range text, the key columns as
        ``keys``, the types as labels, and this result as ``detail``
        (:func:`xlsx_agent.dataset.as_dataset`)."""
        return as_dataset(self)

    @property
    def sources(self) -> dict[Any, dict[str, list[RangeAddress]]]:
        """Each row's sources by its key (the key column's value; a tuple of them with
        several key columns; ``None`` without any), then by column.  Two keys that show
        alike (a number and a bool, a date and its ISO text) keep the first row's: by
        position, :attr:`provenance` has every one."""
        out: dict[Any, dict[str, list[RangeAddress]]] = {}
        keys = self.columns[:self.key_columns]
        for row, where in zip(self.unrounded, self.provenance):
            if not keys:
                key = None
            elif len(keys) == 1:
                key = row[keys[0]]
            else:
                key = tuple(row[k] for k in keys)
            try:
                out.setdefault(key, {c: list(p.sources) for c, p in where.items()})
            except TypeError:  # pragma: no cover - keys are scalars
                continue
        return out

    def as_dict(self) -> dict:
        """The result as plain data: rows as lists in column order, sources as range text."""
        out: dict = {"from": self.address, "columns": list(self.columns),
                     "key_columns": self.key_columns, "types": dict(self.types),
                     "rows": [[row[c] for c in self.columns] for row in self.rows],
                     "rounding": dict(self.rounding),
                     "provenance": [[where[c].as_dict() for c in self.columns]
                                    for where in self.provenance],
                     "matched": self.matched, "used": self.used, "groups": self.groups}
        if self.id is not None:
            out["id"] = self.id
        if self.rows != self.unrounded:
            out["unrounded"] = [[row[c] for c in self.columns] for row in self.unrounded]
        for name in ("left_out", "excluded"):
            value = getattr(self, name)
            if value:
                out[name] = {k: {r: c.as_dict() for r, c in v.items()}
                             for k, v in value.items()}
        for name in ("coerced", "not_compared"):
            value = getattr(self, name)
            if value:
                out[name] = {k: c.as_dict() for k, c in value.items()}
        if self.hidden.count:
            out["hidden"] = self.hidden.as_dict()
        if self.totals:
            out["totals"] = {k: t.as_dict() for k, t in self.totals.items()}
        if self.pivot_columns:
            out["pivot_columns"] = {k: dict(v) for k, v in self.pivot_columns.items()}
        return out


def derive(workbook: "Workbook | None",
           from_: "str | TableAddress | Region | Derived | Dataset | DatasetTable", *,
           where: Any = None, group_by: "Sequence[str] | str | None" = None,
           measures: "Sequence[Mapping[str, Any]] | None" = None,
           bins: "Mapping[str, Any] | None" = None, pivot: str | None = None,
           sort: Any = None, top: int | None = None,
           round: "int | Mapping[str, int] | None" = None,
           coerce: "Sequence[str] | None" = None, id: str | None = None) -> Derived:
    """Figures from a table, name or region, or from a dataset (``workbook`` may then be
    ``None``); see the module's notes."""
    region = _region(workbook, from_)
    plan = _Plan(workbook, region, where=where, group_by=group_by, measures=measures,
                 bins=bins, pivot=pivot, coerce=coerce, id=id)
    plan.check_types(_type_pass(workbook, region, plan.typed_offsets))
    state = plan.run()
    return plan.finish(state, sort=sort, top=top, round=round)


# ------------------------------------------------------------------------------------------
# Arguments
# ------------------------------------------------------------------------------------------


def _region(workbook: "Workbook | None", from_: Any) -> "Region | DatasetTable":
    if isinstance(from_, (Region, DatasetTable)):
        return from_
    if isinstance(from_, (Derived, Dataset)):
        return DatasetTable(from_)
    if isinstance(from_, str):
        try:
            address = parse_address(from_)
        except ValueError:
            address = None
        if isinstance(address, DatasetAddress):
            raise DeriveError("bad_source", "a d: address names a dataset in a tool session; "
                              "here, give the dataset itself (a Derived or a Dataset)",
                              field="from")
        if address is not None and not isinstance(address, TableAddress):
            raise DeriveError("bad_source", "from is a table, name or region (t:Products)",
                              field="from")
        from_ = address if address is not None else from_
    if isinstance(from_, TableAddress):
        if from_.column is not None or from_.key is not None or from_.position is not None:
            raise DeriveError("bad_source", "from names a whole table, name or region, not a "
                              "column or row of one", field="from")
        if workbook is None:
            raise DeriveError("bad_source", "a table is derived from in its workbook",
                              field="from")
        return workbook.region(from_)
    if isinstance(from_, str) and workbook is not None:
        return workbook.region(from_)
    raise DeriveError("bad_source", "from is a table, name or region (t:Products), or a "
                      "dataset", field="from")


@dataclass
class _Measure:
    fn: str
    column: str | None      # as the region spells it
    offset: int | None
    name: str               # as


@dataclass
class _Bins:
    column: str
    offset: int
    raw: list
    name: str
    edges: list = field(default_factory=list)   # keys, once the type is known
    labels: list = field(default_factory=list)


def _names(value: Any, what: str, code: str) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, Iterable) and not isinstance(value, (bytes, Mapping)) \
            and all(isinstance(v, str) for v in value):
        return list(value)
    raise DeriveError(code, f"{what} is a column name or a list of them", field=what)


class _Plan:
    """The arguments checked, and what the passes need."""

    def __init__(self, workbook: "Workbook | None", region: "Region | DatasetTable", *,
                 where, group_by, measures, bins, pivot, coerce, id) -> None:
        self.workbook = workbook
        self.region = region
        self.via = region if isinstance(region, DatasetTable) else None
        self.date1904 = region.date1904 if self.via is not None else workbook.date1904
        if id is not None:
            try:
                DatasetAddress(id=id)
            except (TypeError, ValueError):
                raise DeriveError("bad_source", "id names a dataset, as a table name is "
                                  "written (by_category)", field="id") from None
        self.id = id
        self.tests = _tests(region, where)
        self.group = [region.columns[region.column_index(n)]
                      for n in _names(group_by, "group_by", "bad_group")]
        if len({g.casefold() for g in self.group}) != len(self.group):
            raise DeriveError("bad_group", "group_by names each column once", field="group_by")
        self.group_offsets = [region.column_index(n) for n in self.group]
        coerce = _names(coerce, "coerce", "bad_coerce")
        for item in coerce:
            if item not in COERCIONS:
                raise DeriveError("bad_coerce", f"coerce takes {', '.join(COERCIONS)}",
                                  field="coerce", valid_options=COERCIONS)
        self.coerce = "number_as_text" in coerce
        self.measures = self._measures(measures)
        self.bins = self._bins(bins)
        self.pivot: str | None = None
        self.pivot_offset: int | None = None
        if pivot is not None:
            if not isinstance(pivot, str):
                raise DeriveError("bad_pivot", "pivot is one column name", field="pivot")
            self.pivot = region.columns[region.column_index(pivot)]
            self.pivot_offset = region.column_index(self.pivot)
            if self.pivot in self.group or (self.bins and self.bins.column == self.pivot):
                raise DeriveError("bad_pivot", f"{self.pivot} is grouped already; a pivot "
                                  "column is another", field="pivot")
            if not self.measures:
                raise DeriveError("bad_pivot", "a pivot spreads measures over columns: give "
                                  "at least one measure", field="pivot")
        names = list(self.group) + ([self.bins.name] if self.bins else []) \
            + ([] if self.pivot else [m.name for m in self.measures])
        folded = [n.casefold() for n in names]
        if len(set(folded)) != len(folded):
            raise DeriveError("duplicate_name", "every output column needs its own name: "
                              "give the measures (or bins) another as", field="measures")
        if self.pivot and len({m.name.casefold() for m in self.measures}) != len(self.measures):
            raise DeriveError("duplicate_name", "every measure needs its own as",
                              field="measures")
        # the column a row is counted in, for count and share of rows
        if self.group_offsets:
            self.row_offset = self.group_offsets[0]
        elif self.bins:
            self.row_offset = self.bins.offset
        elif self.pivot_offset is not None:
            self.row_offset = self.pivot_offset
        else:
            self.row_offset = 0
        # the measured columns, each once: offset -> what its accumulator keeps
        self.columns: dict[int, set[str]] = {}
        for m in self.measures:
            if m.offset is not None:
                self.columns.setdefault(m.offset, set()).add(m.fn)
        self.typed_offsets = sorted({t.offset for t in self.tests} | set(self.columns)
                                    | set(self.group_offsets)
                                    | ({self.bins.offset} if self.bins else set())
                                    | ({self.pivot_offset} if self.pivot else set()))

    def _measures(self, measures: Any) -> list[_Measure]:
        if measures is None:
            measures = [{"fn": "count"}]
        if isinstance(measures, (str, bytes, Mapping)) or not isinstance(measures, Iterable):
            raise DeriveError("bad_measure", "measures is a list of {fn, column, as}",
                              field="measures")
        measures = list(measures)
        if not measures or len(measures) > MAX_MEASURES:
            raise DeriveError("bad_measure", f"measures takes 1 to {MAX_MEASURES} measures",
                              field="measures")
        out = []
        for item in measures:
            if not isinstance(item, Mapping) or set(item) - {"fn", "column", "as"}:
                raise DeriveError("bad_measure", "each measure is {fn, column, as}",
                                  field="measures")
            fn = item.get("fn")
            if fn not in FUNCTIONS:
                raise DeriveError("bad_measure", f"a measure's fn is one of {', '.join(FUNCTIONS)}",
                                  field="measures", valid_options=FUNCTIONS)
            column = item.get("column")
            if column is None:
                if fn not in ROW_FUNCTIONS:
                    raise DeriveError("bad_measure", f"{fn} needs a column; only "
                                      f"{' and '.join(ROW_FUNCTIONS)} count rows without one",
                                      field="measures")
                offset = None
            elif not isinstance(column, str):
                raise DeriveError("bad_measure", "a measure's column is a column name",
                                  field="measures")
            else:
                offset = self.region.column_index(column)
                column = self.region.columns[offset]
            name = item.get("as")
            if name is None:
                name = fn if column is None else f"{fn} {column}"
            if not isinstance(name, str) or not name.strip():
                raise DeriveError("bad_measure", "a measure's as is a name for its column",
                                  field="measures")
            out.append(_Measure(fn, column, offset, name))
        return out

    def _bins(self, bins: Any) -> _Bins | None:
        if bins is None:
            return None
        if not isinstance(bins, Mapping) or set(bins) - {"column", "edges", "as"} \
                or not isinstance(bins.get("column"), str):
            raise DeriveError("bad_bins", "bins is {column, edges, as}", field="bins")
        edges = bins.get("edges")
        if isinstance(edges, (str, bytes, Mapping)) or not isinstance(edges, Iterable):
            raise DeriveError("bad_bins", "bins' edges is a list of values", field="bins")
        edges = list(edges)
        if not 2 <= len(edges) <= MAX_EDGES:
            raise DeriveError("bad_bins", f"bins takes 2 to {MAX_EDGES} edges", field="bins")
        offset = self.region.column_index(bins["column"])
        column = self.region.columns[offset]
        name = bins.get("as", column)
        if not isinstance(name, str) or not name.strip():
            raise DeriveError("bad_bins", "bins' as is a name for its column", field="bins")
        if column in self.group:
            raise DeriveError("bad_bins", f"{column} is grouped already; bins group by "
                              "ranges of another column", field="bins")
        return _Bins(column, offset, edges, name)

    # -- the types ------------------------------------------------------------------------

    def check_types(self, types: dict[int, str]) -> None:
        """Refuse what the columns' types do not take, before anything is computed."""
        region = self.region
        self.types = {region.columns[o]: t for o, t in types.items()}
        self.offset_types = types
        for test in self.tests:
            test.check(self.types[test.column], self.date1904)
        for m in self.measures:
            if m.offset is None:
                continue
            t = types[m.offset]
            allowed = TYPES_BY_FUNCTION[m.fn]
            if t not in allowed:
                takes = [fn for fn in FUNCTIONS if t in TYPES_BY_FUNCTION[fn]]
                if t == "blank":
                    reason = "holds no values"
                elif t == "text" and m.fn in ("min", "max"):
                    reason = "has no measured order (text collation is not measured)"
                else:
                    reason = f"is not what {m.fn} takes ({', '.join(allowed)})"
                raise DeriveError("wrong_type", f"{m.column} is a {t} column and {reason}; "
                                  f"it takes {', '.join(takes)}", field="measures",
                                  column=m.column, column_type=t, valid_options=takes)
        if self.bins:
            b = self.bins
            t = types[b.offset]
            if t not in _ORDERED:
                raise DeriveError("wrong_type", f"{b.column} is a {t} column; bins take a "
                                  f"number, percent, date, time or datetime column",
                                  field="bins", column=b.column, column_type=t,
                                  valid_options=list(_ORDERED))
            keys = []
            for raw in b.raw:
                try:
                    keys.append(_operand(raw, t, self.date1904))
                except _Bad:
                    raise DeriveError("wrong_type", f"{b.column} is a {t} column: each edge is "
                                      "a value of its type, as filters are written",
                                      field="bins", column=b.column, column_type=t) from None
            if any(a >= b2 for a, b2 in zip(keys, keys[1:])):
                raise DeriveError("bad_bins", "bins' edges are to increase strictly",
                                  field="bins", column=b.column)
            b.edges = keys
            texts = [_edge_text(k, t, self.date1904) for k in keys]
            b.labels = [f"[{lo}, {hi})" for lo, hi in zip(texts, texts[1:])]
            b.labels[-1] = b.labels[-1][:-1] + "]"

    def run(self) -> "_State":
        return _second_pass(self)

    def finish(self, state: "_State", *, sort: Any, top: Any, round: Any) -> Derived:
        return _finish(self, state, sort, top, round)


def _edge_text(key: Any, t: str, date1904: bool) -> str:
    if t in _NUMBERS:
        return number_text(key)
    return serial_text(_serial(Fraction(key), t), t, date1904) or str(key)


def _serial(key: Fraction, t: str) -> float:
    """A key (:func:`xlsx_agent.query._key`'s) as a serial: a date's is its day, a time's
    and a datetime's milliseconds."""
    return float(key) if t == "date" else float(key / _MS_PER_DAY)


# ------------------------------------------------------------------------------------------
# The passes
# ------------------------------------------------------------------------------------------


def _area(region: Region, offsets: Iterable[int]) -> Area | None:
    data = region.data
    offsets = list(offsets) or [0]
    if data is None:
        return None
    first = region.area.first.column
    return Area(Cell(data.first.row, first + min(offsets)),
                Cell(data.last.row, first + max(offsets)))


def _rows(workbook: "Workbook | None", region: "Region | DatasetTable", offsets: list[int]
          ) -> tuple[Iterable[list[Value]], int] | None:
    """The rows to read for the columns at ``offsets``, streamed, and the offset of each
    row's first value; ``None`` when there are no data rows.  A dataset gives every
    column, measured (unrounded)."""
    if isinstance(region, DatasetTable):
        return region.measured_rows(), 0
    area = _area(region, offsets)
    if area is None:
        return None
    left = area.first.column - region.area.first.column
    return workbook.read(RangeAddress(region.sheet_ref, area)), left


def _type_pass(workbook: "Workbook | None", region: "Region | DatasetTable",
               offsets: list[int]) -> dict[int, str]:
    """Each column's type, from one streamed pass counting its values' types."""
    found = _rows(workbook, region, offsets)
    counts: dict[int, dict[str, int]] = {o: {} for o in offsets}
    if found is not None and offsets:
        rows, left = found
        for row in rows:
            for o in offsets:
                t = row[o - left].t
                if t != "blank":
                    tally = counts[o]
                    tally[t] = tally.get(t, 0) + 1
    return {o: majority_type(c) for o, c in counts.items()}


class _Counter:
    __slots__ = ("count", "first")

    def __init__(self) -> None:
        self.count = 0
        self.first: list[str] = []

    def add(self, value: Value) -> None:
        self.count += 1
        if len(self.first) < FIRST_ADDRESSES:
            self.first.append(value.a)

    def done(self) -> Count:
        return Count(self.count, tuple(self.first))


class _Runs:
    """Rows as runs: a flat array of (first, last) pairs."""

    __slots__ = ("bounds",)

    def __init__(self) -> None:
        self.bounds = array("l")

    def add(self, row: int) -> int:
        """Add a row (rows come in order); 1 when it starts a run."""
        bounds = self.bounds
        if bounds and bounds[-1] == row - 1:
            bounds[-1] = row
            return 0
        bounds.append(row)
        bounds.append(row)
        return 1


class _Acc:
    """One measured column's state in one group (and pivot value)."""

    __slots__ = ("runs", "count", "partials", "extremes", "low", "high", "kept", "seen",
                 "excluded", "elapsed")

    def __init__(self, fns: set[str]) -> None:
        self.runs = _Runs()
        self.count = 0
        self.partials: list[float] | None = [] if fns & {"sum", "mean", "share"} else None
        self.extremes = bool(fns & {"min", "max"})
        self.low: tuple | None = None
        self.high: tuple | None = None
        self.kept = array("d") if "median" in fns else None
        self.seen: set | None = set() if "distinct" in fns else None
        self.excluded: dict[str, int] = {}
        self.elapsed = True


class _Cell:
    """A group's state for one pivot value (or the group's, without a pivot)."""

    __slots__ = ("runs", "count", "accs")

    def __init__(self, columns: dict[int, set[str]]) -> None:
        self.runs = _Runs()
        self.count = 0
        self.accs = {o: _Acc(fns) for o, fns in columns.items()}


class _Group:
    __slots__ = ("shown", "runs", "count", "cells")

    def __init__(self, shown: list) -> None:
        self.shown = shown
        self.runs = _Runs()
        self.count = 0
        self.cells: dict[Any, _Cell] = {}


@dataclass
class _State:
    groups: dict = field(default_factory=dict)          # key -> _Group
    order: list = field(default_factory=list)            # group_by keys, first met
    met: set = field(default_factory=set)                # the same, as a set
    pivots: dict = field(default_factory=dict)           # pivot key -> (v, t)
    matched: int = 0
    used: int = 0
    runs: int = 0
    kept: int = 0
    left_out: dict = field(default_factory=dict)         # column -> reason -> _Counter
    excluded: dict = field(default_factory=dict)         # offset -> reason -> _Counter
    coerced: dict = field(default_factory=dict)          # offset -> _Counter
    not_compared: dict = field(default_factory=dict)     # column -> _Counter
    hidden: _Counter = field(default_factory=_Counter)
    totals: dict = field(default_factory=dict)           # offset or None -> [partials, runs, count]


def _group_key(value: Value, date1904: bool) -> tuple:
    """A cell as a group's key: its type and value (a date by its day...)."""
    t = value.t
    if t == "blank":
        return (t, None)
    if t in _MOMENTS:
        return (t, _key(value, t, date1904))
    return (t, value.v)


def _reason(value: Value, column_type: str) -> str:
    if value.issue == "no_cached_value":
        return "no_cached_value"
    if value.t == "blank":
        return "blank"
    if value.t == "error":
        return "error"
    if value.issue == "number_as_text" and column_type in _NUMBERS:
        return "number_as_text"
    return "wrong_type"


def _second_pass(plan: _Plan) -> _State:
    workbook, region = plan.workbook, plan.region
    date1904 = plan.date1904
    state = _State()
    offsets = plan.typed_offsets + [plan.row_offset]
    found = _rows(workbook, region, offsets)
    if found is None:
        return state
    rows, left = found
    types = plan.offset_types
    typed_tests = [(t, types[t.offset]) for t in plan.tests]
    # a dataset's values are typed already: no text is read as a number there
    locale, grouping = (workbook.locale, workbook.grouping) if workbook is not None \
        else (None, True)
    columns = plan.columns
    shares = {m.offset for m in plan.measures if m.fn == "share"}
    for o in shares:
        state.totals[o] = [[], _Runs(), 0]
    bins = plan.bins
    coerced_bins = _Counter()
    read = set(offsets)

    def leave(column: str, reason: str, value: Value) -> None:
        counter = state.left_out.setdefault(column, {}).setdefault(reason, _Counter())
        counter.add(value)

    def runs(n: int) -> None:
        state.runs += n
        if state.runs > MAX_RUNS:
            raise DeriveError("too_many_sources", f"the sources would take more than "
                              f"{MAX_RUNS} ranges; filter the rows (where) or group "
                              "them so that each group's rows are fewer runs",
                              field="from")

    for row in rows:
        # the filters: as a query's, a cell of another type than its column's never passes
        passed = True
        for test, t in typed_tests:
            value = row[test.offset - left]
            if test.compares and value.t not in ("blank", t):
                state.not_compared.setdefault(test.column, _Counter()).add(value)
                passed = False
            elif passed and not test.passes(value, date1904):
                passed = False
        if not passed:
            continue
        state.matched += 1
        number = row[0].row
        parts = []
        skip = False
        for column, o in zip(plan.group, plan.group_offsets):
            value = row[o - left]
            if value.issue == "no_cached_value":
                leave(column, "no_cached_value", value)
                skip = True
                break
            parts.append((_group_key(value, date1904), value.v))
        if skip:
            continue
        index = None
        if bins is not None:
            value = row[bins.offset - left]
            key = _use(value, types[bins.offset], plan.coerce, locale, grouping, date1904)
            if key is None:
                leave(bins.column, _reason(value, types[bins.offset]), value)
                continue
            if value.t != types[bins.offset]:
                coerced_bins.add(value)
            edges = bins.edges
            if key < edges[0] or key > edges[-1]:
                leave(bins.column, "outside_bins", value)
                continue
            index = min(bisect.bisect_right(edges, key) - 1, len(edges) - 2)
        pivot_key = None
        if plan.pivot_offset is not None:
            value = row[plan.pivot_offset - left]
            if value.issue == "no_cached_value":
                leave(plan.pivot, "no_cached_value", value)
                continue
            pivot_key = _group_key(value, date1904)
            if pivot_key not in state.pivots:
                if len(state.pivots) >= MAX_PIVOT_VALUES:
                    raise DeriveError("too_many_columns", f"{plan.pivot} holds more than "
                                      f"{MAX_PIVOT_VALUES} values to pivot; filter the rows "
                                      "(where) or group by it instead", field="pivot",
                                      column=plan.pivot)
                state.pivots[pivot_key] = (value.v, value.t)
        state.used += 1
        for o in read:
            value = row[o - left]
            if value.hidden:
                state.hidden.add(value)
                break
        group_key = tuple(k for k, _ in parts)
        key = (group_key, index)
        group = state.groups.get(key)
        if group is None:
            if len(state.groups) >= MAX_GROUPS:
                raise DeriveError("too_many_groups", f"more than {MAX_GROUPS} groups; filter "
                                  "the rows (where) or group by fewer columns",
                                  field="group_by")
            group = state.groups[key] = _Group([v for _, v in parts])
            if group_key not in state.met:
                state.met.add(group_key)
                state.order.append(group_key)
        runs(group.runs.add(number))
        group.count += 1
        cell = group.cells.get(pivot_key)
        if cell is None:
            cell = group.cells[pivot_key] = _Cell(columns)
        runs(cell.runs.add(number))
        cell.count += 1
        for o, acc in cell.accs.items():
            value = row[o - left]
            t = types[o]
            used = _use(value, t, plan.coerce, locale, grouping, date1904)
            if used is None:
                reason = _reason(value, t)
                acc.excluded[reason] = acc.excluded.get(reason, 0) + 1
                state.excluded.setdefault(o, {}).setdefault(reason, _Counter()).add(value)
                continue
            if value.t != t:
                state.coerced.setdefault(o, _Counter()).add(value)
            runs(acc.runs.add(number))
            acc.count += 1
            if acc.partials is not None:
                _add_partial(acc.partials, float(used))
            if acc.extremes:
                if acc.low is None or used < acc.low[0]:
                    acc.low = (used, value.v)
                if acc.high is None or used > acc.high[0]:
                    acc.high = (used, value.v)
            if t == "time" and not (isinstance(value.v, str) and value.v.startswith("PT")):
                acc.elapsed = False
            if acc.kept is not None:
                acc.kept.append(float(used))
                state.kept += 1
            if acc.seen is not None:
                v = used if t in _ORDERED else value.v
                if isinstance(v, str) and len(v) > _KEY_TEXT:
                    v = hashlib.blake2b(v.encode("utf-8", "surrogatepass"),
                                        digest_size=16).digest()
                before = len(acc.seen)
                acc.seen.add(v)
                state.kept += len(acc.seen) - before
            if state.kept > MAX_KEPT:
                raise DeriveError("too_many_values", f"medians and distinct counts would keep "
                                  f"more than {MAX_KEPT} values; filter the rows (where)",
                                  field="measures")
            if o in shares:
                total = state.totals[o]
                _add_partial(total[0], float(used))
                runs(total[1].add(number))
                total[2] += 1
    if coerced_bins.count:
        state.coerced["bins"] = coerced_bins
    return state


def _use(value: Value, t: str, coerce: bool, locale: Any, grouping: bool,
         date1904: bool) -> Any:
    """What a cell counts as in a column of type ``t``: a number (a number or percent
    column), a key (a date, time or datetime column: :func:`xlsx_agent.query._key`'s), the
    value itself (any other type); ``None`` when it is left out."""
    if value.t == t:
        if t in _NUMBERS:
            return float(value.v) if not isinstance(value.v, float) else value.v
        if t in _MOMENTS:
            return _key(value, t, date1904)
        return value.v
    if coerce and t == "number" and value.issue == "number_as_text":
        return text_number(value.v, locale, grouping=grouping)
    return None


# ------------------------------------------------------------------------------------------
# The result
# ------------------------------------------------------------------------------------------


def _exact(partials: list[float]) -> Fraction:
    return sum((Fraction(p) for p in partials), Fraction(0))


def _value(m: _Measure, cell: _Cell | None, t: str, plan: _Plan, state: _State
           ) -> tuple[Any, Any]:
    """A measure's value in one cell, and its sort key (the number it stands for)."""
    fn = m.fn
    if m.offset is None:
        n = cell.count if cell is not None else 0
        if fn == "count":
            return n, n
        if not n or not state.used:
            return None, None
        share = float(Fraction(n, state.used))
        return share, share
    acc = cell.accs[m.offset] if cell is not None else None
    if fn == "count":
        n = acc.count if acc is not None else 0
        return n, n
    if fn == "distinct":
        n = len(acc.seen) if acc is not None else 0
        return n, n
    if acc is None or not acc.count:
        return None, None
    if fn == "sum":
        value = math.fsum(acc.partials)
        return value, value
    if fn == "share":
        total = _exact(state.totals[m.offset][0])
        if not total:
            return None, None
        value = float(_exact(acc.partials) / total)
        return value, value
    if fn in ("min", "max"):
        key, v = acc.low if fn == "min" else acc.high
        return v, key
    if fn == "mean":
        exact = _exact(acc.partials) / acc.count
    else:  # median
        ordered = sorted(acc.kept)
        n = len(ordered)
        if n % 2:
            exact = Fraction(ordered[n // 2])
        else:
            exact = (Fraction(ordered[n // 2 - 1]) + Fraction(ordered[n // 2])) / 2
    if t in _NUMBERS:
        value = float(exact)
        return value, value
    shown = serial_text(_serial(exact, t), t, plan.date1904, t == "time" and acc.elapsed)
    return shown, float(exact)


def _output_type(m: _Measure, plan: _Plan) -> str:
    if m.fn in ("count", "distinct"):
        return "number"
    if m.fn == "share":
        return "percent"
    return plan.offset_types[m.offset]


def _label(v: Any, t: str) -> str:
    if t == "blank":
        return "(blank)"
    if t == "bool":
        return "true" if v else "false"
    if t in _NUMBERS:
        return number_text(v)
    return str(v)


def _finish(plan: _Plan, state: _State, sort: Any, top: Any, round_: Any) -> Derived:
    region = plan.region
    via = plan.via
    sheet = region.sheet_ref if via is None else None
    first_column = region.area.first.column if via is None else 0
    bins = plan.bins
    keys = list(plan.group) + ([bins.name] if bins else [])
    # the output columns: (name, measure, pivot key)
    measure_columns: list[tuple[str, _Measure, Any]] = []
    pivot_columns: dict[str, dict[str, Any]] = {}
    taken = {k.casefold() for k in keys}
    if plan.pivot is None:
        measure_columns = [(m.name, m, None) for m in plan.measures]
    else:
        for pkey, (v, t) in state.pivots.items():
            for m in plan.measures:
                base = _label(v, t)
                if len(plan.measures) > 1:
                    base = f"{base} {m.name}"
                name = base
                if name.casefold() in taken:
                    name = f"{base} ({t})"
                n = 2
                while name.casefold() in taken:
                    name = f"{base} ({t} {n})"
                    n += 1
                taken.add(name.casefold())
                measure_columns.append((name, m, pkey))
                pivot_columns[name] = {"measure": m.name, "v": v, "t": t}
    columns = keys + [name for name, _, _ in measure_columns]
    types = {g: plan.types[g] for g in plan.group}
    if bins:
        types[bins.name] = "text"
    for name, m, _ in measure_columns:
        types[name] = _output_type(m, plan)
    by_measure = _digits(round_, [m.name for m in plan.measures if m.fn not in
                                  ("count", "distinct") and _output_type(m, plan) in _NUMBERS])
    digits = {name: by_measure[m.name] for name, m, _ in measure_columns
              if m.name in by_measure}

    # the result rows, in the order met (bins: every bin of every group)
    if bins:
        slots = [(g, i) for g in state.order for i in range(len(bins.labels))]
        if not plan.group and not slots:
            slots = [((), i) for i in range(len(bins.labels))]
    elif plan.group:
        slots = list(state.groups)
    else:
        slots = [((), None)]
    if len(slots) > MAX_GROUPS:
        raise DeriveError("too_many_groups", f"more than {MAX_GROUPS} groups with their bins; "
                          "give fewer edges or filter the rows (where)", field="bins")
    shown_of = {}
    for (g, _), group in state.groups.items():
        shown_of.setdefault(g, group.shown)

    built = []
    for g, index in slots:
        group = state.groups.get((g, index))
        row: dict[str, Any] = {}
        sort_keys: dict[str, Any] = {}
        where: dict[str, Provenance] = {}
        shown = shown_of.get(g, [])
        for k, (column, offset) in enumerate(zip(plan.group, plan.group_offsets)):
            row[column] = shown[k]
            t, key = g[k]
            sort_keys[column] = None if t == "blank" else (TYPES.index(t), _text_key(key, t))
            where[column] = _row_sources(group, sheet, first_column + offset, via)
        if bins:
            row[bins.name] = bins.labels[index]
            sort_keys[bins.name] = index
            where[bins.name] = _row_sources(group, sheet, first_column + bins.offset, via)
        for name, m, pkey in measure_columns:
            cell = group.cells.get(pkey) if group is not None else None
            t = plan.offset_types[m.offset] if m.offset is not None else None
            value, key = _value(m, cell, t, plan, state)
            row[name] = value
            sort_keys[name] = key
            where[name] = _cell_sources(m, cell, sheet, first_column, plan, via)
        built.append((row, sort_keys, where))

    groups = len(built)
    built = _sorted(built, sort, columns, keys + [m.name for m in plan.measures]
                    if plan.pivot else columns)
    if top is not None:
        if isinstance(top, bool) or not isinstance(top, int) or top < 1:
            raise DeriveError("bad_top", "top is a whole number of rows, 1 or more",
                              field="top")
        built = built[:top]
    if len(built) > MAX_ROWS:
        raise DeriveError("too_many_rows", f"{len(built)} rows, more than {MAX_ROWS}; give "
                          "top, filter the rows (where) or group by fewer columns",
                          field="top")
    if len(built) * len(columns) > MAX_CELLS:
        raise DeriveError("too_many_cells", f"{len(built) * len(columns)} values, more than "
                          f"{MAX_CELLS}; give top or fewer measures or pivot values",
                          field="top")

    unrounded = [row for row, _, _ in built]
    rows = [{c: (_round(v, digits[c]) if c in digits and isinstance(v, float) else v)
             for c, v in row.items()} for row in unrounded] if digits else \
        [dict(row) for row in unrounded]
    names_of = {}
    for m in plan.measures:
        if m.offset is not None:
            names_of.setdefault(m.offset, []).append(m.name)
    excluded = {}
    coerced = {}
    for m in plan.measures:
        if m.offset is None:
            continue
        found = state.excluded.get(m.offset)
        if found:
            excluded[m.name] = {r: found[r].done() for r in REASONS if r in found}
        if m.offset in state.coerced:
            coerced[m.name] = state.coerced[m.offset].done()
    if "bins" in state.coerced:
        coerced[bins.name] = state.coerced["bins"].done()
    totals = {}
    for m in plan.measures:
        if m.fn != "share":
            continue
        if m.offset is None:
            n = state.used
            runs = _Runs()
            for group in state.groups.values():
                runs = _merge(runs, group.runs)
            totals[m.name] = Total(n, sheet, first_column + plan.row_offset, runs.bounds, n,
                                   via)
        else:
            partials, runs, n = state.totals[m.offset]
            totals[m.name] = Total(math.fsum(partials), sheet, first_column + m.offset,
                                   runs.bounds, n, via)
    left_out = {c: {r: counters[r].done() for r in REASONS if r in counters}
                for c, counters in state.left_out.items()}
    return Derived(
        address=region.address.text, id=plan.id, columns=tuple(columns),
        key_columns=len(keys), types=types, rows=rows, unrounded=unrounded,
        provenance=[where for _, _, where in built],
        rounding={"rule": ROUNDING, "digits": by_measure},
        matched=state.matched, used=state.used, groups=groups, left_out=left_out,
        excluded=excluded, coerced=coerced,
        not_compared={c: counter.done() for c, counter in state.not_compared.items()},
        hidden=state.hidden.done(), totals=totals, pivot_columns=pivot_columns,
        date1904=plan.date1904)


def _merge(a: _Runs, b: _Runs) -> _Runs:
    """The rows of both, as runs."""
    pairs = sorted([(a.bounds[i], a.bounds[i + 1]) for i in range(0, len(a.bounds), 2)]
                   + [(b.bounds[i], b.bounds[i + 1]) for i in range(0, len(b.bounds), 2)])
    out = _Runs()
    for lo, hi in pairs:
        if out.bounds and out.bounds[-1] >= lo - 1:
            out.bounds[-1] = max(out.bounds[-1], hi)
        else:
            out.bounds.append(lo)
            out.bounds.append(hi)
    return out


def _row_sources(group: _Group | None, sheet, column: int, via) -> Provenance:
    if group is None:
        return Provenance(sheet, column, array("l"), 0, via=via)
    return Provenance(sheet, column, group.runs.bounds, group.count, via=via)


def _cell_sources(m: _Measure, cell: _Cell | None, sheet, first_column: int,
                  plan: _Plan, via) -> Provenance:
    column = first_column + (plan.row_offset if m.offset is None else m.offset)
    if cell is None:
        return Provenance(sheet, column, array("l"), 0, via=via)
    if m.offset is None:
        return Provenance(sheet, column, cell.runs.bounds, cell.count, via=via)
    acc = cell.accs[m.offset]
    return Provenance(sheet, column, acc.runs.bounds, acc.count,
                      {r: acc.excluded[r] for r in REASONS if r in acc.excluded}, via=via)


def _text_key(key: Any, t: str) -> Any:
    if t in ("text", "error"):
        return (key.casefold(), key)
    return key


def _digits(round_: Any, numeric: list[str]) -> dict[str, int]:
    """The decimals per measure (by ``as``) ``round`` asks for, of the measures in
    ``numeric`` (those whose values are numbers)."""
    if round_ is None:
        return {}

    def check(n: Any) -> int:
        if isinstance(n, bool) or not isinstance(n, int) or abs(n) > MAX_DIGITS:
            raise DeriveError("bad_round", f"round is a whole number of decimals, "
                              f"-{MAX_DIGITS} to {MAX_DIGITS}", field="round")
        return n

    if isinstance(round_, Mapping):
        out = {}
        folded = {n.casefold(): n for n in numeric}
        for name, n in round_.items():
            if not isinstance(name, str) or name.casefold() not in folded:
                raise DeriveError("bad_round", "round names measures (by as) whose values "
                                  "are numbers: a sum, mean, median, min, max or share of "
                                  "numbers", field="round", valid_options=numeric)
            out[folded[name.casefold()]] = check(n)
        return out
    n = check(round_)
    return {name: n for name in numeric}


def _round(x: float, digits: int) -> float:
    """``x`` rounded to ``digits`` decimals, half away from zero, from the decimal
    :func:`repr` shows for it."""
    if not math.isfinite(x):
        return x
    exact = Decimal(repr(x))
    with localcontext() as context:
        context.prec = max(context.prec, exact.adjusted() + digits + 3)
        rounded = exact.quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP)
    return float(rounded) + 0.0


def _sorted(built: list, sort: Any, columns: list[str], named: list[str]) -> list:
    """``built`` sorted by ``sort``; a refusal lists ``named`` (the columns not named by a
    pivot's values, which are content) as the options."""
    if not sort:
        return built
    if isinstance(sort, (str, tuple)):
        sort = [sort]
    if isinstance(sort, Mapping) or not isinstance(sort, Iterable):
        raise DeriveError("bad_sort", "sort is a list of output columns, each a name or "
                          "(name, \"asc\" or \"desc\")", field="sort")
    folded = {c.casefold(): c for c in columns}
    keys = []
    for item in sort:
        if isinstance(item, str):
            name, order = item, "asc"
        elif isinstance(item, Mapping) and isinstance(item.get("column"), str) \
                and not set(item) - {"column", "desc"}:
            name, order = item["column"], "desc" if item.get("desc") else "asc"
        elif isinstance(item, (tuple, list)) and len(item) == 2 and isinstance(item[0], str):
            name, order = item
        else:
            raise DeriveError("bad_sort", "each sort key is an output column's name or "
                              "(name, \"asc\" or \"desc\")", field="sort")
        if order not in ("asc", "desc"):
            raise DeriveError("bad_sort", "a sort order is asc or desc", field="sort",
                              valid_options=["asc", "desc"])
        if not isinstance(name, str) or name.casefold() not in folded:
            raise DeriveError("bad_sort", "sort names output columns (a pivot's are named "
                              "by its values: see pivot_columns)", field="sort",
                              valid_options=named)
        keys.append((folded[name.casefold()], order == "desc"))
    for name, desc in reversed(keys):
        own = [entry for entry in built if entry[1][name] is not None]
        none = [entry for entry in built if entry[1][name] is None]
        own.sort(key=lambda entry: entry[1][name], reverse=desc)
        built = own + none
    return built
