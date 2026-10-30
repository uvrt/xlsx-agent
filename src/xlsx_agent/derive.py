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
    ds.sources["Hardware"]["average"]   # [t:Products/{SKU-0003,SKU-0007,...}/Price]: by key
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

**Joins** (:mod:`xlsx_agent.join`).  ``join={"from": ..., "on": [left, right], "columns":
[...]}`` adds another table's or dataset's columns to each row first, by key; the rest of
the pipeline then reads the joined rows, in this order: ``where``, ``group_by`` (with
``bins`` and ``pivot``), the measures, ``sort``, ``top``, ``round``.  Without
``group_by``, ``measures``, ``bins`` or ``pivot`` the result is the joined rows
themselves.

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
min, max    as mean                              the least, the greatest value (of
                                                 dates, a date)
first, last any                                  the value in the first, the last of the
                                                 rows used, in the table's row order
share       number, percent; no column: the rows this group's sum (row count) as a
                                                 fraction of the total over every row
                                                 used (:attr:`Derived.totals`)
=========== ==================================== =======================================

A measure on a column whose type is not listed is refused before anything is computed
(:class:`DeriveError`, ``wrong_type``, naming the column, its type and the functions that
take it): text has no ``min``/``max`` (Excel's collation is not measured: ``first`` and
``last`` take it), dates no ``sum``.  A measure takes the cells of its column's type only;
nothing is coerced silently (ROADMAP.md, "Design rules", 4).  Every other cell of the
group's rows is left out of that measure and counted by reason: ``blank``,
``no_cached_value``, ``error``, ``number_as_text`` (text that reads as a number in the
workbook's locale), ``date_as_text`` (text that reads as a date, in a date, time or
datetime column), ``wrong_type`` (any other).  ``coerce=["number_as_text"]`` opts in to
reading such text in a ``number`` column as its number
(:func:`~xlsx_agent.read.types.text_number`), ``coerce=["date_as_text"]`` such text in a
``date`` or ``datetime`` column as its date (:func:`~xlsx_agent.read.types.text_date`; a
two-digit year is not read: its century is Excel's rule, not measured here); the result
says so, with the cells (:attr:`Derived.coerced`), and a ``min``, ``max``, ``first`` or
``last`` that is such a cell is given as read (a number; a date as ISO 8601).  With a
coercion, the column's type is decided with those cells counted as what they read as, so
that a column of dates all stored as text is a date column (a filter on it compares dates,
its text cells counted in ``not_compared``), and a ``group_by`` column of that type groups
such text by what it reads as (a date by its day, shown as ISO 8601; counted in
``coerced`` under the column's name), so the derived column is a date column through and
through; a ``pivot`` column of that type alike (round 27, X6).  A percent column's text
stays out: its stored values are fractions, and text such as ``25`` is not one.

**Lists and cleaned text** (:mod:`xlsx_agent.clean`).  A declared column whose text is
split is a ``list`` column: ``count`` counts its cells that hold a list (rows, not items:
an item count would be a new function, and the tool's definitions stay as they are);
every other function, ``group_by`` and ``pivot`` refuse it (``wrong_type``), since a cell
holding several values belongs to no one group; ``where`` filters it with ``eq``, ``in``
and ``contains``, a row passing when any item does.  A list in a column of another type is
left out of a measure, or of the groups, as ``wrong_type``.  The cells a declared rule
changed, in every column read and every row, are counted per column with their first
addresses (:attr:`Derived.cleaned`, ``cleaned`` in the tool's result).

**Why a date column is refused.**  A column whose dates are stored as text is a ``text``
column: ``min`` and ``max`` refuse it.  The refusal says so: how many of its values are
dates stored as text (``date_as_text``), the first of them, and that
``coerce=["date_as_text"]`` reads them as dates (in a tool session, the error's
``details`` give the ``issue`` and its ``cells``); the same for numbers stored as text.
Text that is a date only in the locale's day order (``05/01/2026``) is read as one only
when the workbook is opened with ``local_dates``: the refusal names it
``local_date_text``, and no coercion reads it.  ``first`` and ``last`` take a text column
as it is.  A ``coerce`` that read none of the refused column's values says so
(``coerce ["date_as_text"] did not apply: none of its values reads as a date``).  A
dataset derived without ``coerce`` keeps the reading: the ``first`` or ``last`` of dates
stored as text is a ``text`` column whose ISO dates are ``date_as_text``
(:func:`xlsx_agent.dataset._typed`), so a derive from it can opt in.

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
cells in the group's rows that were used -- how many, and how many cells of the group's
rows were left out, by reason.

**Sources by key.**  When the table has a key column -- the one an address by key uses:
declared (``key=``), or the one column that is unique, never blank, text or whole numbers
(:meth:`~xlsx_agent.read.regions.Regions.key_column`) -- and every data row has a key no
other row shares, a value's sources are written by key: one key set per value,
``t:Products/{SKU-0003,SKU-0007,SKU-0012}/Price``, in the table's row order (one key:
``t:Products/SKU-0003/Price``).  An address by key names the same cell after rows are
inserted, deleted or sorted, where a range then names others; so a plan, a check or a
dataset made now still points at the cells it was made from (:attr:`Derived.source_key`
names the key column).  Without such a key column -- none, several candidates and none
declared, or a declared one with a blank, a date or a shared key -- the sources are
compact ranges, a run of consecutive rows being one range (``Stock!F3:F14``), true until
the sheet's rows change.  The cells left out (``excluded``, ``left_out``) are named by
their cell addresses either way.  A group's key value has the key column's cells of its rows; a
``count`` or ``share`` of rows has the rows' cells in the first group column (else the
bins' column, the pivot's, or the table's first).  So a check can read each value's
sources and recompute it, and ``count + left out`` is the group's rows: nothing is dropped
silently.  ``min``, ``max``, ``first`` and ``last`` list every cell compared, not only the
one picked (``first``'s is the first cell of its sources, ``last``'s the last).

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

**Listed pivot values** (round 27, X6): ``pivot_values=["Open", "Closed", "Held"]`` gives
one column (per measure) for each listed value, in that order, whether or not any row
holds it -- a report's fixed columns.  A cell without rows has the value a measure has of
no values: ``count`` and ``distinct`` 0, every other ``None`` (blank), since a sum or a
mean of nothing is not 0.  The values are read as a filter's are written
(:mod:`xlsx_agent.query`), as values of the pivot column's type, and matched as
``group_by`` groups (by type and value): in a number column ``"5"`` and ``5`` are both the
number 5, matching the cells holding it (``5`` and ``5.0``) and never the text ``"5"``; in
a text column ``"5"`` is that text, exactly, case counting; a date is its day
(``"2026-01-05"``); ``"true"``/``"false"`` a bool.  A value that is not one of the
column's type is refused (``wrong_type``), a value listed twice (``"5"`` and ``"5.0"`` in a
number column) ``bad_pivot``.  A row whose value is not listed is in no column: it is left
out, counted in :attr:`Derived.left_out` under the pivot column (``not_listed``, with its
first cells), and each such value is named in :attr:`Derived.pivot_unlisted` (``name``, as
its column would be named, ``v``, ``t`` and its ``rows``, in the order first met; at most
:data:`MAX_PIVOT_VALUES` of them, ``left_out`` counting every row).  There is no option to
add them as columns: list them, or pivot without ``pivot_values``.  The columns are fixed by
the list and the rows by ``group_by``: a group whose rows all hold unlisted values is still
a row, every cell of it empty (``count`` 0, the others blank), its key's sources none (the
rows used), its rows counted as above.  (Bins differ: a row outside them has no bin.)  A
blank cell is not a value that can be listed (it is left out, ``(blank)``), and neither is
a cell of another type than the column's (the text ``"5"`` in a number column: named with
its ``t``).  With ``coerce``, the pivot's text is read as ``group_by``'s is, so a
listed date matches a date stored as text too.

**Sort and top.**  ``sort`` is a list of output columns, each ``"name"`` (ascending) or
``("name", "asc" | "desc")``, the first deciding first; values without a value (``None``,
a blank group) come last either way; a key column of mixed types sorts by type first (in
the order of :data:`~xlsx_agent.read.types.TYPES`), text by :meth:`str.casefold` then as
written; a bin by its order; a date by its day.  Ties keep the groups' order.  The
unrounded values are sorted.  A name that is no output column -- an input column a
measure reads, most often -- is refused (``bad_sort``), listing the output columns and the
one most likely meant: that measure's, else the nearest name.  ``top`` keeps the first that many rows after sorting;
:attr:`Derived.groups` counts them all.

**Bounded.**  Two streamed passes over the columns concerned (the first counts their
types, so that every refusal comes before anything is computed; the second computes).
What is held is per group, never per row: the measures' running state (exact sums as
partials, extremes, counts), the group's rows as runs (a run of consecutive rows is two
numbers), and only for ``median`` its numbers and for ``distinct`` its keys.  So the
memory is bounded by the groups when a group's rows are consecutive; when they
interleave, the runs grow with the rows (16 bytes a run), up to :data:`MAX_RUNS`.  The
limits refuse with :class:`DeriveError`: :data:`MAX_GROUPS` groups (``too_many_groups``),
:data:`MAX_PIVOT_VALUES` pivot values (``too_many_columns``; listed, ``bad_pivot``),
:data:`MAX_ROWS` rows after ``top`` (``too_many_rows``) and :data:`MAX_CELLS` values
(``too_many_cells``),
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
from dataclasses import dataclass, field, replace
from decimal import ROUND_HALF_UP, Decimal, localcontext
from fractions import Fraction
from typing import TYPE_CHECKING, Any, Iterable, Mapping, Sequence, cast

from ooxml_common.locales import DATE_ORDERS
from ooxml_edit.sheets import Area, Cell, number_text
from ooxml_edit.tools import CACHE_EXTERNAL, CACHE_STATUSES, Dataset
from ooxml_edit.tools.schema import close_match

from .address import DatasetAddress, RangeAddress, SheetRef, TableAddress, parse_address
from .dataset import DatasetTable, as_dataset
from .join import JOIN_KEYS, Joined, JoinedTable, joined_table
from .profile import FIRST_ADDRESSES, Count, _add_partial, majority_type
from .query import QueryError, _Bad, _iso_key, _key, _operand, _tests
from .read.regions import Region, RegionError, Span
from .read.types import TYPES, Value, looks_like_date, serial_text, text_date, text_number

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Workbook

#: Every measure function.
FUNCTIONS = ("count", "sum", "mean", "median", "min", "max", "first", "last", "share",
             "distinct")

_NUMBERS = ("number", "percent")
_MOMENTS = ("date", "time", "datetime")
_ORDERED = _NUMBERS + _MOMENTS

#: The types of a cell's own value: every type but ``list`` (a declared split,
#: :mod:`xlsx_agent.clean`).
_CELL_TYPES = tuple(t for t in TYPES if t != "list")

#: The column types each function takes (a list column: ``count`` only).
TYPES_BY_FUNCTION: dict[str, tuple[str, ...]] = {
    "count": TYPES, "distinct": _CELL_TYPES, "sum": _NUMBERS, "share": _NUMBERS,
    "mean": _ORDERED, "median": _ORDERED, "min": _ORDERED, "max": _ORDERED,
    "first": _CELL_TYPES, "last": _CELL_TYPES}

#: The functions that take no column (they count rows).
ROW_FUNCTIONS = ("count", "share")

#: What a caller may opt in to coerce.
COERCIONS = ("number_as_text", "date_as_text")

#: Why a cell is left out of a measure, or a row out of the groups.
REASONS = ("blank", "no_cached_value", "error", "number_as_text", "date_as_text",
           "wrong_type", "outside_bins", "not_listed")

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
    ``bad_round``, ``bad_coerce``, ``bad_join``, ``duplicate_name``, and the limits: ``too_many_groups``,
    ``too_many_columns``, ``too_many_rows``, ``too_many_cells``, ``too_many_sources``,
    ``too_many_values``); ``field`` the argument at fault.  A :class:`QueryError`, as a
    ``where`` refusal is: the message names columns, functions and types, never a value.

    A ``wrong_type`` refusal of a text column that holds dates or numbers stored as text
    names that ``issue`` (``date_as_text``, ``number_as_text``, ``local_date_text``: text
    that is a date in the locale's day order, read as one only with ``local_dates``) and
    its ``cells``, the count with the first addresses, and says how to read them."""

    def __init__(self, code: str, message: str, *, issue: str | None = None,
                 cells: Count | None = None, **options: Any) -> None:
        super().__init__(code, message, **options)
        self.issue = issue
        self.cells = cells


class _Keyed:
    """A keyed table's rows by key, for sources written by key: ``whole`` the table's
    address, ``keys`` each data row's key by position (from the first data row),
    ``first_row`` the first data row's sheet row, ``columns`` the region's column names
    from its ``first_column`` (for a field derived from another, the column it reads: its
    cells are that column's), ``column`` the key column's name."""

    __slots__ = ("whole", "keys", "first_row", "first_column", "columns", "column")

    def __init__(self, region: Region, index: Any) -> None:
        self.whole = region.address
        self.keys = index.keys
        assert region.data is not None   # _keyed asks only of a table with data rows
        self.first_row = region.data.first.row

        self.first_column = region.area.first.column
        # A field derived from another is that column's cells: its sources name that one.
        self.columns = tuple(source or name for name, source in
                             zip(region.columns, region.derived or (None,) * len(region.columns)))
        self.column = index.column

    def address(self, column: int, runs: Sequence[int]) -> "TableAddress | None":
        """The cells of sheet column ``column`` in the rows ``runs`` (first and last sheet
        rows, flat), by key: a key set, one key, or ``None`` for no rows."""
        top, keys = self.first_row, self.keys
        found = tuple(keys[row - top] for i in range(0, len(runs), 2)
                      for row in range(runs[i], runs[i + 1] + 1))
        if not found:
            return None
        name = self.columns[column - self.first_column]
        if len(found) == 1:
            return replace(self.whole, key=found[0], column=name)
        return replace(self.whole, keys=found, column=name)


def _keyed(workbook: "Workbook", region: Region) -> "_Keyed | None":
    """The key form of ``region``'s sources: its key column's index when every data row
    has a key no other row shares; ``None`` (ranges) otherwise."""
    if region.data is None:
        return None
    try:
        index = workbook.regions.key_index(region)
    except (RegionError, KeyError):
        return None
    return _Keyed(region, index) if index.unique else None


class _Sourced:
    """Cells of one column as runs of rows, kept compact (two numbers a run) and given as
    addresses when asked for: by key, or as range addresses."""

    __slots__ = ()

    if TYPE_CHECKING:   # the fields of the dataclasses built on it
        sheet: "SheetRef | Span | None"
        column: int
        runs: array
        via: "DatasetTable | None"
        keyed: "_Keyed | None"

    @property
    def sources(self) -> "tuple[TableAddress | RangeAddress | str, ...]":
        """The cells: by key when the table has a key column (``keyed``), one key set
        (``t:Products/{SKU-0003,SKU-0007}/Price``) in row order; else as compact ranges, a
        run of consecutive rows one range (``Stock!F3:F14``), in row order.  Derived from a
        dataset (``via``), the cells its values came from (:meth:`DatasetTable.sources_of
        <xlsx_agent.dataset.DatasetTable.sources_of>`)."""
        if self.via is not None:
            return self.via.sources_of(self.column, self.runs)
        if self.keyed is not None:
            found = self.keyed.address(self.column, self.runs)
            return () if found is None else (found,)
        b, sheet, column = self.runs, self.sheet, self.column
        if isinstance(sheet, Span):    # a declared table read as one: each sheet's cells
            return tuple(r for i in range(0, len(b), 2)
                         for r in sheet.locate(Area(Cell(b[i], column), Cell(b[i + 1], column))))
        assert sheet is not None   # a table's cells (via: a dataset's, given above)
        return tuple(RangeAddress(sheet, Area(Cell(b[i], column), Cell(b[i + 1], column)))

                     for i in range(0, len(b), 2))

    @property
    def ranges(self) -> int:
        """How many addresses :attr:`sources` gives."""
        if self.via is not None or self.keyed is not None or isinstance(self.sheet, Span):
            return len(self.sources)
        return len(self.runs) // 2


@dataclass(frozen=True)
class Provenance(_Sourced):
    """Where one value came from: :attr:`sources` the cells it was computed from, by key
    or as compact ranges; ``count`` how many cells those are; ``excluded`` how many cells
    of the group's rows were left out of it, by reason.  ``sheet``, ``column`` and
    ``runs`` (first and last rows, flat) are what :attr:`sources` is made from (``sheet``
    a :class:`~xlsx_agent.read.regions.Span` for a declared table read as one: its rows
    are its own, and each run is given as each sheet's cells); ``keyed``
    the table's keys, when they name its rows; ``via`` the dataset it was derived from, if
    any (then ``column`` is the dataset's, from 0, ``runs`` its positions, and ``count``
    how many of its values were used)."""

    sheet: "SheetRef | Span | None"
    column: int
    runs: array = field(repr=False)
    count: int
    excluded: dict[str, int] = field(default_factory=dict)
    via: "DatasetTable | None" = field(default=None, repr=False, compare=False)
    keyed: "_Keyed | None" = field(default=None, repr=False, compare=False)

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
    sheet: "SheetRef | Span | None"
    column: int
    runs: array = field(repr=False)
    count: int
    via: "DatasetTable | None" = field(default=None, repr=False, compare=False)
    keyed: "_Keyed | None" = field(default=None, repr=False, compare=False)

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
    (and the bins, by their ``as``, and each ``group_by`` column, by its name), the cells
    read as numbers or dates by ``coerce``; ``cleaned``
    per column read, the cells a declared rule changed (:mod:`xlsx_agent.clean`), every
    row's, filtered out or not; ``not_compared`` per filtered column, as
    :class:`~xlsx_agent.query.QueryResult` gives it; ``hidden`` the rows used that hold a
    hidden cell among the columns read; ``totals`` per ``share`` measure, its total;
    ``pivot_columns`` per pivoted column, its measure and pivot value; ``pivot_unlisted``
    the values present but not in ``pivot_values``, each ``{name, v, t, rows}``;
    ``date1904`` the date system its dates are in (its source's); ``source_key`` the key column its
    sources are written by (``None``: as ranges; see the module's notes); ``cache`` the
    formula cells read (every row's, in every column the derive reads), counted by their
    cache status (``cached_unverified``, ``invalidated``, ``absent``) and ``external``
    (those referring to another workbook) -- a dataset's own counts when derived from one
    -- so that a result computed from invalidated caches says so, and its dataset is
    stale (ooxml-edit's ``Dataset.cache``); ``joined`` what a ``join`` did
    (:class:`~xlsx_agent.join.Joined`), ``None`` without one.

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
    cleaned: dict[str, Count] = field(default_factory=dict)
    not_compared: dict[str, Count] = field(default_factory=dict)
    hidden: Count = field(default_factory=lambda: Count(0))
    totals: dict[str, Total] = field(default_factory=dict)
    pivot_columns: dict[str, dict[str, Any]] = field(default_factory=dict)
    pivot_unlisted: list[dict[str, Any]] = field(default_factory=list)
    date1904: bool = False
    source_key: str | None = None
    cache: dict[str, int] = field(default_factory=dict)
    joined: "Joined | None" = None

    def dataset(self) -> "Dataset":
        """This result as an ooxml-edit :class:`~ooxml_edit.tools.Dataset`: the rows as
        reported (rounded), each value's sources as range text, the key columns as
        ``keys``, the types as labels, and this result as ``detail``
        (:func:`xlsx_agent.dataset.as_dataset`)."""
        return as_dataset(self)

    @property
    def sources(self) -> "dict[Any, dict[str, list[TableAddress | RangeAddress]]]":
        """Each row's sources by its key (the key column's value; a tuple of them with
        several key columns; ``None`` without any), then by column.  Two keys that show
        alike (a number and a bool, a date and its ISO text) keep the first row's: by
        position, :attr:`provenance` has every one."""
        out: dict[Any, dict[str, list]] = {}
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
        """The result as plain data: rows as lists in column order, sources as address
        text."""
        out: dict = {"from": self.address, "columns": list(self.columns),
                     "key_columns": self.key_columns, "types": dict(self.types),
                     "rows": [[row[c] for c in self.columns] for row in self.rows],
                     "rounding": dict(self.rounding),
                     "provenance": [[where[c].as_dict() for c in self.columns]
                                    for where in self.provenance],
                     "matched": self.matched, "used": self.used, "groups": self.groups}
        if self.id is not None:
            out["id"] = self.id
        if self.source_key is not None:
            out["source_key"] = self.source_key
        if self.rows != self.unrounded:
            out["unrounded"] = [[row[c] for c in self.columns] for row in self.unrounded]
        for name in ("left_out", "excluded"):
            value = getattr(self, name)
            if value:
                out[name] = {k: {r: c.as_dict() for r, c in v.items()}
                             for k, v in value.items()}
        for name in ("coerced", "cleaned", "not_compared"):
            value = getattr(self, name)
            if value:
                out[name] = {k: c.as_dict() for k, c in value.items()}
        if self.hidden.count:
            out["hidden"] = self.hidden.as_dict()
        if self.totals:
            out["totals"] = {k: t.as_dict() for k, t in self.totals.items()}
        if self.pivot_columns:
            out["pivot_columns"] = {k: dict(v) for k, v in self.pivot_columns.items()}
        if self.pivot_unlisted:
            out["pivot_unlisted"] = [dict(u) for u in self.pivot_unlisted]
        if self.cache:
            out["cache"] = dict(self.cache)
        if self.joined is not None:
            out["joined"] = self.joined.as_dict()
        return out


def derive(workbook: "Workbook | None",
           from_: "str | TableAddress | Region | Derived | Dataset | DatasetTable", *,
           where: Any = None, group_by: "Sequence[str] | str | None" = None,
           measures: "Sequence[Mapping[str, Any]] | None" = None,
           bins: "Mapping[str, Any] | None" = None, pivot: str | None = None,
           sort: Any = None, top: int | None = None,
           round: "int | Mapping[str, int] | None" = None,
           coerce: "Sequence[str] | None" = None, id: str | None = None,
           join: "Mapping[str, Any] | None" = None,
           pivot_values: "Sequence[Any] | None" = None) -> Derived:
    """Figures from a table, name or region, or from a dataset (``workbook`` may then be
    ``None``); see the module's notes.  ``join`` adds another table's or dataset's columns
    to each row first (:mod:`xlsx_agent.join`); ``pivot_values`` lists the pivot's columns,
    in order."""
    if join is not None:
        return _join(workbook, from_, join, where=where, group_by=group_by,
                     measures=measures, bins=bins, pivot=pivot, sort=sort, top=top,
                     round=round, coerce=coerce, id=id, pivot_values=pivot_values)
    region = _region(workbook, from_)
    plan = _Plan(workbook, region, where=where, group_by=group_by, measures=measures,
                 bins=bins, pivot=pivot, coerce=coerce, id=id, pivot_values=pivot_values)
    plan.check_types(*_type_pass(workbook, region, plan.typed_offsets, plan.coerce))
    state = plan.run()
    return plan.finish(state, sort=sort, top=top, round=round)


# ------------------------------------------------------------------------------------------
# Joins (round 25, X1: xlsx_agent.join)
# ------------------------------------------------------------------------------------------


def _join(workbook: "Workbook | None", from_: Any, join: Any, *, where: Any, group_by: Any,
          measures: Any, bins: Any, pivot: Any, sort: Any, top: Any, round: Any, coerce: Any,
          id: str | None, pivot_values: Any = None) -> Derived:
    """A derive of ``from_``'s rows joined with ``join["from"]``'s (see
    :mod:`xlsx_agent.join`): the pipeline on the joined rows, or the rows themselves."""
    if not isinstance(join, Mapping) or set(join) - set(JOIN_KEYS) \
            or "from" not in join or "on" not in join:
        raise DeriveError("bad_join", "join is {from: the table or dataset to join, on: "
                          "[this table's key column, the joined one's], columns: the joined "
                          "columns to add}", field="join")
    left = _region(workbook, from_)
    try:
        right = _region(workbook, join["from"])
    except DeriveError as error:
        raise DeriveError("bad_join", f"join's {error.message}", field="join") from None
    table = joined_table(workbook, left, right, join)
    if group_by is not None or measures is not None or bins is not None or pivot is not None \
            or pivot_values is not None:
        derived = derive(workbook, table, where=where, group_by=group_by, measures=measures,
                         bins=bins, pivot=pivot, sort=sort, top=top, round=round,
                         coerce=coerce, id=id, pivot_values=pivot_values)
        return replace(derived, joined=table.joined)
    if round is not None:
        raise DeriveError("bad_round", "round rounds measures; a join without group_by or "
                          "measures gives its rows' cells as stored", field="round")
    plan = _Plan(workbook, table, where=where, group_by=None, measures=None, bins=None,
                 pivot=None, coerce=coerce, id=id)
    every = list(range(len(table.columns)))
    plan.check_types(*_type_pass(workbook, table, every, plan.coerce))
    types = plan.offset_types
    typed_tests = [(t, types[t.offset]) for t in plan.tests]
    date1904 = plan.date1904
    state = _State()
    built = []
    for row in table.measured_rows():
        for value in row:
            if value.cleaned:
                state.cleaned.setdefault(value.column - 1, _Counter()).add(value)
            if value.cache is not None:
                state.cache[value.cache] = state.cache.get(value.cache, 0) + 1
                if value.external:
                    state.cache["external"] = state.cache.get("external", 0) + 1
        passed = True
        for test, t in typed_tests:
            value = row[test.offset]
            if test.compares and value.t not in ("blank", t):
                state.not_compared.setdefault(test.column, _Counter()).add(value)
                passed = False
            elif passed and not test.passes(value, date1904):
                passed = False
        if not passed:
            continue
        state.matched += 1
        if any(value.hidden for value in row):
            state.hidden.add(row[0])
        position = row[0].row
        values: dict[str, Any] = {}
        keys: dict[str, Any] = {}
        where_: dict[str, Provenance] = {}
        for offset, (name, value) in enumerate(zip(table.columns, row)):
            values[name] = value.v
            keys[name] = _row_sort_key(value, date1904)
            where_[name] = Provenance(None, offset, array("l", [position, position]),
                                      0 if value.t == "blank" else 1, via=table)
        built.append((values, keys, where_))
    columns = list(table.columns)
    groups = len(built)
    built = _sorted(built, sort, columns, columns)
    if top is not None:
        if isinstance(top, bool) or not isinstance(top, int) or top < 1:
            raise DeriveError("bad_top", "top is a whole number of rows, 1 or more",
                              field="top")
        built = built[:top]
    if len(built) > MAX_ROWS:
        raise DeriveError("too_many_rows", f"{len(built)} joined rows, more than {MAX_ROWS}; "
                          "give top or filter the rows (where)", field="top")
    if len(built) * len(columns) > MAX_CELLS:
        raise DeriveError("too_many_cells", f"{len(built) * len(columns)} values, more than "
                          f"{MAX_CELLS}; give top, filter the rows (where) or fewer joined "
                          "columns", field="top")
    rows = [values for values, _, _ in built]
    return Derived(
        address=table.address.text, id=id, columns=tuple(columns),
        key_columns=table.left_keys, types={c: types[o] for o, c in enumerate(columns)},
        rows=rows, unrounded=[dict(r) for r in rows],
        provenance=[where_ for _, _, where_ in built],
        rounding={"rule": ROUNDING, "digits": {}}, matched=state.matched,
        used=state.matched, groups=groups,
        cleaned={columns[o]: c.done() for o, c in sorted(state.cleaned.items())},
        not_compared={c: counter.done() for c, counter in state.not_compared.items()},
        hidden=state.hidden.done(), date1904=date1904,
        cache=_cache(state.cache, None), joined=table.joined)


def _row_sort_key(value: Value, date1904: bool) -> Any:
    """A joined row's value as ``sort`` orders it: by type first, then as a group's key
    sorts (a list by its items, as text)."""
    if value.t == "blank":
        return None
    if value.t == "list":
        return (TYPES.index("list"), tuple((str(i).casefold(), str(i))
                                           for i in cast(tuple, value.v)))
    t, key = _group_key(value, date1904)
    return (TYPES.index(t), _text_key(key, t))


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
        if not from_.is_whole:
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
                 where, group_by, measures, bins, pivot, coerce, id,
                 pivot_values: Any = None) -> None:
        self.workbook = workbook
        self.region = region
        self.via = region if isinstance(region, DatasetTable) else None
        if self.via is not None:
            self.date1904 = self.via.date1904
        else:
            assert workbook is not None   # only a dataset is derived from without one
            self.date1904 = workbook.date1904
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
        self.coerce = frozenset(coerce)
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
        self.listed_raw = self._pivot_values(pivot_values)
        #: pivot_values read as keys of the pivot column's type, in order: key -> (v, t)
        self.listed: dict[tuple, tuple[Any, str]] | None = None
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
                                    | ({self.pivot_offset} if self.pivot_offset is not None
                                       else set()))

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

    def _pivot_values(self, values: Any) -> list | None:
        if values is None:
            return None
        if self.pivot is None:
            raise DeriveError("bad_pivot", "pivot_values lists the values of pivot's column: "
                              "give pivot too", field="pivot_values")
        if isinstance(values, (str, bytes, Mapping)) or not isinstance(values, Iterable):
            raise DeriveError("bad_pivot", "pivot_values is a list of values, as filters are "
                              "written", field="pivot_values")
        values = list(values)
        if not 1 <= len(values) <= MAX_PIVOT_VALUES:
            raise DeriveError("bad_pivot", f"pivot_values lists 1 to {MAX_PIVOT_VALUES} values",
                              field="pivot_values")
        if not all(isinstance(v, (str, int, float)) for v in values):
            raise DeriveError("bad_pivot", "each of pivot_values is a value as filters are "
                              "written: text, a number or true/false", field="pivot_values")
        return values

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

    def check_types(self, types: dict[int, str],
                    issues: "dict[int, dict[str, Count]] | None" = None) -> None:
        """Refuse what the columns' types do not take, before anything is computed.
        ``issues``: per column, its cells stored as text that read as another type, by
        issue, for a refusal to name (:func:`_type_pass`)."""
        region = self.region
        self.types = {region.columns[o]: t for o, t in types.items()}
        self.offset_types = types
        for test in self.tests:
            test.check(self.types[test.column], self.date1904)
        grouped = [(c, o, "group_by") for c, o in zip(self.group, self.group_offsets)]
        if self.pivot is not None and self.pivot_offset is not None:
            grouped.append((self.pivot, self.pivot_offset, "pivot"))

        for column, o, field_name in grouped:
            if types[o] == "list":
                raise DeriveError("wrong_type", f"{column} holds lists (its declaration "
                                  "splits its text) and is not grouped by: count it, or "
                                  "filter it with eq, in or contains", field=field_name,
                                  column=column, column_type="list",
                                  valid_options=[n for n, t in self.types.items()
                                                 if t not in ("list", "blank")][:50])
        if self.listed_raw is not None and self.pivot_offset is not None:
            self.listed = _listed(self.listed_raw, self.pivot or "",
                                  types[self.pivot_offset], self.date1904)
        for m in self.measures:
            if m.offset is None:
                continue
            t = types[m.offset]
            allowed = TYPES_BY_FUNCTION[m.fn]
            if t not in allowed:
                takes = [fn for fn in FUNCTIONS if t in TYPES_BY_FUNCTION[fn]]
                if t == "blank":
                    reason = "holds no values"
                elif t == "list":
                    reason = ("holds lists (its declaration splits its text): count counts "
                              "its cells; filter it with eq, in or contains")
                elif t == "text" and m.fn in ("min", "max"):
                    reason = "has no measured order (text collation is not measured)"
                else:
                    reason = f"is not what {m.fn} takes ({', '.join(allowed)})"
                hint, issue, cells = self._stored_as_text(m.fn, t, (issues or {}).get(m.offset),
                                                          m.column)
                raise DeriveError("wrong_type", f"{m.column} is a {t} column and {reason}; "
                                  f"it takes {', '.join(takes)}{hint}", field="measures",
                                  column=m.column, column_type=t, valid_options=takes,
                                  issue=issue, cells=cells)
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

    def _stored_as_text(self, fn: str, t: str, found: "dict[str, Count] | None",
                        column: str | None = None) -> tuple[str, str | None, Count | None]:
        """What a refusal of ``fn`` on a ``t`` column adds when the column's text holds
        dates or numbers: the words, the issue and its cells; nothing otherwise.  A
        dataset's ISO dates may be labelled instead, and the words say how."""
        if t != "text":
            return "", None, None
        if not found:
            found = {}
        wants = {"date_as_text": _MOMENTS, "local_date_text": _MOMENTS,
                 "number_as_text": ("number",)}
        for issue in ("date_as_text", "local_date_text", "number_as_text"):
            cells = found.get(issue)
            if cells is None or not set(TYPES_BY_FUNCTION[fn]) & set(wants[issue]):
                continue
            first = cells.first[0]
            what = "numbers" if issue == "number_as_text" else "dates"
            words = f". {cells.count} of its values are {what} stored as text ({issue}, " \
                    f"first {first})"
            if issue == "local_date_text":
                words += "; they are read as dates only when the workbook is opened with " \
                         "local_dates"
            elif issue in self.coerce:
                words += f", too few to make it a {what[:-1]} column"
            else:
                words += f": coerce [\"{issue}\"] reads them as {what}"
            if self.via is not None and not self.via.derived and issue == "date_as_text" \
                    and column is not None and issue not in self.coerce:
                words += f", as labelling the dataset's column does (Dataset.from_records(" \
                         f"..., types={{{column!r}: \"date\"}}))"
            return words, issue, cells
        for issue, what in (("date_as_text", "a date"), ("number_as_text", "a number")):
            if issue in self.coerce and set(TYPES_BY_FUNCTION[fn]) & set(wants[issue]):
                return (f". coerce [\"{issue}\"] did not apply: none of its values reads as "
                        f"{what}"), None, None
        return "", None, None

    def run(self) -> "_State":
        return _second_pass(self)

    def finish(self, state: "_State", *, sort: Any, top: Any, round: Any) -> Derived:
        return _finish(self, state, sort, top, round)


def _listed(raw: list, column: str, t: str, date1904: bool) -> dict[tuple, tuple[Any, str]]:
    """``pivot_values`` as group keys of a ``t`` column (:func:`_group_key`'s), in order,
    each with the value a cell of it shows: read as a filter's value is
    (:func:`xlsx_agent.query._operand`), so ``"5"`` and ``5`` are the number 5 in a number
    column and the text ``"5"`` in a text column."""
    out: dict[tuple, tuple[Any, str]] = {}
    for item in raw:
        try:
            key = _operand(item, t, date1904)
        except _Bad:
            raise DeriveError("wrong_type", f"{column} is a {t} column: each of pivot_values "
                              "is a value of its type, as filters are written",
                              field="pivot_values", column=column, column_type=t) from None
        if t in _NUMBERS:
            key = float(key)
            shown: Any = int(key) if key.is_integer() and abs(key) < 2 ** 53 else key
        elif t in _MOMENTS:
            shown = _edge_text(key, t, date1904)
        else:
            shown = key
        if (t, key) in out:
            raise DeriveError("bad_pivot", "pivot_values lists each value once (as its "
                              f"{t} column reads it)", field="pivot_values", column=column)
        out[(t, key)] = (shown, t)
    return out


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
          ) -> tuple[Iterable[Sequence[Value]], int] | None:
    """The rows to read for the columns at ``offsets``, streamed, and the offset of each
    row's first value; ``None`` when there are no data rows.  A dataset gives every
    column, measured (unrounded)."""
    if isinstance(region, DatasetTable):
        return region.measured_rows(), 0
    area = _area(region, offsets)
    if area is None:
        return None
    left = area.first.column - region.area.first.column
    first = region.area.first.column
    assert workbook is not None   # only a dataset is derived from without one
    return workbook._read(
RangeAddress(region.sheet_ref, area),
                          columns={first + o for o in offsets or [0]}), left


def _type_pass(workbook: "Workbook | None", region: "Region | DatasetTable",
               offsets: list[int], coerce: "frozenset[str]" = frozenset()
               ) -> tuple[dict[int, str], dict[int, dict[str, Count]]]:
    """Each column's type, from one streamed pass counting its values' types; and per
    column, its text cells that read as another type, by issue (``number_as_text``,
    ``date_as_text``; ``local_date_text`` for a date in the locale's day order that the
    workbook was opened not to read, without ``local_dates``).  A cell a ``coerce`` opts
    in to counts as the type it reads as (a date as ``date``, or ``datetime`` with a time
    of day)."""
    found = _rows(workbook, region, offsets)
    counts: dict[int, dict[str, int]] = {o: {} for o in offsets}
    issues: dict[int, dict[str, _Counter]] = {o: {} for o in offsets}
    locale = workbook.locale if workbook is not None else None
    local = workbook is not None and not workbook.local_dates \
        and getattr(locale, "date_order", None) in DATE_ORDERS
    if found is not None and offsets:
        rows, left = found
        for row in rows:
            for o in offsets:
                value = row[o - left]
                t = value.t
                if t == "blank":
                    continue
                issue = value.issue
                if t == "text":
                    if issue is None and local and looks_like_date(cast(str, value.v), locale,
                                                                   local_dates=True):
                        issue = "local_date_text"
                    if issue in ("number_as_text", "date_as_text", "local_date_text"):
                        issues[o].setdefault(issue, _Counter()).add(value)
                    if issue in coerce:
                        t = _coerced_type(value, issue, locale)
                tally = counts[o]
                tally[t] = tally.get(t, 0) + 1
    return ({o: majority_type(c) for o, c in counts.items()},
            {o: {k: c.done() for k, c in found.items()} for o, found in issues.items()})


def _coerced_type(value: Value, issue: str, locale: Any) -> str:
    """The type a text cell with ``issue`` reads as, coerced; ``text`` when it does not."""
    if issue == "number_as_text":
        return "number"
    shown = text_date(cast(str, value.v), locale, local_dates=True)
    if shown is None:
        return "text"
    return "datetime" if "T" in shown else "date"


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

    __slots__ = ("runs", "count", "partials", "extremes", "low", "high", "ends", "first",
                 "last", "kept", "seen", "excluded", "elapsed")

    def __init__(self, fns: set[str]) -> None:
        self.runs = _Runs()
        self.count = 0
        self.partials: list[float] | None = [] if fns & {"sum", "mean", "share"} else None
        self.extremes = bool(fns & {"min", "max"})
        self.low: tuple | None = None
        self.high: tuple | None = None
        self.ends = bool(fns & {"first", "last"})
        self.first: tuple | None = None
        self.last: tuple | None = None
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
    unlisted: dict = field(default_factory=dict)         # pivot key -> [v, t, rows]
    matched: int = 0
    used: int = 0
    runs: int = 0
    kept: int = 0
    left_out: dict = field(default_factory=dict)         # column -> reason -> _Counter
    excluded: dict = field(default_factory=dict)         # offset -> reason -> _Counter
    coerced: dict = field(default_factory=dict)          # offset -> _Counter
    coerced_groups: dict = field(default_factory=dict)   # group offset -> _Counter
    cleaned: dict = field(default_factory=dict)          # offset -> _Counter
    not_compared: dict = field(default_factory=dict)     # column -> _Counter
    hidden: _Counter = field(default_factory=_Counter)
    totals: dict = field(default_factory=dict)           # offset or None -> [partials, runs, count]
    cache: dict = field(default_factory=dict)            # status or "external" -> count


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
    if value.issue == "date_as_text" and column_type in _MOMENTS:
        return "date_as_text"
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

    cache = state.cache
    for row in rows:
        for o in read:
            value = row[o - left]
            if value.cleaned:
                state.cleaned.setdefault(o, _Counter()).add(value)
            if value.cache is not None:
                cache[value.cache] = cache.get(value.cache, 0) + 1
                if value.external:
                    cache["external"] = cache.get("external", 0) + 1
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
            if value.issue == "no_cached_value" or value.t == "list":
                leave(column, "no_cached_value" if value.t != "list" else "wrong_type", value)
                skip = True
                break
            t = types[o]
            if value.t != t and value.issue in plan.coerce:
                # a coerced column's text is grouped as what it reads as: its date's day
                used = _use(value, t, plan.coerce, locale, grouping, date1904)
                if used is not None:
                    state.coerced_groups.setdefault(o, _Counter()).add(value)
                    parts.append(((t, used), _shown(value, used, t, date1904)))
                    continue
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
        unlisted = False        # a row of a value not in pivot_values: its group, no cell
        if plan.pivot is not None and plan.pivot_offset is not None:
            value = row[plan.pivot_offset - left]
            if value.issue == "no_cached_value" or value.t == "list":
                leave(plan.pivot, "no_cached_value" if value.t != "list" else "wrong_type",
                      value)
                continue
            t = types[plan.pivot_offset]
            pivot_key = None
            shown = value.v
            if value.t != t and value.issue in plan.coerce:
                # as a group_by column's: a coerced column's text is what it reads as
                used = _use(value, t, plan.coerce, locale, grouping, date1904)
                if used is not None:
                    state.coerced_groups.setdefault(plan.pivot_offset, _Counter()).add(value)
                    pivot_key, shown = (t, used), _shown(value, used, t, date1904)
            if pivot_key is None:
                pivot_key = _group_key(value, date1904)
            if plan.listed is not None:
                if pivot_key not in plan.listed:
                    leave(plan.pivot, "not_listed", value)
                    found = state.unlisted.get(pivot_key)
                    if found is not None:
                        found[2] += 1
                    elif len(state.unlisted) < MAX_PIVOT_VALUES:
                        state.unlisted[pivot_key] = [shown, pivot_key[0], 1]
                    unlisted = True
            elif pivot_key not in state.pivots:
                if len(state.pivots) >= MAX_PIVOT_VALUES:
                    raise DeriveError("too_many_columns", f"{plan.pivot} holds more than "
                                      f"{MAX_PIVOT_VALUES} values to pivot; filter the rows "
                                      "(where) or group by it instead", field="pivot",
                                      column=plan.pivot)
                state.pivots[pivot_key] = (shown, pivot_key[0])
        if not unlisted:
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
        if unlisted:
            # the listed columns are fixed, and so are the groups: a group of unlisted
            # values only is a row of empty cells (its rows left out, counted)
            continue
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
            if acc.extremes or acc.ends:
                shown = _shown(value, used, t, date1904)
                if acc.extremes:
                    if acc.low is None or used < acc.low[0]:
                        acc.low = (used, shown)
                    if acc.high is None or used > acc.high[0]:
                        acc.high = (used, shown)
                if acc.ends:
                    if acc.first is None:
                        acc.first = (_text_key(used, t), shown)
                    acc.last = (_text_key(used, t), shown)
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


def _use(value: Value, t: str, coerce: "frozenset[str]", locale: Any, grouping: bool,
         date1904: bool) -> Any:
    """What a cell counts as in a column of type ``t``: a number (a number or percent
    column), a key (a date, time or datetime column: :func:`xlsx_agent.query._key`'s), the
    value itself (any other type); ``None`` when it is left out."""
    if value.t == t:
        if t in _NUMBERS:
            return float(cast(float, value.v)) if not isinstance(value.v, float) else value.v
        if t in _MOMENTS:
            return _key(value, t, date1904)
        return value.v
    if value.issue not in coerce:
        return None
    if t == "number" and value.issue == "number_as_text":
        return text_number(cast(str, value.v), locale, grouping=grouping)
    if t in ("date", "datetime") and value.issue == "date_as_text":
        shown = text_date(cast(str, value.v), locale, local_dates=True)
        return _iso_key(shown, t, date1904) if shown is not None else None
    return None


def _shown(value: Value, used: Any, t: str, date1904: bool) -> Any:
    """A cell's value as a measure gives it: its own, or a coerced one's as read (a number;
    a date as ISO 8601)."""
    if value.t == t:
        return value.v
    if t in _MOMENTS:
        return serial_text(_serial(Fraction(used), t), t, date1904)
    return used


# ------------------------------------------------------------------------------------------
# The result
# ------------------------------------------------------------------------------------------


def _exact(partials: list[float]) -> Fraction:
    return sum((Fraction(p) for p in partials), Fraction(0))


def _value(m: _Measure, cell: _Cell | None, t: str | None, plan: _Plan, state: _State
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
    # Each accumulator keeps what its column's measures need (_Acc): what fn reads is there.
    if fn == "distinct":
        n = len(acc.seen) if acc is not None and acc.seen is not None else 0
        return n, n
    if acc is None or not acc.count:
        return None, None
    if fn == "sum":
        assert acc.partials is not None
        value = math.fsum(acc.partials)
        return value, value
    if fn == "share":
        assert acc.partials is not None
        total = _exact(state.totals[m.offset][0])
        if not total:
            return None, None
        value = float(_exact(acc.partials) / total)
        return value, value
    if fn in ("min", "max", "first", "last"):
        ends = {"min": acc.low, "max": acc.high, "first": acc.first, "last": acc.last}[fn]
        assert ends is not None   # set with the column's first value
        key, v = ends
        return v, key
    assert t is not None   # a measure of a column
    if fn == "mean":
        assert acc.partials is not None
        exact = _exact(acc.partials) / acc.count
    else:  # median
        assert acc.kept is not None
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
    assert m.offset is not None   # every other measure is of a column
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
    sheet: "SheetRef | Span | None" = None
    first_column = 0
    keyed = None
    if not isinstance(region, DatasetTable):
        sheet = region.span or region.sheet_ref
        first_column = region.area.first.column
        if state.used:
            assert plan.workbook is not None   # a table is derived from with its workbook
            keyed = _keyed(plan.workbook, region)
    bins = plan.bins
    keys = list(plan.group) + ([bins.name] if bins else [])
    # the output columns: (name, measure, pivot key)
    measure_columns: list[tuple[str, _Measure, Any]] = []
    pivot_columns: dict[str, dict[str, Any]] = {}
    taken = {k.casefold() for k in keys}
    if plan.pivot is None:
        measure_columns = [(m.name, m, None) for m in plan.measures]
    else:
        pivots = plan.listed if plan.listed is not None else state.pivots
        for pkey, (v, t) in pivots.items():
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
    slots: list[tuple[Any, Any]]
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
    shown_of: dict[Any, Any] = {}
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
            where[column] = _row_sources(group, sheet, first_column + offset, via, keyed)
        if bins:
            row[bins.name] = bins.labels[index]
            sort_keys[bins.name] = index
            where[bins.name] = _row_sources(group, sheet, first_column + bins.offset, via,
                                            keyed)
        for name, m, pkey in measure_columns:
            cell = group.cells.get(pkey) if group is not None else None
            t = plan.offset_types[m.offset] if m.offset is not None else None
            value, key = _value(m, cell, t, plan, state)
            row[name] = value
            sort_keys[name] = key
            where[name] = _cell_sources(m, cell, sheet, first_column, plan, via, keyed)
        built.append((row, sort_keys, where))

    groups = len(built)
    built = _sorted(built, sort, columns, keys + [m.name for m in plan.measures]
                    if plan.pivot else columns,
                    {} if plan.pivot else {m.column.casefold(): m.name
                                           for m in reversed(plan.measures) if m.column},
                    pivot=plan.pivot is not None)
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
    names_of: dict[int, list[str]] = {}
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
    if bins and "bins" in state.coerced:
        coerced[bins.name] = state.coerced["bins"].done()
    for column, o in zip(plan.group, plan.group_offsets):
        if o in state.coerced_groups:
            coerced[column] = state.coerced_groups[o].done()
    if plan.pivot is not None and plan.pivot_offset in state.coerced_groups:
        coerced[plan.pivot] = state.coerced_groups[plan.pivot_offset].done()

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
                                   via, keyed)
        else:
            partials, runs, n = state.totals[m.offset]
            totals[m.name] = Total(math.fsum(partials), sheet, first_column + m.offset,
                                   runs.bounds, n, via, keyed)
    left_out = {c: {r: counters[r].done() for r in REASONS if r in counters}
                for c, counters in state.left_out.items()}
    cleaned = {region.columns[o]: state.cleaned[o].done()
               for o in sorted(state.cleaned)} \
        if isinstance(region, (Region, JoinedTable)) else {}
    return Derived(
        address=region.address.text, id=plan.id, columns=tuple(columns),
        key_columns=len(keys), types=types, rows=rows, unrounded=unrounded,
        provenance=[where for _, _, where in built],
        rounding={"rule": ROUNDING, "digits": by_measure},
        matched=state.matched, used=state.used, groups=groups, left_out=left_out,
        excluded=excluded, coerced=coerced, cleaned=cleaned,
        not_compared={c: counter.done() for c, counter in state.not_compared.items()},
        hidden=state.hidden.done(), totals=totals, pivot_columns=pivot_columns,
        pivot_unlisted=[{"name": _label(v, t), "v": v, "t": t, "rows": n}
                        for v, t, n in state.unlisted.values()],
        date1904=plan.date1904, source_key=keyed.column if keyed is not None else None,
        cache=_cache(state.cache, region))


def _cache(counted: dict, region: Any) -> dict[str, int]:
    """The formula caches a result's values came from: those read, and a dataset's own
    when derived from one, in ooxml-edit's order (``CACHE_STATUSES``, then ``external``)."""
    total = dict(counted)
    if isinstance(region, DatasetTable) and region.dataset.cache:
        for key, count in region.dataset.cache.items():
            total[key] = total.get(key, 0) + count
    return {key: total[key] for key in (*CACHE_STATUSES, CACHE_EXTERNAL) if total.get(key)}


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


def _row_sources(group: _Group | None, sheet, column: int, via, keyed) -> Provenance:
    if group is None:
        return Provenance(sheet, column, array("l"), 0, via=via, keyed=keyed)
    return Provenance(sheet, column, group.runs.bounds, group.count, via=via, keyed=keyed)


def _cell_sources(m: _Measure, cell: _Cell | None, sheet, first_column: int,
                  plan: _Plan, via, keyed) -> Provenance:
    column = first_column + (plan.row_offset if m.offset is None else m.offset)
    if cell is None:
        return Provenance(sheet, column, array("l"), 0, via=via, keyed=keyed)
    if m.offset is None:
        return Provenance(sheet, column, cell.runs.bounds, cell.count, via=via, keyed=keyed)
    acc = cell.accs[m.offset]
    return Provenance(sheet, column, acc.runs.bounds, acc.count,
                      {r: acc.excluded[r] for r in REASONS if r in acc.excluded}, via=via,
                      keyed=keyed)


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
                                  "are numbers: a sum, mean, median, min, max, first, last "
                                  "or share of numbers", field="round", valid_options=numeric)
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


def _not_output(name: Any, named: list[str], read: Mapping[str, str], pivot: bool) -> str:
    """Why ``name`` sorts nothing: the output columns, and the one most likely meant."""
    text = f"sort names an output column: one of {', '.join(named)}"
    if pivot:
        text += " or a pivot's (see pivot_columns)"
    if not isinstance(name, str):
        return text
    meant = read.get(name.casefold())
    if meant is not None and meant in named:
        return f"{text}; {name!r} is an input column: did you mean {meant!r}, its measure?"
    match = close_match(name, named)
    return f"{text}; did you mean {match!r}?" if match is not None else text


def _sorted(built: list, sort: Any, columns: list[str], named: list[str],
            read: Mapping[str, str] | None = None, pivot: bool = False) -> list:
    """``built`` sorted by ``sort``; a refusal lists ``named`` (the columns not named by a
    pivot's values, which are content) as the options, and suggests one: for an input
    column a measure reads (``read``, by its name folded), that measure's column; else the
    closest of ``named`` (:func:`ooxml_edit.tools.schema.close_match`)."""
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
            raise DeriveError("bad_sort", _not_output(name, named, read or {}, pivot),
                              field="sort",
                              valid_options=named)
        keys.append((folded[name.casefold()], order == "desc"))
    for name, desc in reversed(keys):
        own = [entry for entry in built if entry[1][name] is not None]
        none = [entry for entry in built if entry[1][name] is None]
        own.sort(key=lambda entry: entry[1][name], reverse=desc)
        built = own + none
    return built
