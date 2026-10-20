"""Column profiles: what a column of a table, name or region holds, in one streamed pass.

What ``describe`` lists per column, and ``xlsx_profile`` gives in full (ROADMAP.md,
"Tools")::

    products = wb.table("t:Products")
    products.column("Price").stats()     # ColumnStats(type='number', fill=0.98, min=..., ...)
    wb.column("t:Products[Price]").stats()
    wb.profile("t:Products")              # every column's, in one pass over the table

**What is counted.**  A column's cells are its data rows (header and totals rows
excluded).  ``rows`` is how many; ``cells`` how many hold a value or a formula, as the used
range counts them; ``filled`` how many hold a value (a formula without a cached value
holds none: it is flagged ``no_cached_value``).  ``fill`` is ``filled / rows``;
``formula_share`` is ``formulas / cells``.  Fractions are exact (no rounding: ROADMAP.md,
"Design rules", 3; whoever shows them rounds them).

**Type and mix.**  ``types`` counts the values by type (:data:`~xlsx_agent.read.types.TYPES`,
blanks aside), each with its first addresses.  ``type`` is the type most values have (a
tie goes to the type :data:`~xlsx_agent.read.types.TYPES` lists first), ``"blank"`` when
there are none; with more than one type the column is flagged ``mixed_types``, counting the
values of the others (ROADMAP.md, "Messy workbooks": "the column's type is the majority;
the minority is counted by type with first addresses").

**Measures** are over the values of the column's type only, and say how many
(``measured``).  ``number`` and ``percent``: ``min``, ``max`` and ``mean`` as numbers, the
mean the exact sum (as :func:`math.fsum` sums, kept as partials while streaming) over the
count.  ``date``, ``time`` and ``datetime``: as dates, ISO 8601 text as the values are
(:mod:`xlsx_agent.read.types`), ordered by their serials; the mean is the mean serial read
in the workbook's date system as that type.  A value written as ISO text (``t="d"``) has
no serial: such a column's min and max are its text's order (ISO 8601 text sorts as time
does) and it has no mean.  Other types have no min, max or mean.

**Distinct values.**  ``distinct`` counts the different values (type and value: ``1``
and ``"1"`` are two), up to :data:`DISTINCT_CAP`; past it, ``None`` (so many that the count
was not kept).  With :data:`FEW_VALUES` or fewer, ``values`` lists each with its count,
most common first.  **Text lengths**: the median and the longest, in characters (code
points) and in words (the runs of characters between white space, as :meth:`str.split`
splits them), over the text values, and the address of the longest text in characters
(the first, of several as long); kept as histograms of lengths, which a cell's limit
bounds.  ``xlsx_query``'s ``words_gt`` test finds the cells of more words than a limit.

**Dates stored as text.**  A date column's range is its ``min`` and ``max``; its cells
that hold a date as text are flagged ``date_as_text``, counted, and not in the range.
``date_text`` gives their own range: the least and the greatest of them read as dates
(:func:`~xlsx_agent.read.types.text_date`, ISO 8601, ordered as ISO text sorts; a
two-digit year is not read), with how many were read -- in a date column and in a text
column of dates alike.

**Allowed values.**  ``validation``: each data-validation list (``type="list"``) that
covers some of the column's cells: an inline list's values, or the values of the range
(or the defined name over one) it takes them from, up to :data:`ALLOWED_CAP`; a source that
is not a range is given its reason (:data:`~xlsx_agent.read.names.NOT_A_RANGE`), never
evaluated, and one over cells a declared table hides from the model
(:mod:`xlsx_agent.declare`) is not read: its reason is ``withheld``.  Excel writes an
inline list as one quoted, comma-separated literal; how Excel treats space around the
commas is not measured here, and it is kept as written.

**Flags** (facts, not verdicts: ROADMAP.md, "Design rules", 9), each a count with the
first addresses: every value issue the reader met (``number_as_text``, ``date_as_text``,
``no_cached_value``, ``date_out_of_range``, ``bad_value``...; :attr:`Value.issue
<xlsx_agent.read.types.Value.issue>`), ``at_limit`` (text of 32,767 characters: probably
cut), ``volatile`` (a formula calling ``NOW``, ``TODAY``, ``RAND``, ``OFFSET`` or
``INDIRECT``: its value is as of the last save), ``hidden`` (cells on a hidden sheet, row or
column), ``cleaned`` (values a declared column's rule changed: :mod:`xlsx_agent.clean`)
and ``mixed_types``.  A declared column split into lists (type ``list``) has its
distinct values counted as whole lists, and no text lengths.

**Bounded.**  One streamed pass (a row at a time, :meth:`Workbook.read
<xlsx_agent.workbook.Workbook.read>`); what is kept per column is counts, the sum's
partials, the extremes, a set of at most :data:`DISTINCT_CAP` value keys (long text is kept
as a 16-byte digest), the length histograms, the longest text's address, the range of
dates stored as text, and at most :data:`FIRST_ADDRESSES` addresses per count.

**Content and diagnostics** (ROADMAP.md, "Design rules", 7).  A profile is a result: its
``min``, ``max``, ``mean``, ``values``, ``date_text`` and allowed values are cell content,
returned to the caller as :meth:`Workbook.read <xlsx_agent.workbook.Workbook.read>`
returns cells.  Nothing
here puts content in a diagnostic: the flags are codes, counts and addresses; an error
raised here names the address, columns and sheets only; a validation's formula text is
kept off the repr.
"""

from __future__ import annotations

import hashlib
import math
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Callable, Iterable, Sequence

from ooxml_edit.sheets import Area, Cell

from .address import RangeAddress, quote_sheet
from .read.names import range_of, sheet_names
from .read.regions import Region
from .read.stream import Validation
from .read.types import CACHE_STATUSES, TYPES, Value, serial_text, text_date

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: Distinct values are listed when there are this many or fewer (ROADMAP.md, "Tools":
#: ``describe``'s "distinct values when 12 or fewer").
FEW_VALUES = 12

#: The most distinct values counted per column; past it the count is not kept.
DISTINCT_CAP = 10_000

#: The most addresses a count keeps: its first ones, in reading order.
FIRST_ADDRESSES = 3

#: The most allowed values a validation list gives.
ALLOWED_CAP = 1_000

#: The types min, max and mean are given for.
MEASURED_TYPES = ("number", "percent", "date", "time", "datetime")

_DATE_TYPES = ("date", "time", "datetime")
#: Text longer than this is kept as a digest in the distinct set.
_KEY_TEXT = 64


@dataclass(frozen=True)
class Count:
    """How many cells, and the first of them by address (at most
    :data:`FIRST_ADDRESSES`, in reading order)."""

    count: int
    first: tuple[str, ...] = ()

    def as_dict(self) -> dict:
        return {"count": self.count, "first": list(self.first)}


@dataclass(frozen=True)
class Distinct:
    """One distinct value of a column with few: ``v`` and ``t`` as a cell's, and how many
    cells hold it."""

    v: object
    t: str
    count: int

    def as_dict(self) -> dict:
        return {"v": self.v, "t": self.t, "count": self.count}


@dataclass(frozen=True)
class AllowedList:
    """A data-validation list over some of a column's cells.

    ``values`` what it allows (an inline list's text, or the values of its source range
    in order, blanks aside), ``None`` when its source is not a range and ``reason`` says
    why (:data:`~xlsx_agent.read.names.NOT_A_RANGE`, or ``"unknown_name"``), or is not
    read (``"withheld"``: it holds cells hidden from the model); ``source``
    the range the values come from, as an address (``None`` for an inline list); ``cut``
    the source held more than :data:`ALLOWED_CAP` values; ``cells`` how many of the
    column's data cells the rule covers.  ``formula`` is the rule's text, for a program
    only."""

    values: tuple | None
    cells: int
    source: str | None = None
    reason: str | None = None
    cut: bool = False
    formula: str | None = field(default=None, repr=False)

    def as_dict(self) -> dict:
        out: dict = {"values": list(self.values) if self.values is not None else None,
                     "cells": self.cells}
        for name in ("source", "reason"):
            if getattr(self, name) is not None:
                out[name] = getattr(self, name)
        if self.cut:
            out["cut"] = True
        return out


@dataclass(frozen=True)
class ColumnStats:
    """A column's profile (see the module's notes).

    ``address`` the column's (``t:Products[Price]``), ``column`` its name, ``range`` its
    data cells (``Stock!C2:C40``; ``None`` with no data rows); ``rows``, ``cells``,
    ``filled``, ``formulas``, ``fill`` and ``formula_share``; ``type`` and ``types``;
    ``min``, ``max``, ``mean`` and ``measured``; ``distinct`` and ``values``;
    ``text_length_median`` and ``text_length_max`` (characters), ``text_words_median`` and
    ``text_words_max`` (words), ``text_longest`` (the address of the longest text);
    ``date_text_min``, ``date_text_max`` and ``date_text_measured`` (the dates stored as
    text, as ISO 8601); ``validation``; ``flags`` by code.  For a column of a declared
    table (:mod:`xlsx_agent.declare`), ``header`` its header's text when it is addressed
    by an id, and ``description`` the caller's; ``range`` each sheet's, joined by ``,``,
    for a table over several sheets.  ``cache``: the column's formula cells by cache
    status (``cached_unverified``, ``invalidated``, ``absent``; empty without formulas);
    a formula referring to another workbook is flagged ``external_formula``."""

    address: str
    column: str
    range: str | None
    rows: int
    cells: int
    filled: int
    formulas: int
    type: str
    types: dict[str, Count]
    fill: float
    formula_share: float
    min: object = None
    max: object = None
    mean: object = None
    measured: int = 0
    distinct: int | None = 0
    values: tuple[Distinct, ...] | None = None
    text_length_median: float | None = None
    text_length_max: int | None = None
    text_words_median: float | None = None
    text_words_max: int | None = None
    text_longest: str | None = None
    date_text_min: str | None = None
    date_text_max: str | None = None
    date_text_measured: int = 0
    validation: tuple[AllowedList, ...] = ()
    flags: dict[str, Count] = field(default_factory=dict)
    header: str | None = None
    description: str | None = None
    cache: dict[str, int] = field(default_factory=dict)

    def as_dict(self) -> dict:
        """The profile as plain data; what is unset (no measures, no values, no lengths, no
        validation, no flags) is left out."""
        out: dict = {"address": self.address, "column": self.column, "range": self.range,
                     "rows": self.rows, "cells": self.cells, "filled": self.filled,
                     "fill": self.fill, "formulas": self.formulas,
                     "formula_share": self.formula_share, "type": self.type,
                     "types": {t: c.as_dict() for t, c in self.types.items()},
                     "distinct": self.distinct}
        if self.header is not None:
            out["header"] = self.header
        if self.description is not None:
            out["description"] = self.description
        if self.measured:
            out.update(min=self.min, max=self.max, mean=self.mean, measured=self.measured)
        if self.values is not None:
            out["values"] = [d.as_dict() for d in self.values]
        if self.text_length_max is not None:
            out["text_length"] = {"median": self.text_length_median, "max": self.text_length_max,
                                  "median_words": self.text_words_median,
                                  "max_words": self.text_words_max,
                                  "longest": self.text_longest}
        if self.date_text_measured:
            out["date_text"] = {"min": self.date_text_min, "max": self.date_text_max,
                                "measured": self.date_text_measured}
        if self.validation:
            out["validation"] = [v.as_dict() for v in self.validation]
        if self.flags:
            out["flags"] = {code: c.as_dict() for code, c in self.flags.items()}
        if self.cache:
            out["cache"] = dict(self.cache)
        return out


class Column:
    """One column of a region, as :meth:`Region.column
    <xlsx_agent.read.regions.Region.column>` and :meth:`Workbook.column
    <xlsx_agent.workbook.Workbook.column>` give it."""

    def __init__(self, workbook: "Workbook", region: Region, name: str) -> None:
        self.workbook = workbook
        self.region = region
        #: The column's name as the region spells it.
        self.name = region.columns[region.column_index(name)]

    def __repr__(self) -> str:
        return f"Column({self.address!r})"

    @property
    def address(self) -> str:
        """``t:Products[Price]``."""
        return replace(self.region.address, column=self.name).text

    @property
    def range(self) -> RangeAddress | None:
        """The column's data cells; ``None`` when the region has no data rows."""
        area = self.region.column_area(self.name)
        return RangeAddress(self.region.sheet_ref, area) if area is not None else None

    def stats(self) -> ColumnStats:
        """The column's profile, in one streamed pass over its cells."""
        return profile(self.workbook, self.region, [self.name])[0]


def profile(workbook: "Workbook", region: Region, columns: Sequence[str] | None = None, *,
            observe: "Callable[[list[Value]], None] | None" = None) -> list[ColumnStats]:
    """The profiles of ``region``'s ``columns`` (names, matched case aside; every column
    when ``None``), in the order asked, in one streamed pass over the data rows.
    ``observe`` is given each row the pass reads (from the leftmost column asked to the
    rightmost), for another use of the same pass."""
    names = list(region.columns) if columns is None else \
        [region.columns[region.column_index(name)] for name in columns]
    first = region.area.first.column
    offsets = [region.column_index(name) for name in names]
    tallies = [_Tally(workbook.locale) for _ in names]
    data = region.data
    if data is not None and names:
        left, right = first + min(offsets), first + max(offsets)
        area = Area(Cell(data.first.row, left), Cell(data.last.row, right))
        positions = [first + offset - left for offset in offsets]
        reading = {first + o for o in offsets} if observe is None else None
        for row in workbook._read(RangeAddress(region.sheet_ref, area), columns=reading):
            for tally, position in zip(tallies, positions):
                tally.add(row[position])
            if observe is not None:
                observe(row)
    out = []
    for name, tally in zip(names, tallies):
        area = region.column_area(name)
        allowed = tuple(column_lists(workbook, region, name))
        address = replace(region.address, column=name).text
        stats = tally.stats(address, name, region.located(area), region.rows,
                            workbook.date1904, allowed)
        index = region.column_index(name)
        header = region.headers[index] if region.headers else None
        note = region.notes[index] if region.notes else None
        if (header is not None and header != name) or note is not None:
            stats = replace(stats, header=header if header != name else None,
                            description=note)
        out.append(stats)
    return out


def column_lists(workbook: "Workbook", region: Region, name: str) -> list[AllowedList]:
    """The data-validation lists over a column's data cells, each list that covers any of
    them (each sheet's, for a table over several sheets)."""
    if region.span is None:
        area = region.column_area(name)
        if area is None:
            return []
        sheet = region.sheet if region.sheet else workbook.sheet_by_id(region.sheet_id).name
        return [_allowed(workbook, sheet, rule, area)
                for rule in _list_rules(workbook, workbook.sheet(sheet))
                if _covered(rule.areas, area)]
    column = region.span.columns[region.column_index(name)]
    out = []
    for part in region.span.parts:
        if not part.rows:
            continue
        area = Area(Cell(part.first, column), Cell(part.last, column))
        out += [_allowed(workbook, part.sheet, rule, area)
                for rule in _list_rules(workbook, workbook.sheet(part.sheet))
                if _covered(rule.areas, area)]
    return out


def majority_type(counts: "dict[str, int]") -> str:
    """A column's type from how many values it holds of each type (blanks aside): the type
    most values have, a tie going to the type :data:`~xlsx_agent.read.types.TYPES` lists
    first; ``"blank"`` when it holds none.  The rule ``xlsx_query`` checks filters by."""
    held = {t: n for t, n in counts.items() if n and t != "blank"}
    return max(held, key=lambda t: (held[t], -TYPES.index(t)), default="blank")


# ------------------------------------------------------------------------------------------
# The pass
# ------------------------------------------------------------------------------------------


class _Counter:
    """A count with its first addresses."""

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


class _Tally:
    """What one column's pass keeps."""

    def __init__(self, locale: object = None) -> None:
        self.locale = locale
        self.cells = self.filled = self.formulas = 0
        self.cache: dict[str, int] = {}
        self.types: dict[str, _Counter] = {}
        self.flags: dict[str, _Counter] = {}
        self.measures: dict[str, _Measure] = {}
        self.seen: set | None = set()
        self.few: dict | None = {}
        self.lengths: dict[int, int] = {}
        self.words: dict[int, int] = {}
        self.longest: tuple[int, str] | None = None
        self.dates: list | None = None     # [low, high, count] of dates stored as text

    def add(self, value: Value) -> None:
        if value.hidden:
            self._flag("hidden", value)
        if value.issue is not None:
            self._flag(value.issue, value)
        if value.at_limit:
            self._flag("at_limit", value)
        if value.volatile:
            self._flag("volatile", value)
        if value.cleaned:
            self._flag("cleaned", value)
        if value.external:
            self._flag("external_formula", value)
        t = value.t
        if value.formula:
            self.formulas += 1
            if value.cache is not None:
                self.cache[value.cache] = self.cache.get(value.cache, 0) + 1
        if t == "blank":
            if value.formula:
                self.cells += 1
            return
        self.cells += 1
        self.filled += 1
        counter = self.types.get(t)
        if counter is None:
            counter = self.types[t] = _Counter()
        counter.add(value)
        if t in MEASURED_TYPES:
            measure = self.measures.get(t)
            if measure is None:
                measure = self.measures[t] = _Measure()
            measure.add(value)
        elif t == "text":
            text = value.v
            length = len(text)
            self.lengths[length] = self.lengths.get(length, 0) + 1
            words = len(text.split())
            self.words[words] = self.words.get(words, 0) + 1
            if self.longest is None or length > self.longest[0]:
                self.longest = (length, value.a)
            if value.issue == "date_as_text" and self.locale is not None:
                self._date_text(text)
        self._distinct(value)

    def _date_text(self, text: str) -> None:
        shown = text_date(text, self.locale, local_dates=True)
        if shown is None:
            return
        dates = self.dates
        if dates is None:
            self.dates = [shown, shown, 1]
            return
        dates[0], dates[1] = min(dates[0], shown), max(dates[1], shown)
        dates[2] += 1

    def _flag(self, code: str, value: Value) -> None:
        counter = self.flags.get(code)
        if counter is None:
            counter = self.flags[code] = _Counter()
        counter.add(value)

    def _distinct(self, value: Value) -> None:
        seen = self.seen
        if seen is None:
            return
        v = value.v
        if isinstance(v, str) and len(v) > _KEY_TEXT:
            key = (value.t, hashlib.blake2b(v.encode("utf-8", "surrogatepass"),
                                            digest_size=16).digest())
        else:
            key = (value.t, v)
        few = self.few
        if few is not None:
            entry = few.get(key)
            if entry is not None:
                entry[2] += 1
                return
            if len(few) < FEW_VALUES:
                few[key] = [v, value.t, 1]
            else:
                self.few = None
        if key not in seen:
            if len(seen) >= DISTINCT_CAP:
                self.seen = None
                return
            seen.add(key)

    def stats(self, address: str, name: str, range_text: str | None, rows: int,
              date1904: bool, allowed: tuple[AllowedList, ...]) -> ColumnStats:
        types = {t: self.types[t].done() for t in TYPES if t in self.types}
        majority = majority_type({t: c.count for t, c in types.items()})
        flags = {code: counter.done() for code, counter in self.flags.items()}
        if len(types) > 1:
            first: list[str] = []
            for t, count in types.items():
                if t != majority:
                    first.extend(count.first)
            # the minority's first addresses in reading order: addresses of one column sort
            # by row
            first.sort(key=_row_of)
            flags["mixed_types"] = Count(sum(c.count for t, c in types.items() if t != majority),
                                         tuple(first[:FIRST_ADDRESSES]))
        flags = dict(sorted(flags.items()))
        low = high = mean = None
        measured = 0
        measure = self.measures.get(majority)
        if measure is not None:
            low, high, mean = measure.result(majority, date1904)
            measured = measure.count
        if self.seen is None:
            distinct, values = None, None
        else:
            distinct = len(self.seen)
            values = None
            if self.few is not None:
                ordered = sorted(self.few.values(), key=lambda entry: -entry[2])
                values = tuple(Distinct(v, t, n) for v, t, n in ordered)
        median, longest = _median(self.lengths)
        words_median, words_max = _median(self.words)
        dates = self.dates or [None, None, 0]
        return ColumnStats(address=address, column=name, range=range_text, rows=rows,
                           cells=self.cells, filled=self.filled, formulas=self.formulas,
                           type=majority, types=types,
                           fill=self.filled / rows if rows else 0.0,
                           formula_share=self.formulas / self.cells if self.cells else 0.0,
                           min=low, max=high, mean=mean, measured=measured, distinct=distinct,
                           values=values, text_length_median=median, text_length_max=longest,
                           text_words_median=words_median, text_words_max=words_max,
                           text_longest=self.longest[1] if self.longest else None,
                           date_text_min=dates[0], date_text_max=dates[1],
                           date_text_measured=dates[2],
                           validation=allowed, flags=flags,
                           cache={k: self.cache[k] for k in CACHE_STATUSES if k in self.cache})


class _Measure:
    """Min, max and the exact sum of one type's values, streamed."""

    __slots__ = ("partials", "count", "low", "high", "serials", "elapsed", "clock")

    def __init__(self) -> None:
        self.partials: list[float] = []
        self.count = 0
        self.low: tuple | None = None
        self.high: tuple | None = None
        self.serials = True  # every value has a serial (dates) or is a number
        self.elapsed = self.clock = False

    def add(self, value: Value) -> None:
        self.count += 1
        if value.t in _DATE_TYPES:
            if isinstance(value.v, str) and value.v.startswith("PT"):
                self.elapsed = True
            else:
                self.clock = True
            if value.serial is None:
                self.serials = False
            number = value.serial
        else:
            number = value.v
        if number is not None:
            _add_partial(self.partials, float(number))
        key = number if number is not None else value.v
        if self.low is None or _less(key, self.low[0]):
            self.low = (key, value.v)
        if self.high is None or _less(self.high[0], key):
            self.high = (key, value.v)

    def result(self, t: str, date1904: bool) -> tuple:
        if not self.count:
            return None, None, None
        if t not in _DATE_TYPES:
            return self.low[1], self.high[1], math.fsum(self.partials) / self.count
        if not self.serials:
            # ISO text, ordered as text: comparable only when none is a serial or a
            # duration, so that every key is text
            if self.elapsed or any(isinstance(k, float) for k in (self.low[0], self.high[0])):
                return None, None, None
            return self.low[1], self.high[1], None
        if self.elapsed and self.clock:
            return self.low[1], self.high[1], None
        serial = math.fsum(self.partials) / self.count
        return self.low[1], self.high[1], serial_text(serial, t, date1904, self.elapsed)


def _less(a, b) -> bool:
    try:
        return a < b
    except TypeError:  # a serial against ISO text: ordered as text after numbers
        return isinstance(a, float)


def _add_partial(partials: list[float], x: float) -> None:
    """Shewchuk's exact running sum (what :func:`math.fsum` keeps), one value at a time:
    ``partials`` stays as short as the exponents met make it."""
    i = 0
    for y in partials:
        if abs(x) < abs(y):
            x, y = y, x
        hi = x + y
        lo = y - (hi - x)
        if lo:
            partials[i] = lo
            i += 1
        x = hi
    partials[i:] = [x]


def _median(lengths: dict[int, int]) -> tuple[float | None, int | None]:
    total = sum(lengths.values())
    if not total:
        return None, None
    ordered = sorted(lengths.items())
    middle = [(total - 1) // 2, total // 2]  # the two middle positions (one when odd)
    found: list[int] = []
    seen = 0
    for length, n in ordered:
        while len(found) < 2 and middle[len(found)] < seen + n:
            found.append(length)
        seen += n
        if len(found) == 2:
            break
    median = (found[0] + found[1]) / 2
    return (int(median) if median.is_integer() else median), ordered[-1][0]


def _row_of(address: str) -> int:
    digits = ""
    for char in reversed(address):
        if not char.isdigit():
            break
        digits = char + digits
    return int(digits)


# ------------------------------------------------------------------------------------------
# Validation lists
# ------------------------------------------------------------------------------------------


def _list_rules(workbook: "Workbook", sheet: "Sheet") -> list[Validation]:
    return [rule for rule in workbook.scan(sheet).validations if rule.type == "list"]


def _covered(areas: Iterable[Area], column: Area) -> int:
    """How many cells of ``column`` (one column's area) ``areas`` cover, overlaps counted
    once."""
    spans = []
    for area in areas:
        if area.first.column <= column.first.column <= area.last.column:
            top, bottom = max(area.first.row, column.first.row), min(area.last.row, column.last.row)
            if top <= bottom:
                spans.append((top, bottom))
    total, end = 0, 0
    for top, bottom in sorted(spans):
        top = max(top, end + 1)
        if top <= bottom:
            total += bottom - top + 1
            end = bottom
    return total


def _allowed(workbook: "Workbook", sheet: str, rule: Validation,
             column: Area) -> AllowedList:
    cells = _covered(rule.areas, column)
    text = (rule.formula1 or "").strip()
    if text.startswith("="):
        text = text[1:].lstrip()
    if not text:
        return AllowedList(None, cells, reason="formula", formula=rule.formula1)
    literal = _literal(text)
    if literal is not None:
        return AllowedList(tuple(workbook._redacted(literal).split(",")), cells,
                           formula=rule.formula1)
    sheets = sheet_names(workbook.sheets)
    here = quote_sheet(sheet)
    target, reason = range_of(text if "!" in text else f"{here}!{text}", sheets)
    if target is None and reason == "formula" and "!" not in text:
        target, reason = _named_range(workbook, sheet, text)
    if target is None:
        return AllowedList(None, cells, reason=reason, formula=rule.formula1)
    if _withheld(workbook, target):
        return AllowedList(None, cells, source=target.text, reason="withheld",
                           formula=rule.formula1)
    values: list = []
    cut = False
    for row in workbook.read(target):
        for value in row:
            if value.t == "blank":
                continue
            if len(values) >= ALLOWED_CAP:
                cut = True
                break
            values.append(value.v)
        if cut:
            break
    return AllowedList(tuple(values), cells, source=target.text, cut=cut,
                       formula=rule.formula1)


def _withheld(workbook: "Workbook", target: RangeAddress) -> bool:
    """Whether a list's source holds cells of a column hidden from the model
    (:mod:`xlsx_agent.declare`)."""
    try:
        sheet = workbook.sheet_of(target)
    except KeyError:
        return False
    area = target.area
    return any(not (a.last.row < area.first.row or area.last.row < a.first.row
                    or a.last.column < area.first.column or area.last.column < a.first.column)
               for a in workbook.regions.withheld(sheet.name))


def _literal(text: str) -> str | None:
    """A formula that is one string literal (``"Red,Green"``, ``""`` for a quote inside)
    as its text; ``None`` for anything else."""
    if len(text) < 2 or text[0] != '"' or text[-1] != '"':
        return None
    body = text[1:-1]
    if '"' in body.replace('""', ""):
        return None
    return body.replace('""', '"')


def _named_range(workbook: "Workbook", sheet: str,
                 text: str) -> tuple[RangeAddress | None, str | None]:
    """A list's source given as a defined name: the one local to the rule's sheet first,
    then the workbook-wide one, as Excel looks a name up from a sheet."""
    folded = text.casefold()
    local = [n for n in workbook.names
             if n.scope is not None and n.scope.casefold() == sheet.casefold()
             and n.name.casefold() == folded]
    wide = [n for n in workbook.names if n.scope is None and n.bad_scope is None
            and n.name.casefold() == folded]
    found = local or wide
    if len(found) != 1:
        return None, "unknown_name" if not found else "formula"
    name = found[0]
    if name.range is None:
        return None, name.reason or "formula"
    return name.range, None


__all__ = ["ALLOWED_CAP", "DISTINCT_CAP", "FEW_VALUES", "FIRST_ADDRESSES", "AllowedList",
           "Column", "ColumnStats", "Count", "Distinct", "profile"]
