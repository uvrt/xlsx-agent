"""Datasets: a derive's result kept as a table of values, each with the cells it came from
(ROADMAP.md, "Decisions already made", 2; "Addresses": ``d:``)::

    ds = wb.derive("t:Sales", group_by=["Region"],
                   measures=[{"fn": "sum", "column": "Amount", "as": "total"}], id="by_region")
    data = ds.dataset()            # ooxml_edit.tools.Dataset: columns, rows, sources, types
    data.rows                      # (("North", 3.68), ("South", 0.13), ...): rounded
    data.sources[0][1]             # ("Sales!C2:C3", "Sales!C5"): the cells of North's total;
                                   # by key when the table has a key column:
                                   # ("t:Sales/{S-01,S-02,S-04}/Amount",)
    table = DatasetTable(data, "by_region")
    table.read("d:by_region[total]")          # [[DatasetValue(d:by_region/#1/total)], ...]
    wb.derive(ds, measures=[{"fn": "max", "column": "total"}])   # from the dataset

**What a dataset is.**  ooxml-edit's :class:`~ooxml_edit.tools.Dataset` (O2): columns,
rows, and beside every value its sources as address strings -- by key
(``t:Sales/{S-01,S-02}/Amount``) when the table has a key column, else ranges
(``Sales!C2:C3``; :mod:`xlsx_agent.derive`) -- which the tool layer stores and hands on and
never reads.  :func:`as_dataset` makes one of a
:class:`~xlsx_agent.derive.Derived`: the rows as reported (rounded), the key columns as
``keys``, each column's type as its label, and the :class:`Derived` itself as ``detail``,
where the unrounded values stay (ROADMAP.md, "Design rules", 3).  In a tool session it is
kept under ``d:<id>`` (:mod:`xlsx_agent.tools.derive`); in a script it is a value.

**Reading one** (:class:`DatasetTable`).  ``d:by_region`` is its rows, ``[total]`` one
column, ``/#2`` a row by position, ``/North`` a row by key and ``/North/total`` one value,
as for tables (R2).  A row by key needs exactly one key column (a derive grouped by one
column; ``keys`` is 1): the key matches the column's value as a table's key does (text
exactly, a whole number by its digits); none is ``not_found``, two ``duplicate_key`` with
their positions.  Each value comes back as a :class:`DatasetValue`, typed, whose address
names it by position (``d:by_region/#2/total``): a row's position never changes, since a
dataset never does.

**Deriving from one** (:func:`xlsx_agent.derive.derive` with a :class:`Derived`, a
:class:`Dataset` or a :class:`DatasetTable` as ``from_``).  The dataset is a table like any
other: its columns are typed by their values (a column's label first, so a date column's
ISO text is a date), and filters, groups, measures, bins, pivots and sorts work as on a
sheet.  A Python ``date``, ``datetime`` or ``time`` is that type with no label, read as
ISO 8601 text as a cell's date is; ISO text in an unlabelled column is text, flagged
``date_as_text`` as a cell's would be, so a ``min`` on it is refused with the way out --
the ``date`` label (:meth:`Dataset.from_records <ooxml_edit.tools.Dataset.from_records>`
``types=``) or ``coerce=["date_as_text"]``.  The labels a dataset's columns take are
ooxml-edit's ``TYPE_LABELS``: ``number``, ``percent``, ``date``, ``time``, ``datetime``,
``text``, ``bool``, ``error``.  A :class:`Derived`'s **unrounded** values are measured, so rounding happens once,
at the end.  **Provenance carries through**: a value derived from a dataset has as its
sources the cells the dataset's values came from -- the original cells, merged into
compact ranges again, or the keys of one table's column gathered into one key set, in the
order first met -- not the dataset's positions; ``count`` is how many of the dataset's
values it used, and what was left out is named by their ``d:`` addresses.

**Source units.**  A range is one source and a key set one per key: what a result shows
of a value's sources is its first :data:`~xlsx_agent.tools.derive.SOURCES_SHOWN` units
(:func:`cut_sources`), ``more`` counting the rest, and a read of one value's sources pages
through them by unit, a page's keys of one column gathered into one key set again.
"""

from __future__ import annotations

import datetime as _dt
import warnings
from dataclasses import dataclass, replace
from typing import Any, Iterable, Iterator, Sequence

from ooxml_edit.sheets import Area, Cell
from ooxml_edit.tools import DATASET_PREFIX, Dataset

from .address import (Address, AddressError, DatasetAddress, NameNotFound, RangeAddress,
                      SheetRef, TableAddress, format_address, nearest, parse_address)

from .clean import CleanedWarning
from .profile import FIRST_ADDRESSES, Count
from .read.regions import RegionError, key_text
from .read.types import _MS_PER_DAY, TYPES, Value, _clock, looks_like_date

#: What a dataset without an id (a script's :class:`Derived` made without ``id=``) is called
#: in the addresses of its values.
UNNAMED = "derived"

_MOMENT_LABELS = ("date", "time", "datetime")


@dataclass(slots=True, repr=False)
class DatasetValue(Value):
    """One value of a dataset, typed as a cell is (:class:`~xlsx_agent.read.types.Value`):
    ``row`` its position (from 1), ``column`` its column's (from 1), ``name`` the column's
    name, ``dataset`` the dataset's id; :attr:`a` its address, ``d:by_region/#2/total``.
    ``sources``, when asked for, the first of the cells it came from, and ``more`` how many
    others there are (in units: a range, or one key of a key set; :func:`cut_sources`)."""

    dataset: str = UNNAMED
    name: str = ""
    sources: tuple[str, ...] | None = None
    more: int = 0

    @property
    def a(self) -> str:
        return DatasetAddress(self.dataset, column=self.name, position=self.row).text

    def __repr__(self) -> str:
        return f"DatasetValue({self.a}, t={self.t!r})"

    def as_dict(self) -> dict:
        out = {"a": self.a, "v": self.v, "t": self.t}
        if self.sources is not None:
            out["sources"] = list(self.sources)
            if self.more:
                out["more"] = self.more
        return out


@dataclass(frozen=True)
class DatasetSelection:
    """What a ``d:`` address names in a dataset: ``text`` the address as given back,
    ``kind`` (``dataset``, ``column``, ``row`` or ``value``; ``rows`` or ``values`` for a
    row range), ``first`` and ``last`` the
    rows (positions from 1; ``last`` 0 when there are none), ``offsets`` the columns (from
    0)."""

    text: str
    kind: str
    first: int
    last: int
    offsets: tuple[int, ...]

    @property
    def rows(self) -> int:
        return max(self.last - self.first + 1, 0)


def as_dataset(derived: Any) -> Dataset:
    """A :class:`~xlsx_agent.derive.Derived` as an ooxml-edit :class:`Dataset` (see the
    module's notes): rows rounded, sources as range text, the :class:`Derived` as
    ``detail``."""
    columns = tuple(derived.columns)
    rows = tuple(tuple(row[c] for c in columns) for row in derived.rows)
    sources = tuple(tuple(tuple(str(s) for s in where[c].sources) for c in columns)
                    for where in derived.provenance)
    return Dataset(columns, rows, sources=sources, keys=derived.key_columns,
                   types=tuple(derived.types.get(c) for c in columns), detail=derived,
                   cache=getattr(derived, "cache", None) or None)


class DatasetTable:
    """A dataset as a table: what the derive engine reads (as it reads a
    :class:`~xlsx_agent.read.regions.Region`) and what ``xlsx_read`` pages through.

    ``dataset`` an ooxml-edit :class:`Dataset`, or a :class:`Derived` (made one with
    :func:`as_dataset`); ``id`` its id, without ``d:`` (by default the :class:`Derived`'s,
    else :data:`UNNAMED`).  Measured values are the :class:`Derived`'s unrounded ones when
    ``detail`` is one; read values are the dataset's rows, as reported."""

    def __init__(self, dataset: Any, id: str | None = None) -> None:
        from .derive import Derived

        if isinstance(dataset, Derived):
            dataset = as_dataset(dataset)
        if not isinstance(dataset, Dataset):
            raise TypeError(f"a DatasetTable is made of a Dataset or a Derived, not "
                            f"{type(dataset).__name__}")
        detail = dataset.detail
        derived = detail if isinstance(detail, Derived) else None
        if id is None:
            id = derived.id if derived is not None and derived.id else UNNAMED
        if isinstance(id, str) and id.startswith(DATASET_PREFIX):
            id = id[len(DATASET_PREFIX):]
        self.address = DatasetAddress(id)
        self.id = id
        self.dataset = dataset
        self.columns = dataset.columns
        self.keys = dataset.keys
        self.date1904 = bool(getattr(derived, "date1904", False))
        labels = dataset.types or (None,) * len(self.columns)
        self.labels = tuple(label if label in TYPES else None for label in labels)
        measured = dataset.rows
        if derived is not None and len(derived.unrounded) == len(dataset.rows):
            measured = tuple(tuple(row.get(c) for c in self.columns)
                             for row in derived.unrounded)
        self.measured = measured

    @property
    def rows(self) -> int:
        return len(self.dataset.rows)

    def __repr__(self) -> str:
        return f"DatasetTable({self.address.text}, rows={self.rows})"

    # -- columns ---------------------------------------------------------------------------

    def column_index(self, name: str) -> int:
        """The position (from 0) of the column called ``name``, matched case aside, as a
        region's are; :class:`~xlsx_agent.address.NameNotFound` with the nearest names."""
        folded = name.casefold()
        for index, column in enumerate(self.columns):
            if column.casefold() == folded:
                return index
        raise NameNotFound(f"{self.address.text} has no column {name!r}; its columns are "
                           f"{list(self.columns)}", what="column", name=name,
                           candidates=list(self.columns),
                           valid_options=nearest(name, self.columns))

    # -- values ----------------------------------------------------------------------------

    def value(self, position: int, offset: int, *, measured: bool = False,
              sources: int | None = None) -> DatasetValue:
        """The value at row ``position`` (from 1) and column ``offset`` (from 0), typed:
        reported, or ``measured`` (unrounded); with its first ``sources`` cells when
        given."""
        rows = self.measured if measured else self.dataset.rows
        v, t, issue = _typed(rows[position - 1][offset], self.labels[offset])
        out = DatasetValue(sheet="", row=position, column=offset + 1, v=v, t=t,
                           issue=issue, dataset=self.id, name=self.columns[offset])
        if sources is not None:
            out.sources, out.more = cut_sources(self.cell_sources(position, offset), sources)
        return out

    def cell_sources(self, position: int, offset: int) -> tuple[str, ...]:
        """The sources the dataset holds for one value, as given (``()``: none)."""
        held = self.dataset.sources
        return held[position - 1][offset] if held is not None else ()

    def measured_rows(self) -> Iterator[list[DatasetValue]]:
        """Every row, every column, as the derive engine measures them (unrounded)."""
        width = len(self.columns)
        for position in range(1, self.rows + 1):
            yield [self.value(position, o, measured=True) for o in range(width)]

    def sources_of(self, offset: int, runs: Sequence[int]
                   ) -> "tuple[RangeAddress | TableAddress | str, ...]":
        """The cells behind the values of column ``offset`` in the rows ``runs`` (first
        and last positions, flat): every source of each, a column's cells merged into
        compact ranges (a run of rows one range, as a derive writes them), and the keys
        of one table's column into one key set, each in the order first met; a source that
        is neither is kept as it is, once."""
        held = self.dataset.sources
        if held is None:
            return ()
        return tuple(_merged(text for i in range(0, len(runs), 2)
                             for position in range(runs[i], runs[i + 1] + 1)
                             for text in held[position - 1][offset]))

    # -- addresses -------------------------------------------------------------------------

    def select(self, address: "str | DatasetAddress") -> DatasetSelection:
        """What a ``d:`` address names here (see the module's notes):
        :class:`~xlsx_agent.address.NameNotFound` for an unknown column or key;
        :class:`~xlsx_agent.read.regions.RegionError` for a position (or a row range's
        row) past the end (``position_out_of_range``), a key without one key column (``no_key_column``), a
        key two rows share (``duplicate_key``) or a key set (``no_key_set``).  ``*`` is
        every row (``d:by_region/*/total``, the column) or every column
        (``d:by_region/#2/*``, the row); ``d:by_region/#2:#5`` rows 2 to 5, ``/#2:`` row 2
        to the last."""
        if isinstance(address, str):
            parsed = parse_address(address)
            if not isinstance(parsed, DatasetAddress):
                raise AddressError("a dataset is named d:<id>", text=address,
                                   valid_options=[self.address.text])
            address = parsed
        offsets = tuple(range(len(self.columns)))
        if address.column is not None:
            offsets = (self.column_index(address.column),)
        column = self.columns[offsets[0]] if address.column is not None else None
        whole = DatasetAddress(self.id)
        if address.keys is not None:
            raise RegionError("no_key_set", f"a dataset's rows are named one by one, by key "
                              f"or by position ({whole.text}/#1); a key set is for tables",
                              address=whole.text, valid_options=[f"{whole.text}/#1"])
        any_: dict[str, Any] = {"any_row": address.any_row, "any_column": address.any_column}
        if address.row_range is not None:
            first, last = address.row_range
            end = self.rows if last is None else last
            if max(first, end) > self.rows or not self.rows:
                start = min(first, self.rows)
                options = [] if not self.rows else [DatasetAddress(
                    self.id, position=start).text if start == self.rows else DatasetAddress(
                    self.id, row_range=(start, self.rows)).text]
                raise RegionError("position_out_of_range", f"{whole.text} has {self.rows} "
                                  f"rows; there is no row #{max(first, end)} (nothing is "
                                  "cut)", address=whole.text, valid_options=options)
            text = DatasetAddress(self.id, column=column, row_range=address.row_range,
                                  any_column=address.any_column).text
            return DatasetSelection(text, address.kind, first, end, offsets)
        if address.key is None and address.position is None:
            text = DatasetAddress(self.id, column=column, **any_).text
            return DatasetSelection(text, address.kind, 1, self.rows, offsets)
        if address.position is not None:
            position = address.position
            if position > self.rows:
                raise RegionError("position_out_of_range", f"{whole.text} has {self.rows} "
                                  "rows", address=whole.text)
        else:
            assert address.key is not None   # by key: the only form left
            position = self._find_key(address.key)
        text = DatasetAddress(self.id, column=column, position=position,
                              any_column=address.any_column).text
        return DatasetSelection(text, address.kind, position, position, offsets)

    def _find_key(self, wanted: str) -> int:
        whole = DatasetAddress(self.id).text
        if self.keys != 1:
            why = "no key column" if not self.keys else f"{self.keys} key columns"
            raise RegionError("no_key_column", f"{whole} has {why}: name its rows by "
                              f"position ({whole}/#1)", address=whole)
        found = [p for p in range(1, self.rows + 1)
                 if key_text(self.value(p, 0)) == wanted]
        if not found:
            raise NameNotFound(f"no row of {whole} has that key in column "
                               f"{self.columns[0]!r}", what="key", name=wanted)
        if len(found) > 1:
            raise RegionError("duplicate_key", f"{len(found)} rows of {whole} have that key; "
                              "name the row by its position", address=whole,
                              valid_options=[f"{whole}/#{p}" for p in found[:50]])
        return found[0]

    def read(self, address: "str | DatasetAddress", *, sources: int | None = None
             ) -> list[list[DatasetValue]]:
        """The values a ``d:`` address names, row by row, as reported; each with its first
        ``sources`` cells when given."""
        selection = self.select(address)
        return [[self.value(p, o, sources=sources) for o in selection.offsets]
                for p in range(selection.first, selection.last + 1)]


def source_units(sources: Sequence[str]) -> list[tuple[Any, str]]:
    """Sources as units: a key set's keys one by one, as ``(column, key)`` with the
    table's column address; any other source as ``(None, text)``."""
    out: list[tuple[Any, str]] = []
    for text in sources:
        area = _column_area(text)
        if area is not None and isinstance(area[0], TableAddress):
            out.extend((area[0], key) for key in area[1])
        else:
            out.append((None, text))
    return out


def join_units(units: Sequence[tuple[Any, str]]) -> list[str]:
    """Units (:func:`source_units`) as sources again: a run of keys of one column as one
    key set."""
    out: list[str] = []
    column: Any = None
    keys: list[Any] = []   # a run's keys: text, as every unit with a column holds
    for unit_column, text in list(units) + [(None, None)]:
        if unit_column is not None and unit_column == column:
            keys.append(text)
            continue
        if column is not None:
            out.append(_by_keys(column, keys).text)
        column, keys = unit_column, [text]
        if unit_column is None and text is not None:
            out.append(text)
    return out


def cut_sources(sources: Sequence[str], shown: int) -> tuple[tuple[str, ...], int]:
    """The first ``shown`` units of ``sources`` (a range, or one key of a key set) as
    sources, and how many units are left."""
    units = source_units(sources)
    if len(units) <= shown:
        return tuple(sources), 0
    return tuple(join_units(units[:shown])), len(units) - shown


def merge_sources(sources: Iterable[str]) -> list[str]:
    """Source addresses written compactly, each once: a column's cells as ranges, a run of
    rows one range (``Sales!C2``, ``Sales!C3`` -> ``Sales!C2:C3``), and the keys of one
    table's column as one key set (``t:Sales/S-01/Amount``, ``t:Sales/S-02/Amount`` ->
    ``t:Sales/{S-01,S-02}/Amount``), each in the order first met; any other address as it
    is.  The merger a derive from a dataset writes its sources with
    (:meth:`DatasetTable.sources_of`), and the xlsx format's ``merge_sources``: how a deck
    or document tool's result names the cells a table's values came from
    (:mod:`xlsx_agent.tools.docformat`)."""
    return [address if isinstance(address, str) else address.text
            for address in _merged(sources)]


def _merged(sources: Iterable[str]) -> "list[RangeAddress | TableAddress | str]":
    spans: dict[tuple[SheetRef, int], list[tuple[int, int]]] = {}
    keyed: dict[TableAddress, dict[str, None]] = {}
    other: dict[Any, None] = {}
    for text in sources:
        area = _column_area(text)
        if area is None:
            other.setdefault(text, None)
        elif isinstance(area[0], TableAddress):
            group = keyed.setdefault(area[0], {})
            for key in area[1]:
                group.setdefault(key, None)
        else:
            sheet, column, first, last = area
            spans.setdefault((sheet, column), []).append((first, last))
    out: list[RangeAddress | TableAddress | str] = []
    for (sheet, column), pairs in spans.items():
        pairs.sort()
        merged: list[list[int]] = []
        for first, last in pairs:
            if merged and first <= merged[-1][1] + 1:
                merged[-1][1] = max(merged[-1][1], last)
            else:
                merged.append([first, last])
        out.extend(RangeAddress(sheet, Area(Cell(a, column), Cell(b, column)))
                   for a, b in merged)
    out.extend(_by_keys(column, list(keys)) for column, keys in keyed.items())
    out.extend(other)
    return out


#: What a list's items are joined by when a table's rows go into a deck or document: a line
#: break, which their tables show as one cell of several lines.
LIST_JOIN = "\n"


def table_dataset(workbook: Any, address: "str | TableAddress", *,
                  max_cells: int | None = None) -> Dataset:
    """The rows a ``t:`` address names, as an ooxml-edit :class:`Dataset`: what a deck or
    document tool's ``from: "t:Products"`` places (O3; the xlsx format's resolver,
    :mod:`xlsx_agent.tools.docformat`).

    ``t:Products`` (or ``t:Products/*``, ``t:Products/*/*``) is every data row and column;
    ``t:Products[Price]`` (or ``t:Products/*/Price``) one column; ``t:Products/{SKU-0003,
    SKU-0007}`` a key set's rows in its order (``/Price`` their one column),
    ``t:Products/#9:#16`` (``#9:`` to the last) a row range's, and a row by key or by
    position one row.  One value (``t:Products/SKU-0042/Price``) is no rows:
    :class:`RegionError` ``not_rows``.  The columns are the region's names; the header
    and totals rows are not rows.

    Values as :meth:`Workbook.read <xlsx_agent.Workbook.read>` types them: a number or a
    percent the stored number, a date, time or datetime its ISO 8601 text, text as text, a
    logical value a bool, an error its code, a blank ``None``.  Nothing is coerced: a
    number stored as text (``number_as_text``) stays text, a formula without a cached
    value stays blank.  Each column's type label is its values' one type, else ``None``:
    a deck or document tool writes a ``date`` (``datetime``, ``time``) column's ISO text in
    the session's locale, as a Python date (ooxml-edit's ``Rows.label``).

    A declared table's cleaned values (:mod:`xlsx_agent.clean`) are placed cleaned, a list
    as its items joined by :data:`LIST_JOIN` (a line break: deck and document tables show
    one cell, a line per item; its type label ``text``).  When any value was cleaned, the
    dataset's ``detail`` is ``{"cleaned": Count}`` (how many, the first cells) and a
    :class:`~xlsx_agent.clean.CleanedWarning` says so -- in a tool session, among the
    deck's or document's tool's warnings.

    Each value's source is its cell by key (``t:Products/SKU-0042/Price``) when the table
    has a key column -- declared, or the one found -- and every data row has a key no other
    row shares, as a derive's sources are (:mod:`xlsx_agent.derive`); else its cell address
    (``Stock!F7``).  ``keys`` is 1 when the key column comes first.  More than
    ``max_cells`` values (rows times columns) is refused before any is read:
    :class:`RegionError` ``too_many_cells``.

    ``cache`` counts the formula cells read by their cache status and ``external``
    (ooxml-edit's ``Dataset.cache``): rows read from an invalidated cache are placed
    ``stale``, and the placed values' ``from`` carries the counts."""
    target = parse_address(address) if isinstance(address, str) else address
    if not isinstance(target, TableAddress):
        raise AddressError("rows come from a table, name or region: t:Products, "
                           "t:Products[Price] or t:Products/{SKU-0003,SKU-0007}",
                           text=format_address(target))
    address = target
    selection = workbook.resolve(address)
    region = selection.region
    whole = region.address
    if selection.kind == "value":
        raise RegionError("not_rows", f"{selection.address} is one value, not rows: name "
                          "the table, a column or the rows",
                          address=selection.address,
                          valid_options=[whole.text,
                                         replace(whole, column=selection.column).text])
    columns = (selection.column,) if selection.column is not None else tuple(region.columns)
    left = region.area.first.column
    offsets = [region.column_index(name) for name in columns]
    data = region.data
    if selection.positions:
        positions = list(selection.positions)
    elif selection.row is not None:
        positions = [selection.row]
    elif data is not None and selection.range is not None:
        positions = list(range(1, region.rows + 1))
    else:
        positions = []
    cells = len(positions) * len(columns)
    if max_cells is not None and cells > max_cells:
        raise RegionError("too_many_cells", f"{selection.address} has {cells} values; at "
                          f"most {max_cells}: name a column or a key set, or derive a "
                          "smaller dataset", address=selection.address)
    found: dict[int, dict[int, Value]] = {}
    if positions:
        for row in workbook.read(address):
            for value in row:
                found.setdefault(value.row, {})[value.column] = value
    keys, index = None, None
    if positions and data is not None:
        try:
            index = workbook.regions.key_index(region)
        except (RegionError, KeyError):
            index = None
        if index is not None and index.unique:
            keys = index.keys
    rows: list[tuple[Any, ...]] = []
    sources: list[tuple[tuple[str, ...], ...]] = []
    seen: list[set[str]] = [set() for _ in columns]
    cleaned: list[str] = []
    cleaned_count = 0
    cache: dict[str, int] = {}
    for position in positions:
        sheet_row = data.first.row + position - 1
        cells_of = found.get(sheet_row, {})
        values: list[Any] = []
        where: list[tuple[str]] = []
        for n, (name, offset) in enumerate(zip(columns, offsets)):
            value = cells_of.get(left + offset)
            if value is not None and value.cache is not None:
                cache[value.cache] = cache.get(value.cache, 0) + 1
                if value.external:
                    cache["external"] = cache.get("external", 0) + 1
            if value is not None and value.cleaned:
                cleaned_count += 1
                if len(cleaned) < FIRST_ADDRESSES:
                    cleaned.append(value.a)
            if value is None or value.t == "blank":
                values.append(None)
            elif value.t == "list":           # a declared split: one cell, a line per item
                values.append(LIST_JOIN.join(value.v))
                seen[n].add("text")
            else:
                values.append(value.v)
                seen[n].add(value.t)
            if keys is not None:
                where.append((replace(whole, key=keys[position - 1], column=name).text,))
            elif value is not None:
                where.append((value.a,))      # its own cell, a declared table's too
            else:
                cell = Cell(sheet_row, left + offset)
                where.append((region.located(Area(cell, cell)) or "",))
        rows.append(tuple(values))
        sources.append(tuple(where))
    types = tuple(next(iter(t)) if len(t) == 1 else None for t in seen)
    first_is_key = keys is not None and index is not None and columns[0] == index.column
    detail = None
    if cleaned_count:
        detail = {"cleaned": Count(cleaned_count, tuple(cleaned))}
        warnings.warn(CleanedWarning(
            f"{whole.text}: {cleaned_count} values were cleaned by the table's declared "
            f"rules, the first at {', '.join(cleaned)}; a list is placed as its items, one "
            "line each"), stacklevel=2)
    return Dataset(columns, tuple(rows), sources=tuple(sources), keys=1 if first_is_key else 0,

                   types=types, detail=detail, cache=cache or None)


def _typed(v: Any, label: str | None) -> tuple[Any, str, str | None]:
    """A dataset value as a cell's would be read: its value, its type and its issue
    (:func:`_type`).  A Python ``datetime``, ``date`` or ``time`` is that type without a
    label, its value ISO 8601 text as a cell's (``2026-03-01``, ``14:05:00``,
    ``2026-03-01T14:05:00``; to the millisecond, the wall-clock time of an aware one).
    ISO 8601 text in an unlabelled column stays text, flagged ``date_as_text`` as a cell's
    would be: ``coerce=["date_as_text"]`` reads it as a date, as a ``date`` label does."""
    if isinstance(v, _dt.datetime):
        ms = round(v.microsecond / 1000)
        day = v.replace(tzinfo=None, microsecond=0) + _dt.timedelta(milliseconds=ms)
        clock = ((day.hour * 60 + day.minute) * 60 + day.second) * 1000 \
            + day.microsecond // 1000
        return f"{day.date().isoformat()}T{_clock(clock)}", "datetime", None
    if isinstance(v, _dt.date):
        return v.isoformat(), "date", None
    if isinstance(v, _dt.time):
        clock = ((v.hour * 60 + v.minute) * 60 + v.second) * 1000 + round(v.microsecond / 1000)
        return _clock(min(clock, _MS_PER_DAY - 1)), "time", None
    t = _type(v, label)
    issue = None
    if t == "text" and label is None and isinstance(v, str) and looks_like_date(v):
        issue = "date_as_text"
    return v, t, issue


def _type(v: Any, label: str | None) -> str:
    """A dataset value's type: its column's label when the value can be one, else its
    own (a bool, a number, text, nothing)."""
    if v is None:
        return "blank"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, (int, float)):
        return label if label in ("number", "percent") else "number"
    if isinstance(v, str):
        return label if label in _MOMENT_LABELS + ("text", "error") else "text"
    return "text"


_AREAS: dict[str, Any] = {}


def _by_keys(column: TableAddress, keys: list[str]) -> TableAddress:
    """``column`` (a table's column address, ``t:Sales[Amount]``) as the values of
    ``keys``: one key, or a key set."""
    if len(keys) == 1:
        return TableAddress(name=column.name, region=column.region, scope=column.scope,
                            key=keys[0], column=column.column)
    return TableAddress(name=column.name, region=column.region, scope=column.scope,
                        keys=tuple(keys), column=column.column)


def _column_area(text: str) -> tuple | None:
    """``text`` as a sheet, a column and its first and last rows, when it names cells of
    one column (``Sales!C2:C9``); as a table's column (``t:Sales[Amount]``) and the keys
    of its rows, when it names values by key (``t:Sales/{S-01,S-02}/Amount``); ``None``
    otherwise."""
    if text in _AREAS:
        return _AREAS[text]
    found: tuple | None = None
    try:
        address: "Address | None" = parse_address(text)

    except (AddressError, ValueError):
        address = None
    if isinstance(address, RangeAddress) and address.span == "cells" \
            and address.area.first.column == address.area.last.column:
        area = address.area
        found = (address.sheet, area.first.column, area.first.row, area.last.row)
    elif isinstance(address, TableAddress) and address.column is not None \
            and (address.key is not None or address.keys is not None):
        found = (TableAddress(name=address.name, region=address.region,
                              scope=address.scope, column=address.column),
                 address.keys if address.keys is not None else (address.key,))
    if len(_AREAS) < 4096:
        _AREAS[text] = found
    return found
