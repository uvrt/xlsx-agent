"""Comparing two workbooks: sheets by id, tables and regions by name, rows by key, cells by
type and value (ROADMAP.md, "Tools": ``xlsx_compare``)::

    diff = Workbook.compare(old, new, tables=["t:Products"], key="SKU")
    diff.sheets          # [SheetChange(change='renamed', id=3, a='Stock', b='Stock 2026')]
    diff.tables[0].changed   # [CellChange(change='changed', key='SKU-1', column='Price', ...,
                             #             at='t:Products/SKU-1/Price')]
    diff.tables[0].by_column()   # (ColumnSummary(column='Price', changed=3, up=2, down=1,
                                 #   max_up='t:Products/SKU-1/Price +0.5', ...),)
    diff.totals()        # {"sheets": 1, "rows_added": 2, "cells_changed": 3, ...}

``a`` is the older workbook, ``b`` the newer: what is *added* is in ``b`` only, what is
*removed* in ``a`` only.

**Sheets** are matched by ``sheetId`` (what ``#3!`` names: it survives a rename), a sheet
without one by its name: a sheet of one id with another name is ``renamed``; an id in one
workbook only is ``added`` or ``removed``.  Two versions of one workbook keep their ids;
two unrelated workbooks give whatever their ids give.

**Tables and regions** are matched by name, case aside, as ``t:`` addresses are
(:mod:`xlsx_agent.read.regions`).  ``tables`` names them: ``"t:Products"`` for the same
name in both, ``("t:Products", "t:Items")`` for a pair, a detected region by its header's
first cell (``t:Stock!A1``); a name one workbook lacks is
:class:`~xlsx_agent.address.NameNotFound`.  Without ``tables``, every Excel table and
declared region is compared with the one of its name in the other workbook, and those in
one only are listed (``tables_added``, ``tables_removed``); defined names and detected
regions are compared only when named.

**Columns** are matched by name, case aside.  A column whose header changed where it stands
-- the same position in the region, its old name not in ``b`` and its new name not in ``a``
-- is ``renamed``, and its cells are compared as one column's; any other column in one
only is ``added`` or ``removed``, and its cells are not compared.  ``columns`` limits the
cells compared to some columns (by either name); every column change is listed regardless.

**Rows** are matched by key: the key column named (``key``: a column name, or one per
table, by ``a``'s address), else the region's declared key, else the one column both
workbooks find a key (:meth:`~xlsx_agent.read.regions.Regions.key_columns`: unique, never
blank, text or whole numbers); with none, or several, the comparison is refused
(:class:`~xlsx_agent.read.regions.RegionError` ``no_key_column``, the candidates as
``valid_options``), as a key lookup is.  A key is a cell's text or a whole number's digits,
matched exactly, as :func:`~xlsx_agent.read.regions.key_text` reads it.  **A key two rows
of one workbook share is never paired by guess** (never the first of them): it is
reported (``duplicate_key``, with each workbook's rows by position) and its rows are left
out of every other count.  So are rows whose key cell holds no key (a blank, a fraction, a
date: ``unkeyed``, counted with their first positions).  A key in ``b`` only is a row
``added``, in ``a`` only a row ``removed``, each with its cells.

**Cells** of a paired row are compared by type, then value, as the reader types them
(:mod:`xlsx_agent.read.types`): ``changed`` (same type, another value), ``type`` (another
type: a number and the same number stored as text are a ``type`` change, reported as
such, never equal), ``added`` (blank in ``a``) and ``removed`` (blank in ``b``).  Text is
compared exactly (case counts); a date as its ISO text; a formula by its cached value (a
formula without one is ``blank``, as the reader gives it, and equal to a blank cell).
Each change carries both cells, typed, with their addresses.

**By key and by column.**  Each cell changed is named by key as ``b`` names it,
``t:Products/SKU-2/Price`` (``CellChange.at``), with its sheet address in each workbook
on the values; each row added or removed by key in its workbook, ``t:Products/SKU-6``
(``RowChange.at``).  :meth:`TableDiff.by_column` counts the cells changed per column and,
for numbers, percents, dates and times, how many went up and down and the largest move
each way, by key (:class:`ColumnSummary`): what a model reads before, or instead of, the
changes one by one.  A column hidden from the model (:mod:`xlsx_agent.declare`) is not a
column of its table, so it is never compared, counted or named.

**Labels.**  A key names a row exactly but tells a reader little (``SKU-0005``), so each
row added or removed and each cell changed also carries the row's **label**: its text in
the table's label column (:attr:`TableDiff.label`), the first column other than the key,
in ``a``'s order and in both workbooks, whose visible cells are mostly text (more text
than all other types together, over both workbooks; blanks and cells on a hidden row or
column not counted) -- whether or not ``columns`` compares it.  A cell changed and a row
added take the label from ``b``, a row removed from ``a``; at most :data:`LABEL_CHARS`
characters, longer text cut with ``…``.  A row whose label cell is blank, not text or
hidden (on a hidden row or column) has none, and a column hidden from the model is no
column, so never the label column.  Declarations name no label column of their own.

**Content** (ROADMAP.md, "Design rules", 7).  Keys and values are cell content: they are in
the :class:`Diff`, a result, and never in a message; errors name tables, columns and
codes only.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field, replace
from decimal import Decimal
from typing import TYPE_CHECKING, Any, Iterator, Mapping, Sequence

from ooxml_edit.sheets import Area, Cell

from .address import NameNotFound, RangeAddress, TableAddress, nearest
from .profile import FIRST_ADDRESSES, Count
from .read.regions import Region, RegionError, key_text
from .read.types import Value

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: The changes a cell of a paired row can show.
CELL_CHANGES = ("changed", "type", "added", "removed")

#: The most characters of a row's label (see the module's notes); longer text is cut with ``…``.
LABEL_CHARS = 60


@dataclass(frozen=True)
class SheetChange:
    """A sheet ``added`` (in ``b`` only), ``removed`` (in ``a`` only) or ``renamed``:
    ``id`` its ``sheetId`` (``None`` for a sheet without one, matched by name), ``a`` and
    ``b`` its name in each (``None`` where it is not)."""

    change: str
    id: int | None
    a: str | None
    b: str | None

    def as_dict(self) -> dict:
        out: dict[str, Any] = {"change": self.change}
        if self.id is not None:
            out["id"] = self.id
        if self.a is not None:
            out["a"] = self.a
        if self.b is not None:
            out["b"] = self.b
        return out


@dataclass(frozen=True)
class ColumnChange:
    """A column ``added``, ``removed`` or ``renamed``: its name in each workbook."""

    change: str
    a: str | None
    b: str | None

    def as_dict(self) -> dict:
        out: dict[str, Any] = {"change": self.change}
        if self.a is not None:
            out["a"] = self.a
        if self.b is not None:
            out["b"] = self.b
        return out


@dataclass(frozen=True)
class RowChange:
    """A row ``added`` (its key in ``b`` only) or ``removed`` (in ``a`` only): ``key``,
    ``row`` its address by position in its workbook (``t:Products/#7``), ``cells`` its
    compared columns' cells, typed; ``at`` its address by key in its workbook
    (``t:Products/SKU-7``); ``label`` its text in the table's label column
    (:attr:`TableDiff.label`), ``None`` where it has none."""

    change: str
    key: str
    row: str
    cells: tuple[Value, ...]
    at: str = ""
    label: str | None = None


@dataclass(frozen=True)
class CellChange:
    """A cell of a paired row that differs: ``change`` one of :data:`CELL_CHANGES`;
    ``key`` the row's; ``column`` the column's name in ``b``; ``a`` and ``b`` the cell in
    each, typed with its address (``blank`` where it holds nothing); ``at`` the cell's
    address by key in ``b`` (``t:Products/SKU-2/Price``); ``label`` the row's text in the
    table's label column in ``b`` (:attr:`TableDiff.label`), ``None`` where it has none."""

    change: str
    key: str
    column: str
    a: Value
    b: Value
    at: str = ""
    label: str | None = None

    @property
    def delta(self) -> Decimal | None:
        """How far the value moved, for a ``changed`` number, percent, date or time: ``b``
        less ``a``, in its own unit for a number or percent (exact, as the shortest decimals
        of the two: 0.09 - 0.08 is 0.01), days for a date or date and time, seconds for a
        time; ``None`` for any other change."""
        return _delta(self.a, self.b) if self.change == "changed" else None


@dataclass(frozen=True)
class ColumnSummary:
    """One column's changed cells in a :class:`TableDiff`, counted (:meth:`TableDiff.by_column`):
    ``changed`` every kind (:data:`CELL_CHANGES`); for numbers, percents, dates and times
    changed in value, ``up`` and ``down``, and ``max_up`` and ``max_down`` the largest move
    each way, as the cell's address by key and the move signed (``t:Checks/A3/Rating +1.5``,
    a date's in days: ``+3d``, a date and time's in days to four places, a time's as
    ``+1:30`` or ``-0:00:20``); the first met wins a tie.  Text, booleans, errors and lists
    count as ``changed`` only."""

    column: str
    changed: int
    up: int = 0
    down: int = 0
    max_up: str | None = None
    max_down: str | None = None

    def as_dict(self) -> dict:
        out: dict[str, Any] = {"changed": self.changed}
        for name in ("up", "down", "max_up", "max_down"):
            value = getattr(self, name)
            if value:
                out[name] = value
        return out


@dataclass(frozen=True)
class DuplicateKey:
    """A key that two or more rows of one workbook share: never paired.  ``a`` and ``b``
    the rows holding it in each, by position (``t:Products/#3``)."""

    key: str
    a: tuple[str, ...]
    b: tuple[str, ...]


@dataclass(frozen=True)
class TableDiff:
    """One table or region compared: ``a`` and ``b`` its address in each, ``range_a`` and
    ``range_b`` where it lies, ``key`` the key column (as ``a`` names it; ``b`` may have
    renamed it), ``rows_a`` and ``rows_b`` its data rows, ``paired`` the rows matched by
    key and ``unchanged`` those of them with no cell changed; the changes, each in the
    order they are met (``changed`` and ``added`` in ``b``'s row order, ``removed`` in
    ``a``'s); ``compared`` the columns whose cells were compared, as ``b`` names them;
    ``label`` the column whose text names a row for a reader, as ``b`` names it (``None``
    for a table without one: see the module's notes)."""

    a: str
    b: str
    range_a: str
    range_b: str
    key: str
    rows_a: int
    rows_b: int
    paired: int
    unchanged: int
    columns: tuple[ColumnChange, ...] = ()
    duplicates: tuple[DuplicateKey, ...] = ()
    unkeyed_a: Count = field(default_factory=lambda: Count(0))
    unkeyed_b: Count = field(default_factory=lambda: Count(0))
    added: tuple[RowChange, ...] = ()
    removed: tuple[RowChange, ...] = ()
    changed: tuple[CellChange, ...] = ()
    compared: tuple[str, ...] = ()
    label: str | None = None

    def by_column(self) -> tuple[ColumnSummary, ...]:
        """Each compared column with a cell changed (:class:`ColumnSummary`), in ``b``'s
        column order: rows added and removed are not in it (they are by key in ``added``
        and ``removed``)."""
        changes: dict[str, list[CellChange]] = {}
        for change in self.changed:
            changes.setdefault(change.column, []).append(change)
        order = [*self.compared, *[name for name in changes if name not in self.compared]]
        return tuple(_summarise(name, changes[name]) for name in order if name in changes)

    def totals(self) -> dict[str, int]:
        """The changes counted by kind (zeros left out)."""
        out = {"columns": len(self.columns), "duplicate_key": len(self.duplicates),
               "unkeyed": self.unkeyed_a.count + self.unkeyed_b.count,
               "rows_added": len(self.added), "rows_removed": len(self.removed),
               "cells_changed": len(self.changed)}
        return {name: n for name, n in out.items() if n}


@dataclass(frozen=True)
class Diff:
    """Two workbooks compared (see the module's notes): ``sheets`` the sheet changes;
    ``tables`` each table or region compared; ``tables_added`` and ``tables_removed``
    those in one workbook only (when ``tables`` was not given), by address."""

    sheets: tuple[SheetChange, ...]
    tables: tuple[TableDiff, ...]
    tables_added: tuple[str, ...] = ()
    tables_removed: tuple[str, ...] = ()

    def totals(self) -> dict[str, int]:
        """Every change counted by kind, over every table (zeros left out)."""
        out = {"sheets": len(self.sheets),
               "unmatched_tables": len(self.tables_added) + len(self.tables_removed)}
        for table in self.tables:
            for name, n in table.totals().items():
                out[name] = out.get(name, 0) + n
        return {name: n for name, n in out.items() if n}

    @property
    def same(self) -> bool:
        """Nothing differs."""
        return not self.totals()


def compare(a: "Workbook", b: "Workbook", *,
            tables: "str | Sequence[str | tuple[str, str]] | None" = None,
            key: "str | Mapping[str, str] | None" = None,
            columns: "Sequence[str] | None" = None) -> Diff:
    """``a`` (the older) against ``b`` (the newer); see the module's notes."""
    if isinstance(columns, str):
        columns = [columns]
    pairs, added, removed = _pairs(a, b, tables)
    diffs = []
    for region_a, region_b in pairs:
        wanted = _key_for(key, region_a)
        diffs.append(_table(a, b, region_a, region_b, wanted, columns))
    return Diff(sheets=tuple(_sheets(a.sheets, b.sheets)), tables=tuple(diffs),
                tables_added=tuple(added), tables_removed=tuple(removed))


# -- sheets ------------------------------------------------------------------------------------


def _sheets(old: "list[Sheet]", new: "list[Sheet]") -> Iterator[SheetChange]:
    def ident(sheet: "Sheet") -> tuple:
        return ("id", sheet.id) if sheet.id is not None else ("name", sheet.name.casefold())

    later = {ident(sheet): sheet for sheet in new}
    earlier = {ident(sheet): sheet for sheet in old}
    for sheet in old:
        other = later.get(ident(sheet))
        if other is None:
            yield SheetChange("removed", sheet.id, sheet.name, None)
        elif other.name != sheet.name:
            yield SheetChange("renamed", sheet.id, sheet.name, other.name)
    for sheet in new:
        if ident(sheet) not in earlier:
            yield SheetChange("added", sheet.id, None, sheet.name)


# -- which tables ------------------------------------------------------------------------------


def _pairs(a: "Workbook", b: "Workbook", tables: Any
           ) -> tuple[list[tuple[Region, Region]], list[str], list[str]]:
    if tables is None:
        old = {r.name.casefold(): r for r in [*a.tables, *a.regions.declared]}
        new = {r.name.casefold(): r for r in [*b.tables, *b.regions.declared]}
        pairs = [(region, new[name]) for name, region in old.items() if name in new]
        added = [r.address.text for name, r in new.items() if name not in old]
        removed = [r.address.text for name, r in old.items() if name not in new]
        return pairs, added, removed
    if isinstance(tables, (str, TableAddress)):
        tables = [tables]
    pairs = []
    for item in tables:
        if isinstance(item, (tuple, list)):
            if len(item) != 2:
                raise ValueError("a pair of tables is (the one in a, the one in b)")
            first, second = item
        else:
            first = second = item
        pairs.append((_region(a, first), _region(b, second)))
    return pairs, [], []


def _region(workbook: "Workbook", address: "str | TableAddress") -> Region:
    if isinstance(address, str):
        address = TableAddress.parse(address)
    if not isinstance(address, TableAddress) or not address.is_whole:
        raise ValueError("compare takes tables or regions (t:Products), not their columns "
                         "or rows")
    return workbook.region(address)


def _key_for(key: Any, region: Region) -> str | None:
    if key is None or isinstance(key, str):
        return key
    if isinstance(key, Mapping):
        folded = {str(name).casefold(): column for name, column in key.items()}
        for name in (region.address.text, region.name or ""):
            if name.casefold() in folded:
                return folded[name.casefold()]
        return None
    raise ValueError("key is a column name, or a mapping of tables to column names")


# -- one table ---------------------------------------------------------------------------------


def _match_columns(old: Sequence[str], new: Sequence[str]
                   ) -> tuple[list[tuple[int, int]], list[ColumnChange]]:
    """The columns paired (offset in ``old``, offset in ``new``), by name, case aside, and
    a renamed column by its place; and the changes, in ``new``'s order, then ``old``'s."""
    in_old = {name.casefold(): i for i, name in enumerate(old)}
    in_new = {name.casefold(): i for i, name in enumerate(new)}
    pairs: dict[int, int] = {}
    changes: list[ColumnChange] = []
    for i, name in enumerate(old):
        j = in_new.get(name.casefold())
        if j is not None:
            pairs[i] = j
        elif i < len(new) and new[i].casefold() not in in_old:
            pairs[i] = i
            changes.append(ColumnChange("renamed", name, new[i]))
    paired_new = set(pairs.values())
    changes.extend(ColumnChange("added", None, name) for j, name in enumerate(new)
                   if j not in paired_new)
    changes.extend(ColumnChange("removed", name, None) for i, name in enumerate(old)
                   if i not in pairs)
    return sorted(pairs.items()), changes


def _table(a: "Workbook", b: "Workbook", old: Region, new: Region, key: str | None,
           columns: Sequence[str] | None) -> TableDiff:
    paired, column_changes = _match_columns(old.columns, new.columns)
    by_old = dict(paired)
    by_new = {j: i for i, j in paired}

    def offset(name: str) -> int | None:
        """A column named as either workbook names it: its offset in ``old``, if paired."""
        folded = name.casefold()
        for i, title in enumerate(old.columns):
            if title.casefold() == folded:
                return i if i in by_old else None
        for j, title in enumerate(new.columns):
            if title.casefold() == folded:
                return by_new.get(j)
        raise NameNotFound(f"no column {name!r} in {old.address.text} or {new.address.text}",
                           what="column", name=name,
                           candidates=list(dict.fromkeys([*old.columns, *new.columns])),
                           valid_options=nearest(name, [*old.columns, *new.columns]))

    key_offset = _key_offset(a, b, old, new, key, offset, by_old)
    if columns:
        wanted = []
        for name in columns:
            i = offset(name)
            if i is None:
                both = [new.columns[j] for _, j in paired]
                raise NameNotFound(f"column {name!r} is not in both {old.address.text} and "
                                   f"{new.address.text}", what="column", name=name,
                                   candidates=both, valid_options=nearest(name, both))
            wanted.append(i)
        compared = list(dict.fromkeys(wanted))
    else:
        compared = [i for i, _ in paired]
    compared = [i for i in compared if i != key_offset]
    # The label column's candidates are read too: every paired column but the key.
    width = 1 + len(compared)
    read = [key_offset, *compared,
            *[i for i, _ in paired if i != key_offset and i not in compared]]

    rows_old = _rows(a, old, read)
    rows_new = _rows(b, new, [by_old[o] for o in read])
    keys_old, unkeyed_old = _keyed(rows_old, old)
    keys_new, unkeyed_new = _keyed(rows_new, new)
    at_label = _label_column(read, keys_old, keys_new)

    duplicates = []
    for text in dict.fromkeys([*keys_old, *keys_new]):
        found_old, found_new = keys_old.get(text, []), keys_new.get(text, [])
        if len(found_old) > 1 or len(found_new) > 1:
            duplicates.append(DuplicateKey(text, tuple(_at(old, r[0]) for r in found_old),
                                           tuple(_at(new, r[0]) for r in found_new)))
    shared = {d.key for d in duplicates}

    titles = [new.columns[by_old[i]] for i in compared]
    offsets = [key_offset, *compared]
    added, changed = [], []
    paired_rows = unchanged = 0
    for text, found in keys_new.items():
        if text in shared:
            continue
        position, cells = found[0]
        earlier = keys_old.get(text)
        if earlier is None:
            added.append(RowChange("added", text, _at(new, position),
                                   _values(new, position, [by_old[o] for o in offsets],
                                           cells[:width]), _by_key(new, text),
                                   _label(cells, at_label)))
            continue
        paired_rows += 1
        old_position, old_cells = earlier[0]
        before = len(changed)
        label = _label(cells, at_label)
        for index, (one, two) in enumerate(zip(old_cells[1:width], cells[1:width])):
            change = _cell_change(one, two)
            if change is None:
                continue
            i = compared[index]
            changed.append(CellChange(
                change, text, titles[index],
                _value(old, old_position, i, one),
                _value(new, position, by_old[i], two),
                _by_key(new, text, titles[index]), label))
        if len(changed) == before:
            unchanged += 1
    removed = [RowChange("removed", text, _at(old, found[0][0]),
                         _values(old, found[0][0], offsets, found[0][1][:width]),
                         _by_key(old, text), _label(found[0][1], at_label))
               for text, found in keys_old.items() if text not in shared and text not in keys_new]
    return TableDiff(a=old.address.text, b=new.address.text, range_a=old.where,
                     range_b=new.where, key=old.columns[key_offset], rows_a=old.rows,
                     rows_b=new.rows, paired=paired_rows, unchanged=unchanged,
                     columns=tuple(column_changes), duplicates=tuple(duplicates),
                     unkeyed_a=unkeyed_old, unkeyed_b=unkeyed_new, added=tuple(added),
                     removed=tuple(removed), changed=tuple(changed), compared=tuple(titles),
                     label=None if at_label is None else new.columns[by_old[read[at_label]]])


def _label_column(read: list[int], *keyed: dict[str, list]) -> int | None:
    """Where the label column is among the cells ``read`` (the key's first): the first
    other column, in the table's order, whose visible cells over both workbooks' keyed
    rows are mostly text (see the module's notes); ``None`` without one."""
    for index in sorted(range(1, len(read)), key=lambda i: read[i]):
        text = other = 0
        for rows in keyed:
            for found in rows.values():
                for _, cells in found:
                    kept = cells[index]
                    if kept[1] == "blank" or kept[4]:
                        continue
                    if kept[1] == "text":
                        text += 1
                    else:
                        other += 1
        if text > other:
            return index
    return None


def _label(cells: tuple[_Kept, ...], index: int | None) -> str | None:
    """A row's label: its text in the label column, visible and not empty, cut at
    :data:`LABEL_CHARS` characters; ``None`` otherwise."""
    if index is None:
        return None
    v, t, _, _, hidden = cells[index][:5]
    if t != "text" or hidden or not isinstance(v, str) or not v.strip():
        return None
    return v if len(v) <= LABEL_CHARS else v[:LABEL_CHARS - 1] + "…"


def _key_offset(a: "Workbook", b: "Workbook", old: Region, new: Region, key: str | None,
                offset: Any, by_old: dict[int, int]) -> int:
    """The key column's offset in ``old``: named, declared, or the one both find."""
    if key is None and old.key is not None:
        key = old.key
    if key is None and new.key is not None:
        key = new.key
    if key is not None:
        i = offset(key)
        if i is None:
            raise RegionError("no_key_column", f"key column {key!r} is not in both "
                              f"{old.address.text} and {new.address.text}",
                              address=old.address.text,
                              valid_options=[old.columns[i] for i in by_old])
        return i
    found_old = set(a.key_columns(old))
    found_new = {name.casefold() for name in b.key_columns(new)}
    candidates = [i for i, name in enumerate(old.columns)
                  if name in found_old and i in by_old
                  and new.columns[by_old[i]].casefold() in found_new]
    if len(candidates) == 1:
        return candidates[0]
    either = list(dict.fromkeys([*a.key_columns(old), *b.key_columns(new)]))
    why = "no column is" if not candidates else f"{len(candidates)} columns are"
    raise RegionError("no_key_column", f"{old.address.text} and {new.address.text} have no key "
                      f"column named, and {why} a key in both (unique, never blank, text or "
                      "whole numbers); name one with key", address=old.address.text,
                      valid_options=[old.columns[i] for i in candidates] or either)


#: A cell as :func:`_rows` keeps it: its value, type, issue, whether it is a formula's
#: cached value and hidden, its own cell (a declared table's: :attr:`Value.origin`), and a
#: formula's cache status and whether it refers to another workbook.
_Kept = tuple


def _rows(workbook: "Workbook", region: Region, offsets: list[int]
          ) -> Iterator[tuple[int, tuple[_Kept, ...]]]:
    """Each data row's position (from 1) and the cells at ``offsets`` (the key's first),
    streamed once over the columns from the leftmost to the rightmost of them."""
    data = region.data
    if data is None:
        return
    first = region.area.first.column
    left, right = first + min(offsets), first + max(offsets)
    area = Area(Cell(data.first.row, left), Cell(data.last.row, right))
    reading = {first + o for o in offsets}
    for position, row in enumerate(workbook._read(RangeAddress(region.sheet_ref, area),
                                                  columns=reading), 1):
        yield position, tuple(_kept(row[first + o - left]) for o in offsets)


def _kept(value: Value) -> _Kept:
    return (value.v, value.t, value.issue, value.formula, value.hidden, value.origin,
            value.cache, value.external)


def _keyed(rows: Iterator[tuple[int, tuple[_Kept, ...]]], region: Region
           ) -> tuple[dict[str, list], Count]:
    """The rows by key, each key's rows in order; the rows without a key, counted."""
    out: dict[str, list] = {}
    unkeyed = 0
    first: list[str] = []
    for position, cells in rows:
        v, t = cells[0][:2]
        text = key_text(Value("", 0, 0, v=v, t=t))
        if text is None:
            unkeyed += 1
            if len(first) < FIRST_ADDRESSES:
                first.append(_at(region, position))
            continue
        out.setdefault(text, []).append((position, cells))
    return out, Count(unkeyed, tuple(first))


def _same(one: _Kept, two: _Kept) -> bool:
    return one[1] == two[1] and one[0] == two[0]


def _cell_change(one: _Kept, two: _Kept) -> str | None:
    if _same(one, two):
        return None
    if one[1] == "blank":
        return "added"
    if two[1] == "blank":
        return "removed"
    return "changed" if one[1] == two[1] else "type"


def _at(region: Region, position: int) -> str:
    """A data row's address by position: ``t:Products/#5``."""
    return replace(region.address, position=position).text


def _by_key(region: Region, key: str, column: str | None = None) -> str:
    """A row's address by key, or with ``column`` a cell's: ``t:Products/SKU-5/Price``."""
    return replace(region.address, key=key, column=column).text


def _value(region: Region, position: int, offset: int, kept: _Kept) -> Value:
    """A kept cell as the reader gives it, with its address."""
    v, t, issue, formula, hidden, origin, cache, external = kept
    return Value(_written(region), region.data.first.row + position - 1,
                 region.area.first.column + offset, v=v, t=t, issue=issue, formula=formula,
                 hidden=hidden, origin=origin, cache=cache, external=external)


def _values(region: Region, position: int, offsets: list[int],
            cells: tuple) -> tuple[Value, ...]:
    """A row's kept cells as values, at ``offsets`` in ``region``."""
    return tuple(_value(region, position, o, kept) for o, kept in zip(offsets, cells))


def _written(region: Region) -> str:
    """The region's sheet as an address writes it (``Stock``, ``'Q3 sales'``, ``#3``)."""
    return RangeAddress(region.sheet_ref, Area(Cell(1, 1), Cell(1, 1))).text.rpartition("!")[0]


# -- the summary by column -----------------------------------------------------------------------

#: The types whose values move up or down: a number's or percent's, a date's, a time's.
MOVING = ("number", "percent", "date", "datetime", "time")


def _delta(a: Value, b: Value) -> Decimal | None:
    """``b`` less ``a``, when both are of one :data:`MOVING` type (see
    :attr:`CellChange.delta`)."""
    if a.t != b.t or a.t not in MOVING or a.v is None or b.v is None:
        return None
    try:
        if a.t in ("number", "percent"):
            if isinstance(a.v, bool) or isinstance(b.v, bool):
                return None
            return Decimal(repr(b.v)) - Decimal(repr(a.v))
        if a.t == "date":
            return Decimal((dt.date.fromisoformat(b.v) - dt.date.fromisoformat(a.v)).days)
        if a.t == "datetime":
            moved = dt.datetime.fromisoformat(b.v) - dt.datetime.fromisoformat(a.v)
            return Decimal(moved.days) + Decimal(moved.seconds) / 86400 + \
                Decimal(moved.microseconds) / 86_400_000_000
        day = dt.date(2000, 1, 1)
        moved = dt.datetime.combine(day, dt.time.fromisoformat(b.v)) - \
            dt.datetime.combine(day, dt.time.fromisoformat(a.v))
        return Decimal(moved.days * 86400 + moved.seconds) + \
            Decimal(moved.microseconds) / 1_000_000
    except (TypeError, ValueError, ArithmeticError):
        return None


def _plain(number: Decimal) -> str:
    """A decimal written out in full, signed, without trailing zeros: ``+1.5``, ``-0.01``."""
    text = format(number, "+f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def _signed(delta: Decimal, kind: str) -> str:
    """A move as the summary writes it (:class:`ColumnSummary`)."""
    if kind == "date":
        return f"{_plain(delta)}d"
    if kind == "datetime":
        return f"{_plain(round(delta, 4))}d"
    if kind == "time":
        sign = "-" if delta < 0 else "+"
        whole = int(abs(delta))
        hours, rest = divmod(whole, 3600)
        minutes, seconds = divmod(rest, 60)
        return f"{sign}{hours}:{minutes:02d}" + (f":{seconds:02d}" if seconds else "")
    return _plain(delta)


def _summarise(column: str, changes: list[CellChange]) -> ColumnSummary:
    up = down = 0
    top: CellChange | None = None
    bottom: CellChange | None = None
    high = low = Decimal(0)
    for change in changes:
        delta = change.delta
        if delta is None or not delta:
            continue
        if delta > 0:
            up += 1
            if delta > high:
                high, top = delta, change
        else:
            down += 1
            if delta < low:
                low, bottom = delta, change
    return ColumnSummary(
        column, len(changes), up, down,
        f"{top.at} {_signed(high, top.b.t)}" if top is not None else None,
        f"{bottom.at} {_signed(low, bottom.b.t)}" if bottom is not None else None)
