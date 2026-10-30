"""Joins: another table's (or dataset's) columns added to each row of a table, by key, each
added value with the cells it came from (round 25, X1; :mod:`xlsx_agent.derive`'s
``join``)::

    ds = wb.derive("t:Items", join={"from": "t:Issues", "on": ["Ref", "Items"],
                                     "columns": ["Issue", "Raised"]}, id="items_issues")
    ds.rows[0]        # {"Ref": "I-1", ..., "Issue": ("Q-2", "Q-7"), "Raised": (...)}
    ds.sources[...]   # each joined value: the cells of the joined rows it came from
    ds.joined         # {"from": "t:Issues", "on": ["Ref", "Items"], "rows": 12, ...}

**Rows.**  ``from`` (the derive's) is the left side: a table, name or region (``t:``) or
a dataset; ``join["from"]`` the right side, either kind too.  Every left row is kept, in
its order, once: the joined columns of a row no right row matches are blank.  ``on`` is
``[left column, right column]``.

**Matching.**  A right row joins a left row when its key cell equals the left row's, by
type and value as ``group_by`` groups (text exactly, case counting; ``3`` and ``3.0`` one
number, ``3`` and ``"3"`` two keys; a date by its day).  A blank key, an error and a
formula without a cached value match nothing.  A **list** key -- a declared column whose
text is split (:mod:`xlsx_agent.clean`), or a list in a dataset -- matches by any of its
items, on either side: a right row whose ``Items`` cell lists ``("I-1", "I-3")`` joins
both ``I-1`` and ``I-3``.  Text holding several keys that no declaration splits is one
key: declare the split, so that the table says how its text is cut.

**One to many: a list cell.**  When any left row is joined by more than one right row, the
joined columns are **list** columns: each matched cell holds the joined rows' values in the
right side's row order (their blanks left out, a list value's items in place), even a cell
with one match, so the column has one type.  Rows are never repeated: one row per left
row keeps the left table's counts, sums and keys what they were (a measure of a left
column counts each row once, and a row by key still names one row), and "the issues of
each item" is one cell of the item's row.  A list column takes ``count``, and ``where``'s
``eq``, ``in`` and ``contains`` (a row passing when any item does), as a declared list
column does; a deck's or document's table shows it one item per line.  When every left row
is joined by one right row at most, the joined cells hold the right values as they are.

**Columns.**  ``columns`` names the right side's columns to add, in order (default: every
one but its key column).  A name the left side already has (case aside) is renamed
``"<name> (<right side>)"`` -- ``Status (Issues)`` -- and :attr:`Joined.columns` maps each
added column to the right column it holds.

**Provenance.**  A left value's sources are its own cell (by key when the left table has a
unique key column, as a derive writes them); a joined value's are the cells of the right
rows it came from, in that column -- by key, or as cells -- every one of them, merged as a
derive merges sources.  A blank (unmatched) joined cell has none; the row's key cell names
it where a cell address is needed (``excluded``).

**Then the pipeline.**  The join comes first: ``where`` filters the joined rows (a joined
column too), then ``group_by``, ``bins`` and ``pivot`` group them, the ``measures`` are
computed, ``sort`` and ``top`` apply to the result, ``round`` last.  Without ``group_by``,
``measures``, ``bins`` or ``pivot``, the result is the joined rows themselves (``where``,
``sort`` and ``top`` apply; ``round`` is refused: values are cells as stored): the dataset
an agent places as a table, one row per left row, at most
:data:`~xlsx_agent.derive.MAX_ROWS` (``top`` keeps the first).

**Bounded.**  Both sides are read once, whole; at most :data:`MAX_JOINED` values are held
(rows times columns, of both), else :class:`~xlsx_agent.derive.DeriveError`
``too_many_cells``.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any, Mapping, Sequence, cast

from ooxml_edit.sheets import Area, Cell
from ooxml_edit.tools import Dataset

from .address import RangeAddress
from .dataset import DatasetTable
from .read.regions import Region, RegionError
from .read.types import Value

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Workbook

#: The most values (rows times columns, both sides) a join holds.
MAX_JOINED = 1_000_000

#: What a join takes.
JOIN_KEYS = ("from", "on", "columns")


@dataclass(slots=True, repr=False)
class JoinedValue(Value):
    """A value of a joined table: its type and value as read; :attr:`a` the cell it is
    (a left value's own), or for a joined value the first cell it came from (none: the
    left row's key cell)."""

    at: str = ""

    @property
    def a(self) -> str:
        return self.at

    def __repr__(self) -> str:
        return f"JoinedValue({self.at}, t={self.t!r})"


@dataclass(frozen=True)
class Joined:
    """What a join did: ``from`` the right side's address, ``on`` the two key columns,
    ``columns`` each added column's name and the right column it holds, ``matched`` how
    many left rows were joined by a right row, ``unmatched`` how many by none, ``lists``
    whether the added columns are lists (a left row joined by several right rows)."""

    source: str
    on: tuple[str, str]
    columns: dict[str, str]
    matched: int
    unmatched: int
    lists: bool

    def as_dict(self) -> dict:
        return {"from": self.source, "on": list(self.on), "columns": dict(self.columns),
                "matched": self.matched, "unmatched": self.unmatched, "lists": self.lists}


class JoinedTable(DatasetTable):
    """The left side's rows with the right side's columns added: what the derive engine
    reads, as it reads a dataset (:class:`~xlsx_agent.dataset.DatasetTable`), each value as
    read (a :class:`JoinedValue`, with its position as its row) and each value's sources.
    ``address`` is the left side's; ``joined`` what the join did."""

    def __init__(self, columns: Sequence[str], values: list[list[JoinedValue]],
                 sources: list[list[tuple[str, ...]]], address: Any, keys: int,
                 date1904: bool, joined: Joined) -> None:
        dataset = Dataset(tuple(columns), tuple(tuple(v.v for v in row) for row in values),
                          sources=tuple(tuple(row) for row in sources))
        super().__init__(dataset, "joined")
        self.address = address
        self.date1904 = date1904
        self.derived = True        # its types are its cells', not a caller's labels
        self.cells = values
        self.left_keys = keys
        self.joined = joined

    def value(self, position: int, offset: int, *, measured: bool = False,
              sources: int | None = None) -> Any:
        return self.cells[position - 1][offset]

    def measured_rows(self) -> Any:
        return iter(self.cells)


def joined_table(workbook: "Workbook | None", left: "Region | DatasetTable",
                 right: "Region | DatasetTable", join: Mapping[str, Any]) -> JoinedTable:
    """``left``'s rows with ``right``'s columns (see the module's notes); ``join`` the
    caller's ``{"from", "on", "columns"}`` (``from`` already resolved to ``right``)."""
    from .derive import DeriveError

    on = join.get("on")
    if isinstance(on, (str, bytes)) or not isinstance(on, Sequence) or len(on) != 2 \
            or not all(isinstance(n, str) for n in on):
        raise DeriveError("bad_join", "join's on is [this table's key column, the joined "
                          "one's]: two column names", field="join")
    left_key = left.columns[left.column_index(on[0])]
    right_key = right.columns[right.column_index(on[1])]
    wanted = join.get("columns")
    if wanted is None:
        picked = [c for c in right.columns if c != right_key]
    elif isinstance(wanted, (str, bytes)) or not isinstance(wanted, Sequence) or not wanted \
            or not all(isinstance(n, str) for n in wanted):
        raise DeriveError("bad_join", "join's columns is a list of the joined table's "
                          "columns to add", field="join")
    else:
        picked = list(dict.fromkeys(right.columns[right.column_index(n)] for n in wanted))
    if not picked:
        raise DeriveError("bad_join", "the joined table has no column to add but its key: "
                          "name columns", field="join")
    right_name = _name(right)
    taken = {c.casefold() for c in left.columns}
    names: dict[str, str] = {}
    for column in picked:
        name, n = column, 2
        if name.casefold() in taken:
            name = base = f"{column} ({right_name})"
            while name.casefold() in taken:
                name, n = f"{base} {n}", n + 1
        taken.add(name.casefold())
        names[name] = column

    left_rows, left_sources = _read(workbook, left)
    right_rows, right_sources = _read(workbook, right)
    width = len(left.columns) + len(picked)
    held = len(left_rows) * width + len(right_rows) * len(right.columns)
    if held > MAX_JOINED:
        raise DeriveError("too_many_cells", f"a join of {held} values, more than "
                          f"{MAX_JOINED}: join smaller tables (a derived dataset)",
                          field="join")
    date1904 = _date1904(workbook, left)
    lk, rk = left.column_index(left_key), right.column_index(right_key)
    offsets = [right.column_index(c) for c in picked]
    index: dict[Any, list[int]] = {}
    for position, row in enumerate(right_rows):
        for key in _keys(row[rk], date1904):
            found = index.setdefault(key, [])
            if not found or found[-1] != position:
                found.append(position)
    matches: list[list[int]] = []
    for row in left_rows:
        found_rows: set[int] = set()
        for key in _keys(row[lk], date1904):
            found_rows.update(index.get(key, ()))
        matches.append(sorted(found_rows))
    lists = any(len(m) > 1 for m in matches)

    values: list[list[JoinedValue]] = []
    sources: list[list[tuple[str, ...]]] = []
    for position, (row, own, hit) in enumerate(zip(left_rows, left_sources, matches), 1):
        out = [_moved(value, position, n) for n, value in enumerate(row)]
        where = list(own)
        key_cell = row[lk].a
        for k, offset in enumerate(offsets, len(row)):
            cells = [right_rows[m][offset] for m in hit]
            cited = tuple(text for m in hit for text in right_sources[m][offset])
            out.append(_joined(cells, lists, position, k, cited[0] if cited else key_cell))
            where.append(cited)
        values.append(out)
        sources.append(where)
    joined = Joined(_text(right), (left_key, right_key), names,
                    sum(1 for m in matches if m), sum(1 for m in matches if not m), lists)
    keys = 1 if isinstance(left, Region) and _keyed(workbook, left) is not None \
        and left.columns[0] == _keyed(workbook, left) else 0
    return JoinedTable([*left.columns, *names], values, sources, left.address, keys,
                       date1904, joined)


def _name(side: "Region | DatasetTable") -> str:
    if isinstance(side, DatasetTable):
        return side.id
    return side.name or side.address.text


def _text(side: "Region | DatasetTable") -> str:
    return side.address.text


def _date1904(workbook: "Workbook | None", left: "Region | DatasetTable") -> bool:
    if isinstance(left, DatasetTable):
        return left.date1904
    assert workbook is not None   # a table is read in its workbook
    return workbook.date1904


def _keyed(workbook: "Workbook | None", region: Region) -> str | None:
    """The key column of ``region`` when every data row has a key no other row shares."""
    if workbook is None or region.data is None:
        return None
    try:
        index = workbook.regions.key_index(region)
    except (RegionError, KeyError):
        return None
    return index.column if index.unique else None


def _read(workbook: "Workbook | None", side: "Region | DatasetTable"
          ) -> tuple[list[list[Value]], list[list[tuple[str, ...]]]]:
    """Every row of ``side``, every column, as read, and each value's sources: a table's
    cell by key when its keys are unique, else its cell; a dataset's value's own."""
    rows: list[list[Value]]
    sources: list[list[tuple[str, ...]]] = []
    if isinstance(side, DatasetTable):
        rows = [list(row) for row in side.measured_rows()]
        return rows, [[side.cell_sources(p, o) for o in range(len(row))]
                      for p, row in enumerate(rows, 1)]
    data = side.data
    if data is None:
        return [], []
    assert workbook is not None   # a table is read in its workbook
    area = Area(Cell(data.first.row, side.area.first.column),
                Cell(data.last.row, side.area.last.column))
    rows = [list(row) for row in workbook._read(RangeAddress(side.sheet_ref, area))]
    keys = None
    if _keyed(workbook, side) is not None:
        keys = workbook.regions.key_index(side).keys
    whole = side.address
    for p, row in enumerate(rows):
        if keys is not None:
            sources.append([(replace(whole, key=keys[p], column=name).text,)
                            for name in side.columns])
        else:
            sources.append([(value.a,) for value in row])
    return rows, sources


def _keys(value: Value, date1904: bool) -> list[tuple]:
    """What a key cell matches by: its type and value as ``group_by`` groups it, or each
    item's of a list; nothing for a blank, an error or a formula without a cached value."""
    from .derive import _group_key

    if value.t in ("blank", "error") or value.issue == "no_cached_value":
        return []
    if value.t == "list":
        return [_item_key(item) for item in value.v]   # type: ignore[attr-defined]
    return [_group_key(value, date1904)]


def _item_key(item: Any) -> tuple:
    if isinstance(item, bool):
        return ("bool", item)
    if isinstance(item, (int, float)):
        return ("number", item)
    return ("text", item)


def _moved(value: Value, position: int, offset: int) -> JoinedValue:
    """A left value as a joined table's: at its position, named by its own cell."""
    return JoinedValue(sheet="", row=position, column=offset + 1, v=value.v, t=value.t,
                       formula=value.formula, issue=value.issue, hidden=value.hidden,
                       cache=value.cache, external=value.external, cleaned=value.cleaned,
                       at=value.a)


def _joined(cells: list[Value], lists: bool, position: int, offset: int,
            at: str) -> JoinedValue:
    """The joined value of a left row's matched right cells: a list of their values
    (blanks left out, a list's items in place) when the column is one, else the one
    cell's value; blank for none."""
    if not lists:
        if not cells:
            return JoinedValue(sheet="", row=position, column=offset + 1, at=at)
        value = cells[0]
        return JoinedValue(sheet="", row=position, column=offset + 1, v=value.v, t=value.t,
                           formula=value.formula, issue=value.issue, hidden=value.hidden,
                           cache=value.cache, external=value.external,
                           cleaned=value.cleaned, at=at)
    items: list[Any] = []
    for value in cells:
        if value.t == "list":
            items.extend(cast(tuple, value.v))
        elif value.t != "blank":
            items.append(value.v)
    if not items:
        return JoinedValue(sheet="", row=position, column=offset + 1, at=at)
    return JoinedValue(sheet="", row=position, column=offset + 1, v=tuple(items), t="list",
                       hidden=any(v.hidden for v in cells), at=at)


__all__ = ["JOIN_KEYS", "MAX_JOINED", "Joined", "JoinedTable", "JoinedValue", "joined_table"]
