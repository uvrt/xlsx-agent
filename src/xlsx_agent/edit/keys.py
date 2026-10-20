"""Cells written by key (0.2, X4; X6): ``t:stock/P-7/Quantity`` in :meth:`Workbook.set_cells
<xlsx_agent.workbook.Workbook.set_cells>`, and new rows in a declared table's capacity,
each in the row the host names.

::

    wb.declare_tables([{"name": "stock", "sheets": "Stock", "header_row": 9,
                        "key": "Item", "capacity": 200}])
    wb.set_cells({"t:stock/Kettle/Quantity": 4})          # the row whose Item is Kettle
    wb.set_cells({"t:stock/Lamp/Quantity": 2,             # Lamp is no row's key yet:
                  "t:stock/Lamp/Category": "Hardware"},   # a new row, in row 15, the
                 allocate={"t:stock/Lamp": 15})           # free row the host names

**By key, never by place.**  A value by key names the row whose key cell holds the key
(:mod:`xlsx_agent.read.regions`: text exactly, or a whole number's digits), wherever the
row is now: sorting the table, or a reviewer moving rows, does not move the write.  The
table (an Excel table, a declared table or region; any ``t:`` region) must have a key
column declared for it -- ``key`` in its declaration, ``declare_region(key=...)`` or
:meth:`Workbook.declare_key <xlsx_agent.workbook.Workbook.declare_key>` -- since a key
column guessed from the values is no address to write by (``no_key_column``).  Keys are
values the application supplies: a key column with a formula in any data row (an id made
from the row's number, ``="P-"&TEXT(ROW()-9,"000")``) is refused (``formula_key``), since
its keys are cached results that a sort, a row moved or a recalculation gives to other
rows; declare a key column of values instead.  A key two rows share refuses the cell
(``duplicate_key``); a key no row has refuses it (``missing_key``) unless the call
allocates it a row; a row by position (``t:stock/#3``), a key set, a row range, a wildcard
or a whole row or column is refused (``positional``): there is no fallback to a place.  A
key is cell content, so a refusal names such a value by its place in the call
(``cells[3]``: the fourth cell given; ``allocate[0]``: the first row allocated) and its
table, never by its key; a value found is named by its cell's address (``Stock!D12``).

**New rows** go only into a declared table's capacity (:mod:`xlsx_agent.declare`,
``capacity``): the empty rows the host left for it below its data rows.  The contract form
is ``allocate``: the host names, for each key no row has, the sheet row it starts in
(``{"t:stock/Lamp": 15}``, or ``(address, row)`` pairs; the tool's ``row`` on a cell).  A
row allocated is refused, before anything is written, when the table has no capacity
(``no_capacity``) or its capacity is used up (``capacity_full``, with the capacity), when
the row is outside the capacity (``not_capacity``, with its rows), already one of the
table's data rows (``data_row``), not empty (``not_free``: a value or a formula in a
column of the table, a formula whose result is empty text too), named for two keys or a
key named twice (``allocated_twice``), or when a key allocated is already a row's
(``existing_key``).  The rows allocated to a table continue its data rows with no blank
row between (``capacity_gap``, naming the next free row): the data rows end at the first
blank row, so a row past a gap would be no row of the table once saved and declared
again.  The host chooses which keys get which rows, and with its write policy which rows
may be written at all (``cells`` over some capacity rows, or a table's ``new_rows``:
:mod:`xlsx_agent.edit.policy`).

``new_rows=True`` is a convenience of the Python API: each key no row has and none
allocated gets the next free row, in the order the call first names it.  Under a write
policy it needs the table's ``new_rows: true`` in every policy (``not_allocated``
otherwise: the policy asks for rows named).  A new row's key cell is written with the key
-- as a number when it is a whole number's digits and every key cell of the table holds a
number, else as text -- and the call's other cells of that key fill the row; a key
allocated and named by no cell starts a row of its key alone.  A cell of a free capacity
row named by its address is refused (``capacity_row``): a new row is started by its key.
Once written, the row is one of the table's data rows (:meth:`Regions.refresh_capacity
<xlsx_agent.read.regions.Regions.refresh_capacity>`); undone, it is free again.

Every cell named by key is then written as a cell named by its address is: the host's
write policy, a protected sheet, the expectations, the values and the guards are asked of
the cell it is (:mod:`xlsx_agent.edit.cells`).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Mapping

from ooxml_edit.sheets import Area, Cell

from ..address import AddressError, NameNotFound, RangeAddress, SheetRef, TableAddress
from ..read.regions import Region, RegionError, capacity_rows, row_empty

if TYPE_CHECKING:  # pragma: no cover
    from ..workbook import Sheet, Workbook

#: A key written as a number into a table whose key cells all hold numbers.
_WHOLE = re.compile(r"-?(?:0|[1-9][0-9]{0,14})")


@dataclass(frozen=True)
class Found:
    """A cell named by key, found: its sheet and cell."""

    sheet: "Sheet"
    cell: Cell


@dataclass(frozen=True)
class NewCell:
    """A cell named by a key no row has: its table, the key and its own column."""

    region: Region
    key: str
    column: int


def is_keyed(address: str) -> bool:
    """Whether ``address`` names a table's cell (``t:...``) rather than a sheet's."""
    return address.lstrip().startswith("t:")


def resolve(workbook: "Workbook", text: str, new_rows: bool,
            refuse: Any, allocated: "set[tuple[str, str]] | frozenset" = frozenset()
            ) -> "Found | NewCell | None":
    """The cell ``text`` (a ``t:`` address of one value by key) names; a :class:`NewCell`
    for a key no row has when ``new_rows`` or when ``allocated`` holds its ``(table,
    key)``; ``None`` when refused (``refuse(code, detail)`` noted it)."""
    try:
        address = TableAddress.parse(text.strip())
    except (AddressError, ValueError):
        refuse("bad_address", "a value by key: t:stock/<key>/<column>")
        return None
    if address.position is not None or address.row_range is not None \
            or address.keys is not None or address.any_row:
        refuse("positional", "write one value as t:<table>/<key>/<column>")
        return None
    if address.kind != "value" or address.key is None or address.any_column:
        refuse("bad_address", "one value of a table by key: t:<table>/<key>/<column>")
        return None
    try:
        region = workbook.regions.region(address.whole)
        index = region.column_index(address.column)
    except (NameNotFound, KeyError):
        refuse("bad_address", "no such table or column")
        return None
    except RegionError as error:
        refuse("bad_address", error.code)
        return None
    keys = _keys(workbook, region, refuse)
    if keys is None:
        return None
    if address.key in keys.duplicates:
        refuse("duplicate_key", f"in {region.address.text}")
        return None
    if address.key not in keys:
        if new_rows or (region.address.text, address.key) in allocated:
            return NewCell(region, address.key, index)
        refuse("missing_key", f"in {region.address.text}; allocate names a free row of its "
               "capacity for a new row")
        return None
    position = keys.find(address.key)
    data = region.data
    cell = Cell(data.first.row + position - 1, region.area.first.column + index)
    return _real(workbook, region, cell)


def _keys(workbook: "Workbook", region: Region, refuse: Any) -> Any:
    """``region``'s key index, or ``None`` (refused) when it has no key column declared or
    one holding a formula."""
    if region.key is None:
        refuse("no_key_column", f"{region.address.text}: its owner declares one (key)")
        return None
    keys = workbook.regions.key_index(region)
    if keys.formulas:
        refuse("formula_key", f"{region.address.text}: its key column {region.key!r} holds "
               f"a formula in {len(keys.formulas)} data row(s), whose keys are cached results "
               "that move with the rows; declare a key column of values the application "
               "supplies (key)")
        return None
    return keys


def _real(workbook: "Workbook", region: Region, cell: Cell) -> Found:
    """The sheet's cell of ``region``'s own ``cell``."""
    if region.span is None:
        return Found(workbook.sheet_of(region.sheet_ref), cell)
    (found,) = region.span.locate(Area(cell, cell))
    return Found(workbook.sheet_of(found.sheet), found.area.first)


@dataclass(frozen=True)
class Row:
    """A new row given out: its table, its key, the sheet and row it is."""

    region: Region
    key: str
    sheet: "Sheet"
    row: int
    key_column: int
    key_value: Any


@dataclass(frozen=True)
class Allocation:
    """A row the host allocates to a key no row has (``allocate``): its table, the key,
    the sheet row, and its place in the call (``allocate[0]``)."""

    region: Region
    key: str
    row: int
    shown: str


def allocations(workbook: "Workbook", given: Any, refuse: Any) -> "list[Allocation]":
    """``given`` (``allocate``: a mapping of a row by key, ``t:stock/Lamp``, to its sheet
    row, or ``(address, row)`` pairs) read; each item refused (``refuse(code, shown,
    detail)``) is left out."""
    from . import EditError

    if given is None:
        return []
    items = list(given.items()) if isinstance(given, Mapping) else list(given)
    out: list[Allocation] = []
    for n, item in enumerate(items):
        shown = f"allocate[{n}]"
        if not isinstance(item, (tuple, list)) or len(item) != 2:
            raise EditError("bad_address", "allocate is a mapping of a row by key "
                            "(t:stock/<key>) to its sheet row, or (address, row) pairs")
        text, row = item
        if isinstance(row, bool) or not isinstance(row, int) or row < 1:
            refuse("bad_value", shown, "a row is the sheet's row number, a whole number")
            continue
        try:
            address = TableAddress.parse(text.strip()) if isinstance(text, str) else None
        except (AddressError, ValueError):
            address = None
        if address is None or address.kind != "row" or address.key is None:
            refuse("bad_address", shown, "a row by key: t:stock/<key>")
            continue
        try:
            region = workbook.regions.region(address.whole)
        except (NameNotFound, KeyError):
            refuse("bad_address", shown, "no such table")
            continue
        except RegionError as error:
            refuse("bad_address", shown, error.code)
            continue
        keys = _keys(workbook, region, lambda code, detail=None: refuse(code, shown, detail))
        if keys is None:
            continue
        if address.key in keys:
            refuse("existing_key", shown, f"a row of {region.address.text} has the key: "
                   "write its cells by key, without allocate")
            continue
        out.append(Allocation(region, address.key, row, shown))
    return out


def allocate(workbook: "Workbook", wanted: "list[NewCell]",
             explicit: "list[Allocation]" = (), first_free: bool = False
             ) -> "tuple[dict, dict]":
    """The rows the keys get: those of ``explicit`` (checked, see the module), then, with
    ``first_free``, the next free rows for the keys of ``wanted`` none allocates, in the
    order first named: ``(rows, refused)``, ``rows`` a :class:`Row` by ``(table, key)``,
    ``refused`` the code and detail of each ``(table, key)`` that gets none."""
    rows: dict[tuple[str, str], Row] = {}
    refused: dict[tuple[str, str], tuple[str, str]] = {}
    taken: dict[str, int] = {}
    by_table: dict[str, list[Allocation]] = {}
    named: dict[tuple[str, str], int] = {}
    for item in explicit:
        key = (item.region.address.text, item.key)
        named[key] = named.get(key, 0) + 1
        by_table.setdefault(key[0], []).append(item)
    for name, items in by_table.items():
        region = items[0].region
        room = capacity_rows(region)
        if room is None:
            for item in items:
                refused[(name, item.key)] = ("no_capacity", f"{name}: its owner declares one "
                                             "(capacity)")
            continue
        writer = workbook.editor.cells.sheet(workbook.sheet_of(SheetRef(name=room.sheet))
                                             .stored_name)
        rows_named: dict[int, int] = {}
        for item in items:
            rows_named[item.row] = rows_named.get(item.row, 0) + 1
        good: list[Allocation] = []
        for item in items:
            key = (name, item.key)
            if named[key] > 1 or rows_named[item.row] > 1:
                refused[key] = ("allocated_twice", "a key and a row are allocated once")
            elif room.free == 0:
                refused[key] = ("capacity_full", f"{name} holds {region.capacity} rows at "
                                "most")
            elif not room.first <= item.row <= room.end:
                refused[key] = ("not_capacity", f"{name}'s capacity is rows {room.first} to "
                                f"{room.end}, {room.free} free from row {room.last + 1}")
            elif item.row <= room.last:
                refused[key] = ("data_row", f"row {item.row} is one of {name}'s data rows; "
                                f"its free rows start at row {room.last + 1}")
            elif not row_empty(writer, item.row, room.left, room.right):
                refused[key] = ("not_free", f"row {item.row} holds a value or a formula in "
                                f"{name}'s columns")
            else:
                good.append(item)
        expected = set(range(room.last + 1, room.last + 1 + len(good)))
        for item in good:
            key = (name, item.key)
            if item.row not in expected:
                refused[key] = ("capacity_gap", f"rows are taken in order, with no blank row "
                                f"between: the next free row of {name} is {room.last + 1}")
                continue
            sheet = workbook.sheet_of(SheetRef(name=room.sheet))
            rows[key] = _row(workbook, region, item.key, sheet, item.row, room)
        taken[name] = room.last + len(good)
    for item in wanted:
        name = item.region.address.text
        key = (name, item.key)
        if key in rows or key in refused or key in named or not first_free:
            continue
        room = capacity_rows(item.region)
        if room is None:
            refused[key] = ("no_capacity", f"{name}: its owner declares one (capacity)")
            continue
        if not _first_free_allowed(workbook, item.region):
            refused[key] = ("not_allocated", f"the write policy lets {name} take new rows "
                            "only where allocate names them")
            continue
        row = taken.get(name, room.last) + 1
        sheet = workbook.sheet_of(SheetRef(name=room.sheet))
        writer = workbook.editor.cells.sheet(sheet.stored_name)
        if row > room.end or not row_empty(writer, row, room.left, room.right):
            refused[key] = ("capacity_full", f"{name} holds {item.region.capacity} rows at "
                            "most" if row > room.end else
                            f"row {row}, the next of {name}'s capacity, is not empty")
            continue
        taken[name] = row
        rows[key] = _row(workbook, item.region, item.key, sheet, row, room)
    return rows, refused


def _row(workbook: "Workbook", region: Region, key: str, sheet: "Sheet", row: int,
         room: Any) -> Row:
    key_index = region.column_index(region.key)
    return Row(region, key, sheet, row, room.columns[key_index],
               _key_value(workbook, region, key))


def _first_free_allowed(workbook: "Workbook", region: Region) -> bool:
    """Whether every write policy lets ``region`` take its next free rows (a table entry
    of it with ``new_rows``); without a policy, the host's own code, it may."""
    for policy in workbook.policies:
        if not any(entry.new_rows and _same_table(workbook, entry.table, region)
                   for entry in policy.tables):
            return False
    return True


def _same_table(workbook: "Workbook", table: str, region: Region) -> bool:
    try:
        return workbook.regions.region(table).address.text == region.address.text
    except (NameNotFound, RegionError, KeyError, ValueError):
        return False


def _key_value(workbook: "Workbook", region: Region, key: str) -> Any:
    """The key as its new cell holds it: a whole number when it is one's digits and every
    key cell of the table holds a number, else the text."""
    if not _WHOLE.fullmatch(key) or not region.rows:
        return key
    area = region.column_area(region.key)
    for row in workbook.read(RangeAddress(region.sheet_ref, area)):
        if row[0].t != "number":
            return key
    return int(key)


def free_rows(workbook: "Workbook") -> dict[str, list[Area]]:
    """The free capacity rows of every declared table with a capacity, by sheet part: the
    cells an address may not name (``capacity_row``)."""
    out: dict[str, list[Area]] = {}
    for region in workbook.regions.configured:
        room = capacity_rows(region)
        if room is None or room.last >= room.end:
            continue
        part = workbook.sheet_of(SheetRef(name=room.sheet)).part
        out.setdefault(part, []).append(Area(Cell(room.last + 1, room.left),
                                             Cell(room.end, room.right)))
    return out


__all__ = ["Allocation", "Found", "NewCell", "Row", "allocate", "allocations", "free_rows",
           "is_keyed", "resolve"]
