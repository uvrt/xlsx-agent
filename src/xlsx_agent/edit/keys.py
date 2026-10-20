"""Cells written by key (0.2, X4): ``t:stock/P-7/Quantity`` in :meth:`Workbook.set_cells
<xlsx_agent.workbook.Workbook.set_cells>`, and new rows in a declared table's capacity.

::

    wb.declare_tables([{"name": "stock", "sheets": "Stock", "header_row": 9,
                        "key": "Item", "capacity": 200}])
    wb.set_cells({"t:stock/Kettle/Quantity": 4})          # the row whose Item is Kettle
    wb.set_cells({"t:stock/Lamp/Quantity": 2,             # Lamp is no row's key yet:
                  "t:stock/Lamp/Category": "Hardware"},   # a new row, in the first free
                 new_rows=True)                           # row of the table's capacity

**By key, never by place.**  A value by key names the row whose key cell holds the key
(:mod:`xlsx_agent.read.regions`: text exactly, or a whole number's digits), wherever the
row is now: sorting the table, or a reviewer moving rows, does not move the write.  The
table (an Excel table, a declared table or region; any ``t:`` region) must have a key
column declared for it -- ``key`` in its declaration, ``declare_region(key=...)`` or
:meth:`Workbook.declare_key <xlsx_agent.workbook.Workbook.declare_key>` -- since a key
column guessed from the values is no address to write by (``no_key_column``).  A key two
rows share refuses the cell (``duplicate_key``); a key no row has refuses it
(``missing_key``) unless the call asks for new rows; a row by position (``t:stock/#3``), a
key set, a row range, a wildcard or a whole row or column is refused (``positional``):
there is no fallback to a place.  A key is cell content, so a refusal names such a value
by its place in the call (``cells[3]``: the fourth cell given) and its table, never by its
key; a value found is named by its cell's address (``Stock!D12``).

**New rows** (``new_rows=True``, the tool's ``new_rows: true``) go only into a declared
table's capacity (:mod:`xlsx_agent.declare`, ``capacity``): the empty rows the host left
for it below its data rows.  Each key no row has gets the first free row, in the order
the call first names it; its key cell is written with the key -- as a number when it is a
whole number's digits and every key cell of the table holds a number, else as text -- and
the call's other cells of that key fill the row.  A table without a capacity refuses
(``no_capacity``); one whose capacity is used up refuses the keys past it
(``capacity_full``), all before anything is written.  A free row is empty: no value and no
formula in any column of the table, so a row holding a formula whose result is empty
text is never taken.  A cell of a free capacity row named by its address is refused
(``capacity_row``): a new row is started by its key.  Once written, the row is one of the
table's data rows (:meth:`Regions.refresh_capacity
<xlsx_agent.read.regions.Regions.refresh_capacity>`); undone, it is free again.

Every cell named by key is then written as a cell named by its address is: the host's
write policy, a protected sheet, the expectations, the values and the guards are asked of
the cell it is (:mod:`xlsx_agent.edit.cells`).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

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
            refuse: Any) -> "Found | NewCell | None":
    """The cell ``text`` (a ``t:`` address of one value by key) names; a :class:`NewCell`
    for a key no row has when ``new_rows``; ``None`` when refused (``refuse(code, detail)``
    noted it)."""
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
    if region.key is None:
        refuse("no_key_column", f"{region.address.text}: its owner declares one (key)")
        return None
    keys = workbook.regions.key_index(region)
    if address.key in keys.duplicates:
        refuse("duplicate_key", f"in {region.address.text}")
        return None
    if address.key not in keys:
        if new_rows:
            return NewCell(region, address.key, index)
        refuse("missing_key", f"in {region.address.text}; new_rows starts one in its "
               "capacity")
        return None
    position = keys.find(address.key)
    data = region.data
    cell = Cell(data.first.row + position - 1, region.area.first.column + index)
    return _real(workbook, region, cell)


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


def allocate(workbook: "Workbook", wanted: "list[NewCell]") -> "tuple[dict, dict]":
    """The rows the keys of ``wanted`` get, in the order first named: ``(rows, refused)``,
    ``rows`` a :class:`Row` by ``(table, key)``, ``refused`` the code and detail of each
    ``(table, key)`` that gets none."""
    rows: dict[tuple[str, str], Row] = {}
    refused: dict[tuple[str, str], tuple[str, str]] = {}
    taken: dict[str, int] = {}
    for item in wanted:
        name = item.region.address.text
        key = (name, item.key)
        if key in rows or key in refused:
            continue
        room = capacity_rows(item.region)
        if room is None:
            refused[key] = ("no_capacity", f"{name}: its owner declares one (capacity)")
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
        key_index = item.region.column_index(item.region.key)
        rows[key] = Row(item.region, item.key, sheet, row, room.columns[key_index],
                        _key_value(workbook, item.region, item.key))
    return rows, refused


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


__all__ = ["Found", "NewCell", "Row", "allocate", "free_rows", "is_keyed", "resolve"]
