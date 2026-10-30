"""The host's write policy (0.2, X4): where a workbook may be written, fixed by the
application, never by a model.

A policy is an allowlist of destinations, as plain data an application builds from its own
configuration::

    policy = WritePolicy.from_config({
        "cells": ["Stock!B4", "Stock!D10:E209", "Notes!"],   # cells, ranges, whole sheets
        "tables": [{"table": "stock", "columns": ["Quantity", "Unit price"],
                    "new_rows": True}],                       # a table's data cells, by column
        "formulas": False,               # no formula written (default: allowed)
        "replace_formulas": False,       # no formula written over (default: allowed)
    })
    wb = Workbook.open(data, policy=policy)      # or policy={...}, or ["Stock!D10:E209"]
    wb.narrow_policy({"cells": ["Stock!D10:D209"]})   # narrower from then on, never wider

**Destinations.**  ``cells``: cell and range addresses with their sheet (``Stock!D10``,
``Stock!D10:E209``, ``Stock!D:D``, ``Stock!3:5``, ``Stock!`` the whole sheet; ``#3!...``
by the sheet's id).  ``tables``: a table, declared table or region by name (``stock`` or
``t:stock``), its data cells -- every visible column's, or ``columns``' (by name, id or
header) -- on every sheet it spans, as its rows are when the call is made (a table with a
``capacity`` grows: :mod:`xlsx_agent.declare`); never its header or totals rows.
``new_rows: true`` also lets a new row by key be written into its capacity (its key column
and ``columns``' cells there), and lets ``new_rows=True`` give a new key the next free row;
without it, a policy lets a new row be written only into capacity rows its ``cells`` cover,
each allocated by the host (``allocate``: :mod:`xlsx_agent.edit.keys`).  A cell is writable
when a destination covers it; a policy with no destination (``WritePolicy.read_only()``,
``{"cells": []}``) lets nothing be written.

**Whatever writes.**  Every write is checked against the policy before anything changes:
a cell named by its address, by key (``t:stock/P-7/Quantity``), a new row's cells, a
formula, a write in a batch or a tool batch's op.  :meth:`Editor.write
<xlsx_agent.edit.Editor.write>`, the step every write goes through, checks it again.
``formulas: false`` refuses a formula anywhere; ``replace_formulas: false`` refuses a write
over a cell that holds one, whatever the call's ``replace_formulas`` says.  A shared
formula group handed on when its master is written (:mod:`xlsx_agent.edit.cells`) keeps
every member's formula as it was, so its members need not be writable.

**Without a policy** (decided, 2026-10-20): the Python API writes everywhere but on a
protected sheet -- the host's own code; a tool session writes as its ``edit`` role allows
(a ``review`` session has no tool that writes).  A protected sheet is refused under any
policy (:mod:`xlsx_agent.edit.cells`).

**Narrowing.**  :meth:`Workbook.narrow_policy <xlsx_agent.workbook.Workbook.narrow_policy>`
(and ``xlsx_agent.tools.narrow`` for a tool session) adds a policy: a write must then pass
every one, so a policy added never widens what an earlier one allows.  Nothing takes one
away.

**Checked.**  A policy's form is checked when it is made (:class:`PolicyError`, naming the
entry: ``cells[2]``); its sheets and tables against the workbook when a write is planned
(``bad_policy``, the same entry): a table is often declared after the workbook is opened.
The tool layer checks both when a workbook joins the session, refusing ``open_document``.
The policy is shown to the model by ``describe`` (``writable``, :meth:`WritePolicy.as_dict`),
in a few tokens: the addresses and tables as given.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Mapping, Sequence

from ooxml_edit.sheets import Area, Cell

from ..address import (AddressError, NameNotFound, RangeAddress, SheetRef, TableAddress,
                       nearest, parse_address)

if TYPE_CHECKING:  # pragma: no cover
    from ..workbook import Workbook

#: The keys a policy takes.
POLICY_KEYS = ("cells", "tables", "formulas", "replace_formulas")

#: The keys a policy's table entry takes.
TABLE_ENTRY_KEYS = ("table", "columns", "new_rows")

#: The most destinations one policy lists.
MAX_DESTINATIONS = 1_000


class PolicyError(ValueError):
    """A write policy refused: ``code`` (``bad_policy``), ``path`` the entry
    (``cells[2]``, ``tables[0].columns[1]``), ``valid_options`` what would do.  The message
    names addresses, tables and columns, never a cell's content."""

    def __init__(self, path: str, message: str, *, valid_options: Sequence[str] = ()) -> None:
        super().__init__(f"write policy {path}: {message}" if path
                         else f"write policy: {message}")
        self.code = "bad_policy"
        self.path = path
        self.message = str(self)
        self.valid_options = list(valid_options)[:50]


@dataclass(frozen=True)
class TableEntry:
    """A table a policy lets be written: ``table`` as ``t:<name>``, ``columns`` (``None``:
    every visible one), ``new_rows`` whether a new row by key may go into its capacity."""

    table: str
    columns: tuple[str, ...] | None = None
    new_rows: bool = False

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"table": self.table}
        if self.columns is not None:
            out["columns"] = list(self.columns)
        if self.new_rows:
            out["new_rows"] = True
        return out


@dataclass(frozen=True)
class WritePolicy:
    """Where a workbook may be written (see the module): immutable; made by
    :meth:`from_config`, or :meth:`read_only`."""

    cells: tuple[str, ...] = ()
    tables: tuple[TableEntry, ...] = ()
    formulas: bool = True
    replace_formulas: bool = True

    @classmethod
    def from_config(cls, config: Any) -> "WritePolicy":
        """A policy from plain data: a mapping of :data:`POLICY_KEYS`, a list of cell
        addresses (``cells`` alone), or a :class:`WritePolicy` (itself).
        :class:`PolicyError` for anything else, naming the entry."""
        return as_policy(config)

    @classmethod
    def read_only(cls) -> "WritePolicy":
        """A policy that lets nothing be written."""
        return cls()

    @property
    def is_read_only(self) -> bool:
        """No destination: nothing may be written."""
        return not self.cells and not self.tables

    def as_dict(self) -> "dict[str, Any] | str":
        """What ``describe`` shows (``writable``): the destinations as given, and the
        formula rules when they refuse; ``"nothing"`` for a read-only policy."""
        if self.is_read_only:
            return "nothing"
        out: dict[str, Any] = {}
        if self.cells:
            out["cells"] = list(self.cells)
        if self.tables:
            out["tables"] = [entry.as_dict() for entry in self.tables]
        if not self.formulas:
            out["formulas"] = False
        if not self.replace_formulas:
            out["replace_formulas"] = False
        return out


def as_policy(config: Any) -> WritePolicy:
    """``config`` as a :class:`WritePolicy` (see :meth:`WritePolicy.from_config`)."""
    if isinstance(config, WritePolicy):
        return config
    if isinstance(config, (list, tuple)):
        config = {"cells": list(config)}
    if not isinstance(config, Mapping):
        raise PolicyError("", "a policy is a mapping of cells, tables, formulas and "
                          "replace_formulas, or a list of cell addresses",
                          valid_options=list(POLICY_KEYS))
    for key in config:
        if key not in POLICY_KEYS:
            raise PolicyError(str(key), "unknown key",
                              valid_options=nearest(str(key), POLICY_KEYS))
    cells = _list(config.get("cells", ()), "cells")
    tables = _list(config.get("tables", ()), "tables")
    if len(cells) + len(tables) > MAX_DESTINATIONS:
        raise PolicyError("", f"{len(cells) + len(tables):,} destinations; at most "
                          f"{MAX_DESTINATIONS:,}: give ranges")
    out_cells = tuple(_cells(item, f"cells[{n}]") for n, item in enumerate(cells))
    out_tables = tuple(_table(item, f"tables[{n}]") for n, item in enumerate(tables))
    flags = {}
    for key in ("formulas", "replace_formulas"):
        value = config.get(key, True)
        if not isinstance(value, bool):
            raise PolicyError(key, "true or false")
        flags[key] = value
    return WritePolicy(cells=out_cells, tables=out_tables, **flags)


def _list(value: Any, path: str) -> list:
    if isinstance(value, (str, bytes, Mapping)) or not isinstance(value, Sequence):
        raise PolicyError(path, "a list")
    return list(value)


def _cells(item: Any, path: str) -> str:
    if not isinstance(item, str) or not item.strip():
        raise PolicyError(path, "a cell or range address with its sheet: Stock!D10:E209")
    try:
        parsed = parse_address(item.strip())
    except (AddressError, ValueError):
        parsed = None
    if not isinstance(parsed, RangeAddress) or parsed.sheet is None:
        raise PolicyError(path, "a cell or range address with its sheet (Stock!D10:E209, "
                          "Stock!D:D, Stock! for the whole sheet); a table goes in tables")
    return parsed.text


def _table(item: Any, path: str) -> TableEntry:
    if isinstance(item, str):
        item = {"table": item}
    if not isinstance(item, Mapping):
        raise PolicyError(path, "a table: a mapping with table, and columns or new_rows",
                          valid_options=list(TABLE_ENTRY_KEYS))
    for key in item:
        if key not in TABLE_ENTRY_KEYS:
            raise PolicyError(f"{path}.{key}", "unknown key",
                              valid_options=nearest(str(key), TABLE_ENTRY_KEYS))
    name = item.get("table")
    if not isinstance(name, str) or not name.strip():
        raise PolicyError(f"{path}.table", "a table's name, as stock or t:stock")
    text = name.strip()
    if not text.startswith("t:"):
        text = f"t:{text}"
    try:
        address = TableAddress.parse(text)
    except (AddressError, ValueError):
        address = None
    if address is None or address.kind not in ("table", "region"):
        raise PolicyError(f"{path}.table", "a table's name, as stock or t:stock: its columns "
                          "go in columns")
    columns = item.get("columns")
    if columns is not None:
        columns = _list(columns, f"{path}.columns")
        if not columns or not all(isinstance(c, str) and c for c in columns):
            raise PolicyError(f"{path}.columns", "column names, ids or headers")
        columns = tuple(columns)
    new_rows = item.get("new_rows", False)
    if not isinstance(new_rows, bool):
        raise PolicyError(f"{path}.new_rows", "true or false")
    return TableEntry(table=address.text, columns=columns, new_rows=new_rows)


# -- against a workbook ------------------------------------------------------------------------


class Allowed:
    """A workbook's policies resolved against it as it is now: by sheet part, the areas
    each policy lets be written; :meth:`allows` asks every one."""

    def __init__(self, workbook: "Workbook", policies: "Sequence[WritePolicy]") -> None:
        self.policies = tuple(policies)
        self.formulas = all(p.formulas for p in self.policies)
        self.replace_formulas = all(p.replace_formulas for p in self.policies)
        self._areas = [_areas(workbook, policy) for policy in self.policies]

    def allows(self, part: str, cell: Cell) -> bool:
        """Whether every policy lets ``cell`` of the sheet part ``part`` be written."""
        return all(any(area.contains(cell) for area in areas.get(part, ()))
                   for areas in self._areas)


def _areas(workbook: "Workbook", policy: WritePolicy) -> dict[str, list[Area]]:
    """The areas ``policy`` lets be written, by sheet part; :class:`PolicyError` for a
    sheet, table or column the workbook does not have."""
    from ..read.regions import RegionError, capacity_rows

    out: dict[str, list[Area]] = {}
    for n, text in enumerate(policy.cells):
        address = parse_address(text)
        assert isinstance(address, RangeAddress)   # as_policy took only ranges
        try:
            sheet = workbook.sheet_of(address)
        except (NameNotFound, KeyError, ValueError):
            raise PolicyError(f"cells[{n}]", "names no sheet of the workbook") from None
        if sheet.part is None:
            raise PolicyError(f"cells[{n}]", "names a sheet that holds no cells")
        out.setdefault(sheet.part, []).append(address.area)
    for n, entry in enumerate(policy.tables):
        path = f"tables[{n}]"
        try:
            region = workbook.regions.region(entry.table)
        except (NameNotFound, RegionError, KeyError, ValueError):
            raise PolicyError(f"{path}.table", f"{entry.table} is no table, name or region "
                              "of the workbook", valid_options=workbook.regions.known()[:50]
                              ) from None
        derived = region.derived or (None,) * len(region.columns)
        columns = entry.columns if entry.columns is not None else \
            [c for c, source in zip(region.columns, derived) if source is None]
        own: list[int] = []
        for m, column in enumerate(columns):
            try:
                own.append(region.column_index(column))
            except NameNotFound:
                raise PolicyError(f"{path}.columns[{m}]", f"{entry.table} has no column "
                                  f"{column!r}", valid_options=list(region.columns)) from None
            source = derived[own[-1]]
            if source is not None:
                raise PolicyError(f"{path}.columns[{m}]", f"{column!r} is read from column "
                                  f"{source!r} by the table's own rules, and is never "
                                  f"written: name {source!r}", valid_options=[source])
        data = region.data
        for index in own:
            at = region.area.first.column + index
            if data is not None:
                for sheet_name, area in _located(region, Area(Cell(data.first.row, at),
                                                              Cell(data.last.row, at))):
                    out.setdefault(_part(workbook, sheet_name), []).append(area)
        if entry.new_rows:
            room = capacity_rows(region)
            if room is None:
                raise PolicyError(f"{path}.new_rows", f"{entry.table} has no capacity to "
                                  "take new rows: declare one (capacity)")
            if region.key is None:
                raise PolicyError(f"{path}.new_rows", f"{entry.table} has no key column: "
                                  "new rows are written by key")
            wanted = set(own) | {region.column_index(region.key)}
            part = _part(workbook, room.sheet)
            for index in sorted(wanted):
                at = room.columns[index]
                if room.last < room.end:
                    out.setdefault(part, []).append(
                        Area(Cell(room.last + 1, at), Cell(room.end, at)))
    return out


def _located(region: Any, area: Area) -> list[tuple[str, Area]]:
    """Cells of ``region``'s own rectangle as the sheets' (name, area)."""
    if region.span is None:
        return [(region.sheet, area)]
    return [(found.sheet.name, found.area) for found in region.span.locate(area)]


def _part(workbook: "Workbook", sheet_name: str) -> str:
    part = workbook.sheet_of(SheetRef(name=sheet_name)).part
    assert part is not None   # a table's sheet holds cells
    return part



__all__ = ["Allowed", "MAX_DESTINATIONS", "POLICY_KEYS", "PolicyError", "TABLE_ENTRY_KEYS",
           "TableEntry", "WritePolicy", "as_policy"]
