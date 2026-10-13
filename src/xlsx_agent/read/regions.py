"""Regions: Excel tables, defined names over ranges, declared regions and detected ones, and
how a ``t:`` address resolves to cells.

Decision 1 of ROADMAP.md: Excel tables, defined names and regions the caller declares come
first; detection is only for a simple single-row header over a block of data, and a merged
or multi-row header is declared by the caller, never guessed.  All four are a
:class:`Region`: a rectangle on one sheet with its header rows, its data rows and any
totals rows, and a name for each column::

    wb.region("t:Products")              # Region(kind='table', name='Products', ...)
    wb.resolve("t:Products[Price]")      # Selection: the column's data cells, Stock!C2:C40
    wb.resolve("t:Products/SKU-0042/Price", key="SKU")
    wb.declare_region("Returns", "Returns!A3:F200", header_rows=2,
                      headers=["SKU", "Date", "Qty", "Reason", "Store", "Refund"])
    wb.detect_region("t:Stock!A1")       # a single-row header at A1, or RegionError

**Which name.**  ``t:Name`` is matched, case aside (names "are not case-sensitive",
ECMA-376 Part 1, §18.17), against the tables' ``displayName``, the workbook-wide defined
names and the declared regions; ``t:Stock!Name`` against the names local to ``Stock``
only.  A name local to a sheet is never reached by ``t:Name`` (Excel reaches it so only
from its own sheet); the refusal offers ``t:Stock!Name``.  Two of them with one name (a
file Excel would not write) are refused as ``ambiguous_name``, never one picked.

**Columns.**  A table's columns are its ``tableColumn`` names; a declared region's are the
names declared, or its header row's text; a detected region's its header row's text; a
defined name's range has no header row, and its columns are named by their letter
(``Column B``): which row of a name's range is a header is not guessed.  A blank column
name, or one two columns share (case aside), becomes the letter name and is flagged
(``blank_header``, ``duplicate_header``).  Column names are matched case aside, as names
are.  **Column names are names**: they may appear in messages and ``valid_options``, as
table and sheet names may (ROADMAP.md, "Addresses"); keys are cell content and never do.

**Keys.**  ``t:Products/SKU-0042`` finds the row whose key column holds ``SKU-0042``.  The
key column is the one the caller names (``key=``, or the declared region's ``key``), or
else the only column :meth:`Regions.key_columns` finds (unique, never blank, text or
whole numbers); with none, or more than one, the lookup is refused (``no_key_column``)
with those columns as ``valid_options``.  A key is matched exactly (case counts) against
the cell's text, or a whole number's digits (``42`` matches 42 and 42.0).  A key two rows
share is refused (``duplicate_key``) with the rows' positions (``t:Products/#3``) as
``valid_options``; never the first of them.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Iterator, Sequence

from ooxml_edit.charts.workbook import Area, Cell, column_letters

from ..address import (NameNotFound, RangeAddress, SheetRef, TableAddress,
                       is_address_name, is_name, nearest, parse_address, quote_sheet)
from .names import DefinedName
from .tables import TablePart, name_columns
from .types import Problem, Value

if TYPE_CHECKING:  # pragma: no cover
    from ..profile import Column, ColumnStats
    from ..query import QueryResult
    from ..workbook import Sheet, Workbook

#: What a :class:`Region` is.
REGION_KINDS = ("table", "name", "declared", "detected")

#: Why a header is not detected (:attr:`RegionError.reason` of ``header_unclear``).
UNCLEAR_REASONS = ("merged_header", "rows_above", "multi_row_header", "header_not_text")


class RegionError(ValueError):
    """A ``t:`` address, or a region, the workbook cannot give: ``code`` says why.

    ``not_a_range`` (a defined name that is a formula or a constant), ``ambiguous_name``,
    ``scoped_name`` (``t:Name`` for a name local to a sheet), ``header_unclear`` (with
    ``reason``, one of :data:`UNCLEAR_REASONS`), ``no_header`` (no header at the cell),
    ``no_key_column``, ``duplicate_key``, ``position_out_of_range``, ``bad_table``,
    ``name_taken``, ``bad_name``, ``headers_required``, ``bad_headers``.  The message
    names sheets, tables, names and columns, never a key or a cell's value; ``address``
    is the address in question as text, ``valid_options`` what would work instead."""

    def __init__(self, code: str, message: str, *, address: str | None = None,
                 reason: str | None = None, valid_options: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.address = address
        self.reason = reason
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class Region:
    """A table, a defined name over a range, a declared region or a detected one.

    ``kind`` is one of :data:`REGION_KINDS`; ``name`` the name a ``t:`` address uses
    (``None`` for a detected region, which is named by its header's first cell); ``scope``
    the sheet a defined name is local to; ``sheet`` the sheet it lies on, as the workbook
    spells it, and ``sheet_id`` its ``sheetId``; ``area`` all of it (header and totals rows
    included); ``header_rows`` and ``totals_rows`` how many of its rows are those;
    ``columns`` a name per column.  ``key`` the key column declared (or given to
    :meth:`Workbook.table <xlsx_agent.workbook.Workbook.table>`); ``hidden`` a hidden
    defined name; ``flags`` what is odd, by code and address.  A table also has ``id``,
    ``object_name`` (its ``name``), ``auto_filter``, ``totals`` (each column's totals-row
    function) and ``table_type``.

    A region the workbook gave knows its workbook, so that its columns can be read:
    :meth:`column` and :meth:`profile`."""

    kind: str
    name: str | None
    sheet: str
    area: Area
    columns: tuple[str, ...]
    header_rows: int = 1
    totals_rows: int = 0
    key: str | None = None
    scope: str | None = None
    hidden: bool = False
    sheet_id: int | None = None
    flags: tuple[Problem, ...] = ()
    id: int | None = None
    object_name: str | None = None
    auto_filter: Area | None = None
    totals: tuple[str | None, ...] = ()
    table_type: str | None = None
    workbook: "Workbook | None" = field(default=None, repr=False, compare=False)

    @property
    def sheet_ref(self) -> SheetRef:
        return SheetRef(name=self.sheet) if self.sheet else SheetRef(id=self.sheet_id)

    @property
    def address(self) -> TableAddress:
        """``t:Products``, ``t:Stock!Prices`` (a name local to a sheet) or ``t:Stock!A1``
        (a detected region)."""
        if self.name is None:
            return TableAddress(region=RangeAddress(self.sheet_ref,
                                                    Area(self.area.first, self.area.first)))
        scope = SheetRef(name=self.scope) if self.scope is not None else None
        return TableAddress(name=self.name, scope=scope)

    @property
    def range(self) -> RangeAddress:
        """All of it, as a range address."""
        return RangeAddress(self.sheet_ref, self.area)

    @property
    def header(self) -> Area | None:
        """The header rows; ``None`` when it has none."""
        if not self.header_rows:
            return None
        first = self.area.first
        return Area(first, Cell(first.row + self.header_rows - 1, self.area.last.column))

    @property
    def data(self) -> Area | None:
        """The data rows, between the header and totals rows; ``None`` when there are none."""
        top = self.area.first.row + self.header_rows
        bottom = self.area.last.row - self.totals_rows
        if top > bottom:
            return None
        return Area(Cell(top, self.area.first.column), Cell(bottom, self.area.last.column))

    @property
    def totals_area(self) -> Area | None:
        """The totals rows; ``None`` when it has none."""
        if not self.totals_rows:
            return None
        last = self.area.last
        return Area(Cell(last.row - self.totals_rows + 1, self.area.first.column), last)

    @property
    def rows(self) -> int:
        """How many data rows."""
        data = self.data
        return 0 if data is None else data.last.row - data.first.row + 1

    def column_index(self, name: str) -> int:
        """The position (from 0) of the column called ``name``, matched case aside;
        :class:`~xlsx_agent.address.NameNotFound` with the nearest column names."""
        folded = name.casefold()
        for index, column in enumerate(self.columns):
            if column.casefold() == folded:
                return index
        raise NameNotFound(f"{self.address.text} has no column {name!r}; its columns are "
                           f"{list(self.columns)}", what="column", name=name,
                           candidates=list(self.columns),
                           valid_options=nearest(name, self.columns))

    def column_area(self, name: str) -> Area | None:
        """The data cells of the column called ``name``; ``None`` when there are no data
        rows."""
        column = self.area.first.column + self.column_index(name)
        data = self.data
        if data is None:
            return None
        return Area(Cell(data.first.row, column), Cell(data.last.row, column))

    def column(self, name: str) -> "Column":
        """The column called ``name`` (matched case aside), whose
        :meth:`~xlsx_agent.profile.Column.stats` is its profile."""
        from ..profile import Column
        return Column(self._workbook(), self, name)

    def profile(self, columns: "Sequence[str] | None" = None) -> "list[ColumnStats]":
        """The profiles of ``columns`` (every column when ``None``), in one streamed pass
        (:mod:`xlsx_agent.profile`)."""
        from ..profile import profile
        return profile(self._workbook(), self, columns)

    def query(self, columns: "Sequence[str] | None" = None, *, where=None, sort=None,
              limit: int | None = None, offset: int = 0) -> "QueryResult":
        """The rows that pass ``where`` (``{"Category": {"eq": "Hardware"}}``), ordered by
        ``sort`` (``[("Price", "desc")]``), from ``offset`` and at most ``limit`` of them,
        with ``columns`` (every column when ``None``): :func:`xlsx_agent.query.query`."""
        from ..query import query
        return query(self._workbook(), self, columns, where=where, sort=sort, limit=limit,
                     offset=offset)

    def _workbook(self) -> "Workbook":
        if self.workbook is None:
            raise ValueError(f"{self.address.text} was made without its workbook; get it "
                             "from Workbook.table or Workbook.region")
        return self.workbook


@dataclass(frozen=True)
class Selection:
    """What a ``t:`` address names, resolved: ``address`` its canonical form; ``kind``
    ``"table"`` or ``"region"`` (the whole of it), ``"column"``, ``"row"`` or ``"value"``;
    ``region`` the :class:`Region`; ``range`` the data cells selected (header and totals
    rows excluded; ``None`` when the region has no data rows); ``row`` the data row's
    position (from 1) for a row or a value; ``column`` the column's name as the region
    spells it."""

    address: str
    kind: str
    region: Region
    range: RangeAddress | None
    row: int | None = None
    column: str | None = None


@dataclass(frozen=True)
class UnclearRegion:
    """A block :meth:`Regions.find` met whose header is not detected: ``address`` is its
    first cell as a region address (``t:Stock!A1``), ``reason`` one of
    :data:`UNCLEAR_REASONS`.  ``code`` is always ``"header_unclear"``: the caller
    declares it (:meth:`Workbook.declare_region
    <xlsx_agent.workbook.Workbook.declare_region>`)."""

    address: str
    reason: str
    code: str = field(default="header_unclear")


class Regions:
    """A workbook's regions: made by :class:`~xlsx_agent.workbook.Workbook`, which holds
    one, and read through it (the tables and names once, the cells streamed by the R3
    reader)."""

    def __init__(self, workbook: "Workbook") -> None:
        self.workbook = workbook
        self._tables: list[Region] | None = None
        self._names: list[DefinedName] | None = None
        self._declared: dict[str, Region] = {}
        self._detected: dict[tuple, Region] = {}
        self._keys: dict[tuple, list[str]] = {}

    # -- what the workbook holds -----------------------------------------------------------

    @property
    def tables(self) -> list[Region]:
        if self._tables is None:
            self._tables = [self._table_region(part) for part in self.workbook._table_parts()
                            if part.area is not None and not _broken(part)]
        return list(self._tables)

    @property
    def names(self) -> list[DefinedName]:
        if self._names is None:
            self._names = self.workbook._defined_names()
        return list(self._names)

    @property
    def declared(self) -> list[Region]:
        return list(self._declared.values())

    def problems(self) -> list[Problem]:
        """What is odd about the tables and names, by code: each table's flags, a table
        the workbook cannot read (``bad_table``, at its sheet), a name whose
        ``localSheetId`` names no sheet (``bad_scope``), and two tables or workbook-wide
        names with one name, case aside (``duplicate_name``, at the second's address)."""
        out: list[Problem] = []
        for part in self.workbook._table_parts():
            out.extend(part.flags)
            if part.area is not None and _broken(part):
                out.append(Problem("bad_table", quote_sheet(part.sheet) if part.sheet else None))
        for name in self.names:
            if name.bad_scope is not None:
                out.append(Problem("bad_scope", None))
        seen: set[tuple[str | None, str]] = set()
        for name, scope in [(t.name, None) for t in self.tables] + \
                [(n.name, n.scope) for n in self.names]:
            key = (scope.casefold() if scope else None, name.casefold())
            if key in seen and is_address_name(name):
                written = TableAddress(name=name, scope=SheetRef(name=scope) if scope else None)
                out.append(Problem("duplicate_name", written.text))
            seen.add(key)
        return list(dict.fromkeys(out))

    def _table_region(self, part: TablePart) -> Region:
        sheet = self.workbook.sheet(part.sheet)
        return Region(kind="table", name=part.name, sheet=part.sheet, area=part.area,
                      columns=part.columns, header_rows=part.header_rows,
                      totals_rows=part.totals_rows, sheet_id=sheet.id, flags=part.flags,
                      id=part.id, object_name=part.object_name, auto_filter=part.auto_filter,
                      totals=part.totals, table_type=part.kind, workbook=self.workbook)

    # -- finding a region by address -------------------------------------------------------

    def region(self, address: "str | TableAddress") -> Region:
        """The table, name or region a ``t:`` address names (its column or row, if any,
        aside)."""
        address = _table_address(address)
        if address.region is not None:
            return self.detect(address.region)
        if address.scope is not None:
            return self._scoped(address)
        folded = address.name.casefold()
        found: list[tuple[str, object]] = []
        found += [("table", t) for t in self.tables if t.name.casefold() == folded]
        found += [("name", n) for n in self.names
                  if n.scope is None and n.bad_scope is None and n.name.casefold() == folded]
        if folded in self._declared:
            found.append(("declared", self._declared[folded]))
        if len(found) > 1:
            kinds = sorted({kind for kind, _ in found})
            raise RegionError("ambiguous_name", f"t:{address.name} names more than one thing "
                              f"in the workbook ({', '.join(kinds)}); it cannot be resolved",
                              address=address.whole.text)
        if not found:
            local = [n for n in self.names if n.scope is not None and n.name.casefold() == folded
                     and n.address is not None]
            if local:
                options = [n.address.text for n in local]
                raise RegionError("scoped_name", f"t:{address.name} is local to a sheet; it is "
                                  f"written with its sheet, as {options[0]}",
                                  address=address.whole.text, valid_options=options)
            known = self.known()
            raise NameNotFound(f"no table, name or region t:{address.name}; there are "
                               f"{known[:50]}", what="table", name=address.name,
                               candidates=known,
                               valid_options=nearest(f"t:{address.name}", known))
        kind, thing = found[0]
        if kind == "name":
            return self._name_region(thing)
        return thing

    def _scoped(self, address: TableAddress) -> Region:
        sheet = self.workbook.sheet_of(address.scope)
        folded = address.name.casefold()
        found = [n for n in self.names
                 if n.scope is not None and n.scope.casefold() == sheet.name.casefold()
                 and n.name.casefold() == folded]
        if len(found) > 1:
            raise RegionError("ambiguous_name", f"{address.whole.text} names more than one "
                              "defined name; it cannot be resolved", address=address.whole.text)
        if not found:
            local = [n.address.text for n in self.names if n.scope is not None
                     and n.scope.casefold() == sheet.name.casefold() and n.address is not None]
            known = local + self.known()
            raise NameNotFound(f"no name {address.name!r} local to sheet "
                               f"{quote_sheet(sheet.name)}", what="table", name=address.name,
                               candidates=known,
                               valid_options=nearest(address.whole.text, known))
        return self._name_region(found[0])

    def known(self) -> list[str]:
        """Every table, workbook-wide name over a range and declared region, as ``t:``
        addresses, in that order (sheet-local names after them)."""
        out = [t.address.text for t in self.tables if is_address_name(t.name)]
        names = [n for n in self.names if n.is_range and n.address is not None]
        out += [n.address.text for n in names if n.scope is None]
        out += [r.address.text for r in self.declared]
        out += [n.address.text for n in names if n.scope is not None]
        return list(dict.fromkeys(out))

    def _name_region(self, name: DefinedName) -> Region:
        if name.range is None:
            written = name.address.text if name.address is not None else "the name"
            raise RegionError("not_a_range", f"{written} is a defined name that is not a range "
                              f"({name.reason}); it is listed, never evaluated",
                              address=written, reason=name.reason)
        sheet = self.workbook.sheet_of(name.range.sheet)
        area = name.range.clipped(self.workbook.used_range(sheet)) \
            if sheet.type in ("worksheet", "macrosheet") else name.range.area
        if area is None:  # whole columns or rows of an empty sheet: the first row
            area = Area(name.range.area.first, Cell(1, name.range.area.last.column)) \
                if name.range.span == "columns" else Area(name.range.area.first,
                                                          Cell(name.range.area.first.row, 1))
        columns = tuple(f"Column {column_letters(c)}"
                        for c in range(area.first.column, area.last.column + 1))
        return Region(kind="name", name=name.name, sheet=sheet.name, area=area, columns=columns,
                      header_rows=0, scope=name.scope, hidden=name.hidden, sheet_id=sheet.id,
                      workbook=self.workbook)

    # -- declared regions ------------------------------------------------------------------

    def declare(self, name: str, where: "str | RangeAddress", *, header_rows: int = 1,
                headers: Sequence[str] | None = None, key: str | None = None) -> Region:
        if not is_name(name):
            raise RegionError("bad_name", "a region's name is a name as ECMA-376 defines one: "
                              "a letter, _ or \\ first, then letters, digits, _ and .; not "
                              "TRUE, FALSE or a cell reference; at most 255 characters",
                              valid_options=["Returns", "Stock_2026"])
        folded = name.casefold()
        taken = [t.address.text for t in self.tables if t.name.casefold() == folded]
        taken += [n.address.text for n in self.names if n.address is not None
                  and n.scope is None and n.name.casefold() == folded]
        if taken:
            raise RegionError("name_taken", f"{taken[0]} is already a table or defined name; "
                              "a region takes a name of its own", address=taken[0])
        if isinstance(header_rows, bool) or not isinstance(header_rows, int) or header_rows < 0:
            raise RegionError("bad_headers", "header_rows is a whole number from 0")
        if isinstance(where, str):
            where = parse_address(where)
        if not isinstance(where, RangeAddress):
            raise RegionError("bad_name", "a region is declared over a cell or range, as "
                              "Returns!A3:F200", valid_options=["Returns!A3:F200"])
        sheet = self.workbook._cell_sheet(self.workbook.sheet_of(where))
        area = where.clipped(self.workbook.used_range(sheet))
        if area is None:
            raise RegionError("no_header", f"{where.text} holds no cells", address=where.text)
        height = area.last.row - area.first.row + 1
        width = area.last.column - area.first.column + 1
        if header_rows > height:
            raise RegionError("bad_headers", f"header_rows is {header_rows} for a region "
                              f"{height} rows high")
        flags: list[Problem] = []
        written = quote_sheet(sheet.name)
        if headers is not None:
            headers = list(headers)
            if len(headers) != width or not all(isinstance(h, str) and h for h in headers):
                raise RegionError("bad_headers", f"headers names each of the region's {width} "
                                  "columns with a non-empty string")
            if len({h.casefold() for h in headers}) != width:
                raise RegionError("bad_headers", "headers are unique, case aside")
            columns = headers
        elif header_rows > 1:
            raise RegionError("headers_required", f"a region with {header_rows} header rows "
                              "is declared with its header names (headers=[...]): a "
                              "multi-row header is never read as names")
        elif header_rows == 1:
            row = next(iter(self.workbook.read(RangeAddress(
                SheetRef(name=sheet.name), Area(area.first, Cell(area.first.row,
                                                                  area.last.column))))))
            columns = _header_names(row, area.first.column, flags, written)
        else:
            columns = [f"Column {column_letters(c)}"
                       for c in range(area.first.column, area.last.column + 1)]
        region = Region(kind="declared", name=name, sheet=sheet.name, area=area,
                        columns=tuple(columns), header_rows=header_rows, sheet_id=sheet.id,
                        flags=tuple(flags), workbook=self.workbook)
        if key is not None:
            region = replace(region, key=region.columns[region.column_index(key)])
        self._declared[folded] = region
        return region

    # -- detected regions ------------------------------------------------------------------

    def detect(self, where: "str | RangeAddress | TableAddress") -> Region:
        """The region whose single-row header starts at a cell (``t:Stock!A1`` or
        ``Stock!A1``), detected; :class:`RegionError` ``header_unclear`` (with its reason)
        or ``no_header`` when there is no simple header there.

        The header row runs right from the cell while the header cell or the cell under it
        holds something; the data rows run down from the row under the header to the row
        before the first row blank across those columns (or the used range's end).  The
        header is unclear, and nothing is detected, when a merged area touches the header
        row, the row above it or the row below it within those columns
        (``merged_header``); when the row directly above holds something within them
        (``rows_above``: a header of two rows, or a title, may start there); when a header
        cell holds something other than text (``header_not_text``); or when the first
        data row is all text and some column's later rows are all something else
        (``multi_row_header``: a second header row)."""
        if isinstance(where, str):
            where = parse_address(where)
        if isinstance(where, TableAddress):
            if where.region is None:
                raise RegionError("no_header", "a detected region is named by its header's "
                                  "first cell, as t:Stock!A1", valid_options=["t:Stock!A1"])
            where = where.region
        if not isinstance(where, RangeAddress) or where.kind != "cell":
            raise RegionError("no_header", "a detected region is named by one cell, its "
                              "header's first", valid_options=["t:Stock!A1"])
        sheet = self.workbook._cell_sheet(self.workbook.sheet_of(where))
        cache = (sheet.part or sheet.id, where.cell)
        if cache not in self._detected:
            self._detected[cache] = self._detect(sheet, where.cell)
        return self._detected[cache]

    def _detect(self, sheet: "Sheet", cell: Cell) -> Region:
        written = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
        address = f"t:{written}!{cell.ref}"
        used = self.workbook.used_range(sheet)
        if used is None or not used.contains(cell):
            raise RegionError("no_header", f"{address}: nothing is at the header cell",
                              address=address)
        top = max(cell.row - 1, 1)
        area = Area(Cell(top, cell.column), Cell(max(used.last.row, cell.row), used.last.column))
        rows = self.workbook.read(RangeAddress(SheetRef(name=sheet.name), area))
        above = next(rows) if top < cell.row else None
        header = next(rows)
        below = next(rows, None)
        if header[0].t == "blank":
            raise RegionError("no_header", f"{address}: nothing is at the header cell",
                              address=address)
        width = 0
        for offset, value in enumerate(header):
            under = below[offset] if below is not None else None
            if value.t == "blank" and (under is None or under.t == "blank"):
                break
            width = offset + 1
        right = cell.column + width - 1

        def unclear(reason: str) -> RegionError:
            return RegionError("header_unclear", f"{address}: the header is not a simple "
                               f"single row ({reason}); declare the region with "
                               "declare_region and its header names", address=address,
                               reason=reason)

        span = Area(Cell(max(cell.row - 1, 1), cell.column), Cell(cell.row + 1, right))
        if any(_overlaps(merge, span) for merge in self.workbook.merges(sheet)):
            raise unclear("merged_header")
        if above is not None and any(v.t != "blank" for v in above[:width]):
            raise unclear("rows_above")
        if any(v.t not in ("text", "blank") for v in header[:width]):
            raise unclear("header_not_text")
        data: list[list[Value]] = []
        if below is not None and any(v.t != "blank" for v in below[:width]):
            data.append(below[:width])
            for row in rows:
                if all(v.t == "blank" for v in row[:width]):
                    break
                data.append(row[:width])
        if len(data) > 1 and _second_header(data):
            raise unclear("multi_row_header")
        flags: list[Problem] = []
        columns = _header_names(header[:width], cell.column, flags, written)
        bottom = cell.row + len(data)
        return Region(kind="detected", name=None, sheet=sheet.name,
                      area=Area(cell, Cell(bottom, right)), columns=tuple(columns),
                      header_rows=1, sheet_id=sheet.id, flags=tuple(flags),
                      workbook=self.workbook)

    def find(self, sheet: "Sheet | str | SheetRef") -> list["Region | UnclearRegion"]:
        """The blocks of a sheet outside its tables, each detected (a :class:`Region`) or
        not (an :class:`UnclearRegion`, with the reason): a block's header is the first
        row, from the top, holding something outside the tables and the blocks found
        above it, starting at its first such cell; further blocks in that row start after
        a blank column.  A block whose header is unclear is taken to run to the next row
        blank across the used range."""
        sheet = self.workbook._cell_sheet(sheet if not isinstance(sheet, str)
                                          else self.workbook.sheet(sheet))
        used = self.workbook.used_range(sheet)
        if used is None:
            return []
        taken = [t.area for t in self.tables if t.sheet.casefold() == sheet.name.casefold()]
        out: list[Region | UnclearRegion] = []
        written = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
        row_number = used.first.row
        rows = self._rows(sheet, used, row_number)
        while True:
            row = next(rows, None)
            if row is None:
                break
            row_number = row[0].row
            starts = [v for v in row if v.t != "blank"
                      and not any(a.contains(Cell(v.row, v.column)) for a in taken)]
            if not starts:
                continue
            bottom = row_number
            column = 0
            for value in starts:
                if value.column <= column:
                    continue
                try:
                    region = self.detect(RangeAddress(SheetRef(name=sheet.name),
                                                      Area(Cell(value.row, value.column),
                                                           Cell(value.row, value.column))))
                except RegionError as error:
                    if error.code != "header_unclear":
                        raise
                    out.append(UnclearRegion(f"t:{written}!{_ref(value)}", error.reason))
                    column = _block_end(row, value.column)
                    bottom = max(bottom, self._blank_after(sheet, used, row_number))
                    continue
                out.append(region)
                taken.append(region.area)
                column = region.area.last.column
                bottom = max(bottom, region.area.last.row)
            if bottom > row_number:
                rows = self._rows(sheet, used, bottom + 1)
        return out

    def _rows(self, sheet: "Sheet", used: Area, top: int) -> Iterator[list[Value]]:
        if top > used.last.row:
            return iter(())
        return self.workbook.read(RangeAddress(SheetRef(name=sheet.name),
                                               Area(Cell(top, used.first.column), used.last)))

    def _blank_after(self, sheet: "Sheet", used: Area, row: int) -> int:
        """The last row before the first row, after ``row``, blank across the used range."""
        last = row
        for values in self._rows(sheet, used, row + 1):
            if all(v.t == "blank" for v in values):
                break
            last = values[0].row
        return last

    # -- resolving -------------------------------------------------------------------------

    def resolve(self, address: "str | TableAddress", *, key: str | None = None) -> Selection:
        address = _table_address(address)
        region = self.region(address)
        kind = address.kind
        sheet = region.sheet_ref
        whole = region.address
        if address.key is None and address.position is None:
            if address.column is None:
                data = region.data
                return Selection(whole.text, kind, region,
                                 RangeAddress(sheet, data) if data is not None else None)
            name = region.columns[region.column_index(address.column)]
            area = region.column_area(name)
            text = replace(whole, column=name).text
            return Selection(text, "column", region,
                             RangeAddress(sheet, area) if area is not None else None,
                             column=name)
        column = None
        if address.column is not None:
            column = region.columns[region.column_index(address.column)]
        if address.position is not None:
            position = address.position
            if position > region.rows:
                options = [replace(whole, position=region.rows).text] if region.rows else []
                raise RegionError("position_out_of_range", f"{whole.text} has {region.rows} "
                                  f"data rows; there is no row #{position}",
                                  address=whole.text, valid_options=options)
        else:
            position = self._find_key(region, address.key, key)
        data = region.data
        row = data.first.row + position - 1
        if column is not None:
            c = region.area.first.column + region.column_index(column)
            area = Area(Cell(row, c), Cell(row, c))
        else:
            area = Area(Cell(row, data.first.column), Cell(row, data.last.column))
        text = replace(whole, key=address.key, position=address.position, column=column).text
        return Selection(text, "value" if column is not None else "row", region,
                         RangeAddress(sheet, area), row=position, column=column)

    def key_column(self, region: Region, key: str | None = None) -> str:
        """The key column of ``region``: ``key`` (a column name) when given, else the
        region's declared key, else the one column :meth:`key_columns` finds; else
        :class:`RegionError` ``no_key_column`` with the candidates as ``valid_options``."""
        if key is not None:
            return region.columns[region.column_index(key)]
        if region.key is not None:
            return region.key
        candidates = self.key_columns(region)
        if len(candidates) == 1:
            return candidates[0]
        why = "no column is" if not candidates else f"{len(candidates)} columns are"
        raise RegionError("no_key_column", f"{region.address.text} has no key column named, "
                          f"and {why} unique, never blank, text or whole numbers; name one "
                          "with key=", address=region.address.text, valid_options=candidates)

    def key_columns(self, region: Region) -> list[str]:
        """The columns of ``region`` that can be its key: every data cell holds text or a
        whole number, and no two hold the same key (as :meth:`resolve` matches keys).  A
        region with no data rows has none."""
        cache = (region.kind, region.sheet, region.area, region.columns)
        if cache in self._keys:
            return list(self._keys[cache])
        data = region.data
        out: list[str] = []
        if data is not None:
            width = data.last.column - data.first.column + 1
            seen: list[set[str] | None] = [set() for _ in range(width)]
            for row in self.workbook.read(RangeAddress(region.sheet_ref, data)):
                for index, value in enumerate(row):
                    keys = seen[index]
                    if keys is None:
                        continue
                    text = key_text(value)
                    if text is None or text in keys:
                        seen[index] = None
                    else:
                        keys.add(text)
            out = [region.columns[i] for i, keys in enumerate(seen) if keys is not None]
        self._keys[cache] = out
        return list(out)

    def _find_key(self, region: Region, wanted: str, key: str | None) -> int:
        column = self.key_column(region, key)
        area = region.column_area(column)
        found: list[int] = []
        if area is not None:
            for position, row in enumerate(self.workbook.read(RangeAddress(region.sheet_ref,
                                                                           area)), 1):
                if key_text(row[0]) == wanted:
                    found.append(position)
        whole = region.address
        if not found:
            raise NameNotFound(f"no row of {whole.text} has that key in column {column!r}",
                               what="key", name=wanted)
        if len(found) > 1:
            options = [replace(whole, position=p).text for p in found[:50]]
            raise RegionError("duplicate_key", f"{len(found)} rows of {whole.text} have that "
                              f"key in column {column!r}; name the row by its position",
                              address=whole.text, valid_options=options)
        return found[0]


def key_text(value: Value) -> str | None:
    """A cell as a key: its text, or a whole number's digits (``42`` for 42 and 42.0);
    ``None`` for anything else (a blank, a fraction, a date, a boolean, an error)."""
    if value.t == "text":
        return value.v if value.v else None
    if value.t == "number" and not isinstance(value.v, bool):
        number = value.v
        if isinstance(number, int):
            return str(number)
        if isinstance(number, float) and number.is_integer():
            return str(int(number))
    return None


def _header_names(row: Sequence[Value], first_column: int, flags: list[Problem],
                  written: str) -> list[str]:
    """A header row's text as column names (see :func:`~.tables.name_columns`); a cell
    that is not text counts as blank."""
    names = [v.v if v.t == "text" and isinstance(v.v, str) else "" for v in row]
    addresses = [f"{written}!{_ref(v)}" for v in row]
    return name_columns(names, first_column, flags, written, blank="blank_header",
                        duplicate="duplicate_header", addresses=addresses)


def _second_header(data: list[list[Value]]) -> bool:
    """The first data row looks like a second header row: all of it text (or blank), and
    some column whose first value is text holds no text further down, yet holds something."""
    first = data[0]
    if any(v.t not in ("text", "blank") for v in first):
        return False
    for index, value in enumerate(first):
        if value.t != "text":
            continue
        later = [row[index].t for row in data[1:] if row[index].t != "blank"]
        if later and "text" not in later:
            return True
    return False


def _overlaps(one: Area, two: Area) -> bool:
    return not (one.last.row < two.first.row or two.last.row < one.first.row
                or one.last.column < two.first.column or two.last.column < one.first.column)


def _block_end(row: Sequence[Value], start: int) -> int:
    """The last column of the run of non-blank cells in ``row`` from column ``start``."""
    end = start
    for value in row:
        if value.column < start:
            continue
        if value.t == "blank":
            break
        end = value.column
    return end


def _broken(part: TablePart) -> bool:
    return any(flag.code == "bad_ref" for flag in part.flags)


def _ref(value: Value) -> str:
    return value.ref


def _table_address(address: "str | TableAddress") -> TableAddress:
    if isinstance(address, str):
        address = TableAddress.parse(address)
    if not isinstance(address, TableAddress):
        raise ValueError("a t: address is expected, as t:Products")
    return address
