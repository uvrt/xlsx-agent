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

**What a key is** (:func:`key_text`).  A text cell is its text, exactly: ``007`` is the
text ``007`` and never the number 7, and text that looks like a date (``2026-01-05``) is
that text.  A whole number is its digits, so the text ``7`` and the number 7 are one key
(two rows holding them share it).  A blank cell, a fraction, a date, a time, a boolean and
an error are no key: a date's text depends on its format and the locale, so a date is
never matched by text.  A row whose key cell is no key cannot be named by key; a column
with one is no key column unless the caller names it (and then that row has no key
address: :meth:`Regions.expand` refuses it as ``no_key``).

**Key index.**  A region's key column is read once into a :class:`KeyIndex` (each row's
key by position, and which keys two rows share), kept until the locale changes: a row by
key, a key set and :meth:`Regions.expand` look keys up there, never by reading the column
again.

**Declared tables** (:mod:`xlsx_agent.declare`): tables the caller declares from its own
configuration are declared regions too, with column ids (``Region.headers`` keeps each
header, which finds the column as well), descriptions and a ``shape``.  One over several
sheets, or with columns hidden from the model, is read as one rectangle of its own
(:class:`Span`): a range on its own sheet (``t:<name>``) is read sheet by sheet, each
value given in the table's place with its own cell kept, so every address a result gives
is a real cell's.  Its blocks are not detected again (:meth:`Regions.find`), and a region
detected or declared over a hidden column's cells is refused (``declared``), as is a
defined name over them.  An Excel table it covers any cell of is replaced by it
(:meth:`Regions.replaced_by`): left out of :attr:`Regions.tables`, and ``t:<the Excel
table>`` refused toward it (``declared``, the declared table as ``valid_options``).

**Key-value blocks.**  A block of two columns whose first row holds something other than
text (so no header is detected) and whose first column is a different text on each of
two or more rows is detected as a key-value region (``shape`` ``"key_value"``): no header
row, its columns ``field`` and ``value``, its key ``field``
(``t:Run!A1/report_date/value``).

**Titles and captions.**  A block of one cell with nothing under it (a title, a caption)
is no region: :meth:`Regions.find` leaves it out, so ``describe`` does not list it, and it
stays a cell that ``find_text`` and a read reach.  When it is text and lies over a detected
region's columns two rows above its header (one blank row between), it is that region's
``title``.  A region of one column is a region whenever it has data rows.

**Selectors and key sets** (:mod:`xlsx_agent.address`).  ``t:Products/*/Price`` resolves
as ``t:Products[Price]``, ``t:Products/*`` and ``t:Products/*/*`` as the table, and
``t:Products/SKU-0042/*`` as the row, each keeping its own text.  A key set
(``t:Products/{SKU-0042,SKU-0043}/Price``) resolves to its rows' positions, in the set's
order (:attr:`Selection.positions`), each key looked up as a single key is: one unknown or
shared refuses the whole set.  :meth:`Regions.expand` gives what any of these covers as
one address by key per cell (or per row).
"""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Iterator, Sequence

from ooxml_edit.charts.workbook import Area, Cell, column_letters

from ..address import (NameNotFound, RangeAddress, SheetRef, TableAddress,
                       is_address_name, is_name, nearest, parse_address, quote_sheet)
from .names import DefinedName
from .tables import TablePart, name_columns
from .types import Problem, Value

if TYPE_CHECKING:  # pragma: no cover
    from ..clean import Clean
    from ..profile import Column, ColumnStats
    from ..query import QueryResult
    from ..workbook import Sheet, Workbook

#: What a :class:`Region` is.
REGION_KINDS = ("table", "name", "declared", "detected")

#: The names of a key-value region's first two columns (the rest are named by letter):
#: ``t:run/report_date/value``.
KEY_VALUE_COLUMNS = ("field", "value")

#: How wide a block detected as key-value is: a field and a value.  A wider block whose
#: first row holds numbers is as likely a header of years over a table, so it is never
#: detected as one; a key-value table with more columns (a confidence, a source...) is
#: declared (:mod:`xlsx_agent.declare`).
KEY_VALUE_WIDTH = 2

#: The most rows a block detected as key-value has.
KEY_VALUE_ROWS = 200

#: Why a header is not detected (:attr:`RegionError.reason` of ``header_unclear``).
UNCLEAR_REASONS = ("merged_header", "rows_above", "multi_row_header", "header_not_text")


class RegionError(ValueError):
    """A ``t:`` address, or a region, the workbook cannot give: ``code`` says why.

    ``not_a_range`` (a defined name that is a formula or a constant), ``ambiguous_name``,
    ``scoped_name`` (``t:Name`` for a name local to a sheet), ``header_unclear`` (with
    ``reason``, one of :data:`UNCLEAR_REASONS`), ``no_header`` (no header at the cell),
    ``no_key_column``, ``duplicate_key``, ``no_key`` (a row whose key cell is no key,
    :meth:`Regions.expand`), ``no_key_set`` (a key set of a dataset's rows,
    :mod:`xlsx_agent.dataset`), ``position_out_of_range``, ``bad_table``,
    ``name_taken``, ``bad_name``, ``headers_required``, ``bad_headers``, ``declared`` (a
    detected region or declared range in a table declared from configuration), and as a source of
    rows (:func:`~xlsx_agent.dataset.table_dataset`) ``not_rows`` (one value) and
    ``too_many_cells`` (more values than the caller's limit).  The message
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

    A table declared from the caller's configuration (:mod:`xlsx_agent.declare`) also has
    ``headers`` (each column's header text, where its name is an id the caller gave: a
    column is found by either), ``notes`` (each column's description, or ``None``),
    ``about`` (the table's) and ``shape`` (``"rows"``, or ``"key_value"``: a field column
    and a value column, addressed ``t:run/<field>/value``).  One that spans several sheets,
    or has columns hidden from the model, is read as one rectangle of its own: ``span``
    (a :class:`Span`) maps its rows and columns to the cells they are, its ``sheet`` is its
    own (``t:<name>``, a name no sheet can have), and every value read from it carries its
    cell's own address (:attr:`Value.origin <xlsx_agent.read.types.Value.origin>`);
    :attr:`where` gives its cells as ranges of the sheets.

    A detected region with a title cell over it (one blank row between) has ``title``, its
    text; :attr:`caption` is true of a region that is one cell and no data rows, a title or
    caption that :meth:`Regions.find` leaves out.

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
    span: "Span | None" = None
    headers: tuple[str, ...] = ()
    notes: tuple[str | None, ...] = ()
    about: str | None = None
    shape: str = "rows"
    title: str | None = None
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
        """All of it, as a range address (on its own sheet, for a :attr:`span`: see
        :attr:`where`)."""
        return RangeAddress(self.sheet_ref, self.area)

    @property
    def where(self) -> str:
        """All of it as the sheets' ranges, as text: :attr:`range`'s, or each sheet's of a
        :attr:`span`, joined by ``,`` (Excel's union)."""
        if self.span is None:
            return self.range.text
        return ",".join(RangeAddress(SheetRef(name=sheet), area).text
                        for sheet, area in self.span.extent())

    def located(self, area: Area | None) -> str | None:
        """Cells of the region (in its own rows and columns) as the sheets' ranges, as
        text: ``Stock!C2:C40``; for a :attr:`span`, each sheet's, joined by ``,``."""
        if area is None:
            return None
        if self.span is None:
            return RangeAddress(self.sheet_ref, area).text
        return ",".join(r.text for r in self.span.locate(area))

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

    @property
    def caption(self) -> bool:
        """Whether this is a lone cell, a title or a caption: one cell detected as a header,
        with no data rows."""
        return (self.kind == "detected" and self.shape == "rows" and self.rows == 0
                and self.area.first == self.area.last)

    def column_index(self, name: str) -> int:
        """The position (from 0) of the column called ``name``, matched case aside;
        :class:`~xlsx_agent.address.NameNotFound` with the nearest column names."""
        folded = name.casefold()
        for index, column in enumerate(self.columns):
            if column.casefold() == folded:
                return index
        for index, header in enumerate(self.headers):
            if header.casefold() == folded:
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
    ``"table"`` or ``"region"`` (the whole of it), ``"column"``, ``"row"`` or ``"value"``,
    ``"rows"`` or ``"values"`` (a key set); ``region`` the :class:`Region`; ``range`` the
    data cells selected (header and totals rows excluded; ``None`` when the region has no
    data rows, and for a key set); ``row`` the data row's position (from 1) for a row or a
    value; ``column`` the column's name as the region spells it; ``positions`` a key set's
    rows, in its order (:attr:`areas` gives their cells).  For a declared table read as
    one (a :attr:`Region.span`), ``range`` is in the table's own rectangle: reading it
    gives each cell's own address, and :meth:`Region.located` gives the sheets'."""

    address: str
    kind: str
    region: Region
    range: RangeAddress | None
    row: int | None = None
    column: str | None = None
    positions: tuple[int, ...] = ()

    @property
    def areas(self) -> list[RangeAddress]:
        """The cells selected, one range per row of a key set (in the set's order), else
        :attr:`range` alone (nothing when it is ``None``)."""
        if not self.positions:
            return [self.range] if self.range is not None else []
        region, data = self.region, self.region.data
        if self.column is not None:
            left = right = region.area.first.column + region.column_index(self.column)
        else:
            left, right = data.first.column, data.last.column
        return [RangeAddress(region.sheet_ref,
                             Area(Cell(data.first.row + p - 1, left),
                                  Cell(data.first.row + p - 1, right)))
                for p in self.positions]


@dataclass(frozen=True)
class Part:
    """One sheet of a declared table read as one (:class:`Span`): ``sheet`` its name as
    the workbook spells it, ``header`` its header row (``None`` without one), ``first``
    and ``last`` its data rows (``last`` is ``first - 1`` when it has none)."""

    sheet: str
    sheet_id: int | None
    header: int | None
    first: int
    last: int

    @property
    def rows(self) -> int:
        return self.last - self.first + 1


@dataclass(frozen=True)
class Span:
    """How a declared table read as one rectangle of its own maps to the cells it is: a
    table over several sheets (rows in the sheets' order), or one with columns hidden
    from the model.  In its own rectangle, row 1 is the header row (the first sheet's;
    ``header_rows`` 0: none), the data rows follow sheet by sheet, and column ``n`` is the
    sheet column ``columns[n - 1]``; ``left`` to ``right`` are the sheet columns declared,
    ``hidden`` the ones hidden from the model among them; ``clean`` each of its own
    columns' cleaning rules (``None``: none; empty: no column has any), applied to the data
    rows as they are read (:mod:`xlsx_agent.clean`).  ``name`` is the sheet name its
    region reads through (``t:<name>``: no sheet can be called that, since a sheet's
    name never holds ``:``)."""

    name: str
    parts: tuple[Part, ...]
    columns: tuple[int, ...]
    header_rows: int
    left: int
    right: int
    hidden: tuple[int, ...] = ()
    clean: "tuple[Clean | None, ...]" = ()

    @property
    def rows(self) -> int:
        """How many data rows, every sheet's."""
        return sum(part.rows for part in self.parts)

    @property
    def area(self) -> Area:
        """Its own rectangle: the header row, then every data row; ``columns`` wide."""
        return Area(Cell(1, 1), Cell(max(self.header_rows + self.rows, 1), len(self.columns)))

    def runs(self, top: int, bottom: int) -> Iterator[tuple[Part, int, int, int]]:
        """Its own rows ``top`` to ``bottom`` as each sheet's: the part, its first and last
        sheet rows, and the first of its own rows they are."""
        if self.header_rows and top <= 1 <= bottom:
            first = self.parts[0]
            yield first, first.header, first.header, 1
        offset = self.header_rows
        for part in self.parts:
            low, high = max(top, offset + 1), min(bottom, offset + part.rows)
            if low <= high:
                yield part, part.first + low - offset - 1, part.first + high - offset - 1, low
            offset += part.rows

    def locate(self, area: Area) -> list[RangeAddress]:
        """Cells of its own rectangle as the sheets' ranges, sheet by sheet; a run of
        columns broken by a hidden one is two ranges."""
        real = [self.columns[c - 1] for c in range(area.first.column, area.last.column + 1)]
        spans: list[list[int]] = []
        for column in real:
            if spans and spans[-1][1] + 1 == column:
                spans[-1][1] = column
            else:
                spans.append([column, column])
        return [RangeAddress(SheetRef(name=part.sheet), Area(Cell(top, a), Cell(bottom, b)))
                for part, top, bottom, _ in self.runs(area.first.row, area.last.row)
                for a, b in spans]

    def areas(self) -> list[tuple[str, Area]]:
        """Each sheet's cells of the table, header row to last data row, every column
        declared (hidden ones too)."""
        out = []
        for part in self.parts:
            top = part.header if part.header is not None else part.first
            bottom = max(part.last, top)
            out.append((part.sheet, Area(Cell(top, self.left), Cell(bottom, self.right))))
        return out

    def extent(self) -> list[tuple[str, Area]]:
        """Each sheet's cells of the table as the model sees them: header row to last data
        row, from its first column to its last (a hidden column at either end left
        out)."""
        low, high = min(self.columns), max(self.columns)
        return [(sheet, Area(Cell(area.first.row, low), Cell(area.last.row, high)))
                for sheet, area in self.areas()]

    def withheld(self) -> list[tuple[str, Area]]:
        """The cells of the columns hidden from the model, sheet by sheet: never shown."""
        return [(sheet, Area(Cell(area.first.row, column), Cell(area.last.row, column)))
                for sheet, area in self.areas() for column in self.hidden]

    def read(self, workbook: "Workbook", area: Area, *, formats: bool = False,
             columns: "set[int] | None" = None) -> Iterator[list[Value]]:
        """The rows of ``area`` (its own rows and columns), each read from its sheet and
        given in its own place, its cell's address kept (:attr:`Value.origin
        <xlsx_agent.read.types.Value.origin>`).  ``columns``: its own columns to type."""
        own = range(area.first.column, area.last.column + 1)
        real = [self.columns[c - 1] for c in own]
        left, right = min(real), max(real)
        typed = None if columns is None else \
            {self.columns[c - 1] for c in columns if 1 <= c <= len(self.columns)}
        name = quote_sheet(self.name)
        rules = [(n, column, self.clean[column - 1]) for n, column in enumerate(own)
                 if self.clean and self.clean[column - 1] is not None]
        if rules:
            from ..clean import clean_value
        for part, top, bottom, first in self.runs(area.first.row, area.last.row):
            cells = RangeAddress(SheetRef(name=part.sheet), Area(Cell(top, left),
                                                                 Cell(bottom, right)))
            for row in workbook._read(cells, formats=formats, columns=typed):
                at = first + row[0].row - top
                out = []
                for column, sheet_column in zip(own, real):
                    value = row[sheet_column - left]
                    value.origin = (value.sheet, value.row, value.column)
                    value.sheet, value.row, value.column = name, at, column
                    out.append(value)
                if rules and at > self.header_rows:
                    for n, _, rule in rules:
                        clean_value(out[n], rule, workbook)
                yield out


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
        self._indexes: dict[tuple, KeyIndex] = {}
        self._declared_keys: dict[str, str] = {}      # region's address, folded -> column
        #: The tables declared from the caller's configuration (:mod:`xlsx_agent.declare`),
        #: by name folded; those read as one rectangle of their own by their sheet name.
        self._configured: dict[str, Region] = {}
        self._spans: dict[str, Span] = {}
        #: The tables being declared again (by name folded): their hidden cells are read
        #: while the new declaration is checked (:meth:`redeclaring`).
        self._redeclaring: frozenset[str] = frozenset()

    def forget_cells(self) -> None:
        """Forget what was found from the cells (detected regions, key columns, key
        indexes): the workbook's locale changed.  The tables, names and declared regions
        stay."""
        self._detected.clear()
        self._keys.clear()
        self._indexes.clear()

    # -- what the workbook holds -----------------------------------------------------------

    @property
    def tables(self) -> list[Region]:
        """The Excel tables, less those a declared table covers (:meth:`replaced_by`): what
        the workbook shows of its tables."""
        return [t for t in self.excel_tables if self.replaced_by(t) is None]

    @property
    def excel_tables(self) -> list[Region]:
        """Every Excel table the workbook can read, those a declared table covers too: their
        names stay taken."""
        if self._tables is None:
            self._tables = [self._table_region(part) for part in self.workbook._table_parts()
                            if part.area is not None and not _broken(part)]
        return list(self._tables)

    def replaced_by(self, table: Region) -> Region | None:
        """The declared table that covers any cell of the Excel table ``table``, and so
        replaces it: ``describe`` lists the declared table alone, and ``t:<table>`` is
        refused toward it (``declared``), so a column the declared table hides from the
        model is never reached through the Excel table's name."""
        for region, area in self.table_areas(table.sheet):
            if _overlaps(area, table.area):
                return region
        return None

    @property
    def names(self) -> list[DefinedName]:
        if self._names is None:
            self._names = self.workbook._defined_names()
        return list(self._names)

    @property
    def declared(self) -> list[Region]:
        return list(self._declared.values())

    # -- tables declared from configuration -------------------------------------------------

    def keep_table(self, region: Region) -> Region:
        """Keep a table :mod:`xlsx_agent.declare` made, usable as ``t:<name>`` from then on
        (declaring its name again replaces it, and whatever region had it)."""
        folded = region.name.casefold()
        old = self._declared.pop(folded, None)
        if old is not None and old.span is not None:
            self._spans.pop(old.span.name, None)
        self._configured.pop(folded, None)
        self._declared[folded] = region
        self._configured[folded] = region
        if region.span is not None:
            self._spans[region.span.name] = region.span
        return region

    @property
    def configured(self) -> list[Region]:
        """The tables declared from configuration, in the order declared."""
        return list(self._configured.values())

    @contextmanager
    def reading(self, span: "Span") -> Iterator[None]:
        """Read through ``span`` while a table is checked, before it is kept."""
        old = self._spans.get(span.name)
        self._spans[span.name] = span
        try:
            yield
        finally:
            if old is None:
                self._spans.pop(span.name, None)
            else:
                self._spans[span.name] = old

    def span(self, sheet: SheetRef) -> "Span | None":
        """The :class:`Span` whose own sheet ``sheet`` names, if any."""
        return self._spans.get(sheet.name) if sheet.name is not None else None

    def table_areas(self, sheet: str) -> list[tuple[Region, Area]]:
        """The cells of ``sheet`` the declared tables cover, each with its table."""
        folded = sheet.casefold()
        out = []
        for region in self._configured.values():
            if region.span is not None:
                out += [(region, a) for s, a in region.span.areas() if s.casefold() == folded]
            elif region.sheet.casefold() == folded:
                out.append((region, region.area))
        return out

    def withheld(self, sheet: str) -> list[Area]:
        """The cells of ``sheet`` in columns hidden from the model: never shown, never
        searched, read blank (:meth:`Workbook.read <xlsx_agent.workbook.Workbook.read>`)."""
        return [area for _, area in self.withholding(sheet)]

    def withholding(self, sheet: str) -> list[tuple[Region, Area]]:
        """:meth:`withheld`, each area with the declared table that hides it."""
        folded = sheet.casefold()
        return [(region, a) for name, region in self._configured.items()
                if region.span is not None and name not in self._redeclaring
                for s, a in region.span.withheld() if s.casefold() == folded]

    @contextmanager
    def redeclaring(self, names: "Sequence[str]") -> Iterator[None]:
        """Read the hidden cells of the tables named ``names`` while they are declared
        again: a declaration finds a hidden column by its header's text."""
        old = self._redeclaring
        self._redeclaring = old | {name.casefold() for name in names}
        try:
            yield
        finally:
            self._redeclaring = old

    def problems(self) -> list[Problem]:
        """What is odd about the tables and names, by code: each table's flags, a table
        the workbook cannot read (``bad_table``, at its sheet), a name whose
        ``localSheetId`` names no sheet (``bad_scope``), and two tables or workbook-wide
        names with one name, case aside (``duplicate_name``, at the second's address)."""
        out: list[Problem] = []
        for part in self.workbook._table_parts():
            out.extend(f for f in part.flags if not self._in_withheld(f.address))
            if part.area is not None and _broken(part):
                out.append(Problem("bad_table", quote_sheet(part.sheet) if part.sheet else None))
        for name in self.names:
            if name.bad_scope is not None:
                out.append(Problem("bad_scope", None))
        seen: set[tuple[str | None, str]] = set()
        for name, scope in [(t.name, None) for t in self.excel_tables] + \
                [(n.name, n.scope) for n in self.names]:
            key = (scope.casefold() if scope else None, name.casefold())
            if key in seen and is_address_name(name):
                written = TableAddress(name=name, scope=SheetRef(name=scope) if scope else None)
                out.append(Problem("duplicate_name", written.text))
            seen.add(key)
        return list(dict.fromkeys(out))

    def _in_withheld(self, address: str | None) -> bool:
        """Whether ``address`` (a cell or range, as a problem gives it) holds a withheld
        cell."""
        if address is None or "!" not in address:
            return False
        try:
            found = parse_address(address)
        except ValueError:
            return False
        if not isinstance(found, RangeAddress) or found.sheet.name is None:
            return False
        return any(_overlaps(found.area, a) for a in self.withheld(found.sheet.name))

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
        aside), with the key column declared for it (:meth:`declare_key`), if any."""
        return self.with_key(self._region(address))

    @property
    def declared_keys(self) -> dict[str, str]:
        """The key columns :meth:`declare_key` declared, by region address (folded)."""
        return dict(self._declared_keys)

    def with_key(self, region: Region) -> Region:
        """``region`` with the key column declared for it, when it has none of its own."""
        column = self._declared_keys.get(region.address.text.casefold())
        if column is None or region.key is not None:
            return region
        return replace(region, key=region.columns[region.column_index(column)])

    def declare_key(self, address: "str | TableAddress", column: str) -> Region:
        """Name ``column`` the key column of the table, name or region at ``address``
        (``t:Products``; a detected region's ``t:Stock!A1``), from then on: what an address
        by key looks keys up in, and what a derived value's sources are written by
        (:mod:`xlsx_agent.derive`), when the region declares none itself.  Declaring
        again replaces it.  The column must exist (:class:`NameNotFound` with the nearest);
        its cells need not all be keys -- a row without one, or sharing one, is refused
        where a key would name it."""
        region = self._region(address)
        name = region.columns[region.column_index(column)]
        self._declared_keys[region.address.text.casefold()] = name
        return replace(region, key=region.key or name)

    def _region(self, address: "str | TableAddress") -> Region:
        address = _table_address(address)
        if address.region is not None:
            return self.detect(address.region)
        if address.scope is not None:
            return self._scoped(address)
        folded = address.name.casefold()
        found: list[tuple[str, object]] = []
        found += [("table", t) for t in self.excel_tables if t.name.casefold() == folded]
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
        if kind == "table":
            declared = self.replaced_by(thing)
            if declared is not None:
                text = declared.address.text
                raise RegionError("declared", f"t:{thing.name} lies in {text}, a declared "
                                  f"table: address it as {text}", address=address.whole.text,
                                  valid_options=[text])
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
        hiding = [r.address.text for r, a in self.withholding(sheet.name) if _overlaps(area, a)]
        if hiding:
            written = name.address.text if name.address is not None else "the name"
            raise RegionError("declared", f"{written} holds cells of a column hidden from the "
                              "model by a declared table", address=written,
                              valid_options=list(dict.fromkeys(hiding)))
        columns = tuple(f"Column {column_letters(c)}"
                        for c in range(area.first.column, area.last.column + 1))
        return Region(kind="name", name=name.name, sheet=sheet.name, area=area, columns=columns,
                      header_rows=0, scope=name.scope, hidden=name.hidden, sheet_id=sheet.id,
                      workbook=self.workbook)

    # -- declared regions ------------------------------------------------------------------

    def declare(self, name: str, where: "str | RangeAddress", *, header_rows: int = 1,
                headers: Sequence[str] | None = None, key: str | None = None,
                keep: bool = True) -> Region:
        """A region declared over ``where`` (see :meth:`Workbook.declare_region
        <xlsx_agent.workbook.Workbook.declare_region>`), usable as ``t:<name>`` from then on;
        with ``keep`` false it is only made, not kept (a schema's region spec,
        :mod:`xlsx_agent.schema`)."""
        if not is_name(name):
            raise RegionError("bad_name", "a region's name is a name as ECMA-376 defines one: "
                              "a letter, _ or \\ first, then letters, digits, _ and .; not "
                              "TRUE, FALSE or a cell reference; at most 255 characters",
                              valid_options=["Returns", "Stock_2026"])
        folded = name.casefold()
        taken = [t.address.text for t in self.excel_tables if t.name.casefold() == folded]
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
        if any(_overlaps(area, hidden) for hidden in self.withheld(sheet.name)):
            raise RegionError("declared", f"{where.text} holds cells of a column hidden "
                              "from the model by a declared table", address=where.text)
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
        if keep:
            gone = self._configured.pop(folded, None)
            if gone is not None and gone.span is not None:
                self._spans.pop(gone.span.name, None)
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
        (``multi_row_header``: a second header row).  A block of two columns whose header
        is not text is a key-value region instead, when it is one (see the module's
        notes).  A cell in a table declared from configuration is refused
        (``declared``, with the table as ``valid_options``)."""
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
        for region, area in self.table_areas(sheet.name):
            if area.contains(where.cell):
                text = region.address.text
                raise RegionError("declared", f"{where.text} lies in {text}, a declared "
                                  f"table: address it as {text}", address=where.text,
                                  valid_options=[text])
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
            pairs = _key_value(header[:width], below, rows, width)
            if pairs is None:
                raise unclear("header_not_text")
            names = KEY_VALUE_COLUMNS + tuple(f"Column {column_letters(c)}" for c in
                                              range(cell.column + 2, cell.column + width))
            return self._titled(sheet, Region(
                kind="detected", name=None, sheet=sheet.name,
                area=Area(cell, Cell(cell.row + pairs - 1, right)), columns=names[:width],
                header_rows=0, key=KEY_VALUE_COLUMNS[0], sheet_id=sheet.id,
                shape="key_value", workbook=self.workbook))
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
        return self._titled(sheet, Region(kind="detected", name=None, sheet=sheet.name,
                                          area=Area(cell, Cell(bottom, right)),
                                          columns=tuple(columns), header_rows=1,
                                          sheet_id=sheet.id, flags=tuple(flags),
                                          workbook=self.workbook))

    def _titled(self, sheet: "Sheet", region: Region) -> Region:
        """``region`` with its ``title``: the one text cell over its columns two rows above
        its first row, the row between blank across them, when that cell is a caption
        (:attr:`Region.caption`); else ``region`` as it is."""
        first, last = region.area.first, region.area.last
        if region.caption or first.row < 3:
            return region
        two, one = self.workbook.read(RangeAddress(SheetRef(name=sheet.name), Area(
            Cell(first.row - 2, first.column), Cell(first.row - 1, last.column))))
        held = [v for v in two if v.t != "blank"]
        if any(v.t != "blank" for v in one) or len(held) != 1 or held[0].t != "text":
            return region
        try:
            above = self.detect(RangeAddress(SheetRef(name=sheet.name),
                                             Area(Cell(first.row - 2, held[0].column),
                                                  Cell(first.row - 2, held[0].column))))
        except RegionError:
            return region
        return replace(region, title=held[0].v) if above.caption else region

    def find(self, sheet: "Sheet | str | SheetRef") -> list["Region | UnclearRegion"]:
        """The blocks of a sheet outside its tables, each detected (a :class:`Region`) or
        not (an :class:`UnclearRegion`, with the reason): a block's header is the first
        row, from the top, holding something outside the tables and the blocks found
        above it, starting at its first such cell; further blocks in that row start after
        a blank column.  A block whose header is unclear is taken to run to the next row
        blank across the used range.  A title or caption (:attr:`Region.caption`), or a
        lone cell holding a number or date, is no block and is left out."""
        sheet = self.workbook._cell_sheet(sheet if not isinstance(sheet, str)
                                          else self.workbook.sheet(sheet))
        used = self.workbook.used_range(sheet)
        if used is None:
            return []
        taken = [t.area for t in self.excel_tables
                 if t.sheet.casefold() == sheet.name.casefold()]
        taken += [area for _, area in self.table_areas(sheet.name)]
        out: list[Region | UnclearRegion] = []
        written = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
        row_number = used.first.row
        rows = self._open_rows(sheet, used, row_number, taken)
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
                    if error.reason == "header_not_text" and self._lone(sheet, value):
                        column = value.column   # a lone number or date: no region either
                        continue
                    out.append(UnclearRegion(f"t:{written}!{_ref(value)}", error.reason))
                    column = _block_end(row, value.column)
                    bottom = max(bottom, self._blank_after(sheet, used, row_number))
                    continue
                if region.caption:   # a title or caption: a cell, never a region
                    column = value.column
                    continue
                out.append(region)
                taken.append(region.area)
                column = region.area.last.column
                bottom = max(bottom, region.area.last.row)
            if bottom > row_number:
                rows = self._open_rows(sheet, used, bottom + 1, taken)
        return out

    def _lone(self, sheet: "Sheet", value: Value) -> bool:
        """Whether the cell of ``value`` stands alone: the cells right of it and under it
        blank (a cell a header is not detected at already has nothing above it)."""
        cell = Cell(value.row, value.column)
        near = list(self.workbook.read(RangeAddress(SheetRef(name=sheet.name), Area(
            cell, Cell(cell.row + 1, cell.column + 1)))))
        return near[0][1].t == "blank" and all(v.t == "blank" for v in near[1][:2])

    def _open_rows(self, sheet: "Sheet", used: Area, top: int,
                   taken: list[Area]) -> Iterator[list[Value]]:
        """The rows from ``top`` that may start a block: those of the used range, less the
        rows the areas in ``taken`` cover from its first column to its last (such a row
        holds nothing outside them), skipped by starting a read after them -- when the
        sheet's index makes a read start near its rows; else every row, in one pass.
        ``taken`` is read as it is when each run of rows starts: an area added while a run
        is read covers rows up to the one being read, or :meth:`find` starts again below
        it."""
        if self.workbook.scan(sheet).index is None:
            yield from self._rows(sheet, used, top)
            return
        row, last = top, used.last.row
        while row <= last:
            edges = sorted({e for a in taken for e in (a.first.row, a.last.row + 1)})
            while row <= last and _covered(row, used, taken):
                row = next((e for e in edges if e > row), last + 1)
            if row > last:
                return
            end = row
            while end <= last and not _covered(end, used, taken):
                end = next((e for e in edges if e > end), last + 1)
            yield from self.workbook.read(RangeAddress(SheetRef(name=sheet.name), Area(
                Cell(row, used.first.column), Cell(min(end, last + 1) - 1, used.last.column))))
            row = end

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
        sheet = region.sheet_ref
        whole = region.address
        any_ = {"any_row": address.any_row, "any_column": address.any_column}
        if address.key is None and address.position is None and address.keys is None:
            if address.column is None:
                data = region.data
                text = replace(whole, **any_).text
                return Selection(text, whole.kind, region,
                                 RangeAddress(sheet, data) if data is not None else None)
            name = region.columns[region.column_index(address.column)]
            area = region.column_area(name)
            text = replace(whole, column=name, **any_).text
            return Selection(text, "column", region,
                             RangeAddress(sheet, area) if area is not None else None,
                             column=name)
        column = None
        if address.column is not None:
            column = region.columns[region.column_index(address.column)]
        if address.keys is not None:
            positions = tuple(self._find_key(region, k, key, k_th=n)
                              for n, k in enumerate(address.keys, 1))
            text = replace(whole, keys=address.keys, column=column, **any_).text
            return Selection(text, "values" if column is not None else "rows", region, None,
                             column=column, positions=positions)
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
        text = replace(whole, key=address.key, position=address.position, column=column,
                       **any_).text
        return Selection(text, "value" if column is not None else "row", region,
                         RangeAddress(sheet, area), row=position, column=column)

    def expand(self, address: "str | TableAddress", *, key: str | None = None
               ) -> list[TableAddress]:
        """What a ``t:`` address covers, as one address by key for each cell, in the
        table's row order (a key set's, in its order), then column by column:
        ``t:Products/*/Price`` gives ``t:Products/SKU-0001/Price``, ... for every data row;
        ``t:Products/*/*`` every value; ``t:Products/SKU-0042/*`` each value of one row; a
        key set each of its values.  Without a column it gives rows (``t:Products`` and
        ``t:Products/*``: ``t:Products/SKU-0001``, ...).  A column (``t:Products[Price]``),
        a row by position and a value come back by key too.

        The keys are the key column's (``key``, else the region's own, as for a row by
        key): :class:`RegionError` ``no_key_column`` without one.  A row it covers whose
        key cell is no key (blank, a date, a fraction...) is refused as ``no_key``, and a
        key two rows share as ``duplicate_key``, each with the rows by position as
        ``valid_options``: an address by key is never made for a row it would not name
        alone."""
        address = _table_address(address)
        region = self.region(address)
        index = self.key_index(region, key)
        whole = region.address
        if address.keys is not None:
            positions = [index.find(k, n) for n, k in enumerate(address.keys, 1)]
        elif address.key is not None:
            positions = [index.find(address.key)]
        elif address.position is not None:
            positions = [self.resolve(replace(address, column=None, any_column=False)).row]
        else:
            positions = list(range(1, region.rows + 1))
        blank = [p for p in positions if index.keys[p - 1] is None]
        if blank:
            options = [replace(whole, position=p).text for p in blank[:50]]
            raise RegionError("no_key", f"{len(blank)} rows of {whole.text} have no key in "
                              f"column {index.column!r} (blank, or not text or a whole "
                              "number); name them by position", address=whole.text,
                              valid_options=options)
        shared = [p for p in positions if index.keys[p - 1] in index.duplicates]
        if shared:
            options = [replace(whole, position=p).text for p in shared[:50]]
            raise RegionError("duplicate_key", f"{len(shared)} rows of {whole.text} share "
                              f"a key in column {index.column!r}; name them by position",
                              address=whole.text, valid_options=options)
        if address.column is not None:
            columns: list[str | None] = [region.columns[region.column_index(address.column)]]
        elif address.any_column:
            columns = list(region.columns)
        else:
            columns = [None]
        return [replace(whole, key=index.keys[p - 1], column=c)
                for p in positions for c in columns]

    def key_index(self, region: Region, key: str | None = None) -> "KeyIndex":
        """The :class:`KeyIndex` of ``region``'s key column (:meth:`key_column`), read once
        and kept."""
        column = self.key_column(region, key)
        cache = (_key_cache(region), column)
        if cache not in self._indexes:
            keys: list[str | None] = []
            area = region.column_area(column)
            if area is not None:
                for row in self.workbook.read(RangeAddress(region.sheet_ref, area)):
                    keys.append(key_text(row[0]))
            self._indexes[cache] = KeyIndex(region, column, keys)
        return self._indexes[cache]

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
        cache = _key_cache(region)
        if cache in self._keys:
            return list(self._keys[cache])
        finder = KeyFinder(region)
        data = region.data
        if data is not None:
            for row in self.workbook.read(RangeAddress(region.sheet_ref, data)):
                finder.add(row)
        return self.keep_keys(region, finder)

    def key_finder(self, region: Region) -> "KeyFinder | None":
        """A :class:`KeyFinder` for ``region`` to be given its data rows by a pass that
        reads them anyway (the catalogue's profile), then :meth:`keep_keys`; ``None`` when
        :meth:`key_columns` knows them already."""
        return None if _key_cache(region) in self._keys else KeyFinder(region)

    def keep_keys(self, region: Region, finder: "KeyFinder") -> list[str]:
        """What ``finder`` found, given every data row of ``region``: kept as
        :meth:`key_columns`' answer."""
        out = finder.keys()
        self._keys[_key_cache(region)] = out
        return list(out)

    def _find_key(self, region: Region, wanted: str, key: str | None,
                  k_th: int | None = None) -> int:
        return self.key_index(region, key).find(wanted, k_th)


class KeyIndex:
    """A region's key column, read once: ``keys`` each data row's key by position (from
    0: row 1 is ``keys[0]``; ``None`` for a cell that is no key, :func:`key_text`),
    ``duplicates`` the keys two or more rows share, with their positions.  :meth:`find`
    looks a key up as an address by key does."""

    def __init__(self, region: Region, column: str, keys: list[str | None]) -> None:
        self.region = region
        self.column = column
        self.keys = keys
        self._first: dict[str, int] = {}
        self.duplicates: dict[str, list[int]] = {}
        for position, k in enumerate(keys, 1):
            if k is None:
                continue
            first = self._first.setdefault(k, position)
            if first != position:
                self.duplicates.setdefault(k, [first]).append(position)

    @property
    def unique(self) -> bool:
        """Every data row has a key, and no two share one: each row has an address by
        key."""
        return not self.duplicates and None not in self.keys

    def find(self, wanted: str, k_th: int | None = None) -> int:
        """The position of the row whose key is ``wanted``: :class:`NameNotFound` when no
        row has it, :class:`RegionError` ``duplicate_key`` (with the rows by position) when
        two do.  ``k_th``: the key's place in a key set, for the message (never the key
        itself)."""
        whole = self.region.address
        which = "that key" if k_th is None else f"key {k_th} of the set"
        if wanted in self.duplicates:
            found = self.duplicates[wanted]
            options = [replace(whole, position=p).text for p in found[:50]]
            raise RegionError("duplicate_key", f"{len(found)} rows of {whole.text} have "
                              f"{which} in column {self.column!r}; name the row by its "
                              "position", address=whole.text, valid_options=options)
        if wanted not in self._first:
            raise NameNotFound(f"no row of {whole.text} has {which} in column "
                               f"{self.column!r}", what="key", name=wanted)
        return self._first[wanted]


class KeyFinder:
    """The columns of a region that can be its key (:meth:`Regions.key_columns`), from its
    data rows, each given whole and in order to :meth:`add`: a column is dropped at its
    first cell that is not a key, or a key met before."""

    def __init__(self, region: Region) -> None:
        self.columns = region.columns
        data = region.data
        width = 0 if data is None else data.last.column - data.first.column + 1
        self._seen: list[set[str] | None] = [set() for _ in range(width)]

    def add(self, row: Sequence[Value]) -> None:
        seen = self._seen
        for index, value in enumerate(row):
            keys = seen[index]
            if keys is None:
                continue
            text = key_text(value)
            if text is None or text in keys:
                seen[index] = None
            else:
                keys.add(text)

    def keys(self) -> list[str]:
        return [self.columns[i] for i, keys in enumerate(self._seen) if keys is not None]


def _key_cache(region: Region) -> tuple:
    return (region.kind, region.sheet, region.area, region.columns, region.span)


def key_text(value: Value) -> str | None:
    """A cell as a key: its text, exactly, or a whole number's digits (``42`` for 42 and
    42.0); ``None`` for anything else (a blank, empty text, a fraction, a date, a time, a
    boolean, an error).  See the module's notes."""
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


def _key_value(first: list[Value], below: "list[Value] | None", rows: Iterator[list[Value]],
               width: int) -> int | None:
    """How many rows a key-value block has, from its first: :data:`KEY_VALUE_WIDTH`
    columns, :data:`KEY_VALUE_ROWS` rows at most, down to the first row blank across them,
    every row's first cell text and a key no other row's is (and at least two rows);
    ``None`` when the block is not one."""
    if width != KEY_VALUE_WIDTH or below is None:
        return None
    seen: set[str] = set()
    count = 0
    for row in _chain([first, below], rows):
        values = row[:width]
        if all(v.t == "blank" for v in values):
            break
        key = values[0].v if values[0].t == "text" else None
        if not key or key in seen or count >= KEY_VALUE_ROWS:
            return None
        seen.add(key)
        count += 1
    return count if count >= 2 else None


def _chain(first: list, rest: Iterator) -> Iterator:
    yield from first
    yield from rest


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


def _covered(row: int, used: Area, taken: list[Area]) -> bool:
    """Whether the areas in ``taken`` cover ``row`` from the used range's first column to
    its last."""
    reach = used.first.column
    for first, last in sorted((a.first.column, a.last.column) for a in taken
                              if a.first.row <= row <= a.last.row):
        if first > reach:
            return False
        reach = max(reach, last + 1)
    return reach > used.last.column


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
