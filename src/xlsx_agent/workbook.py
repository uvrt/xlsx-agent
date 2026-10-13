"""The open workbook: its package, its kind, its sheets.

``Workbook.open`` takes bytes, a path, a binary file or an ooxml-edit
:class:`~ooxml_edit.opc.OpcPackage` (a chart's embedded workbook, say), checks it as
:mod:`xlsx_agent.package` says, and reads the workbook part and each sheet's protection
(streamed: a sheet is never parsed whole)::

    wb = Workbook.open("inventory.xlsx")
    wb.sheets        # [Sheet(name='Stock', id=3, state='visible', protected=False, type='worksheet')]
    wb.to_bytes()    # the bytes that were opened, while nothing has changed

and its cells, typed, streamed from the sheet (:mod:`xlsx_agent.read`)::

    wb = Workbook.open("inventory.xlsx", locale=Locale(decimal_separator=",",
                       thousands_separator=".", date_order="dmy", language="nl-BE",
                       list_separator=";"))      # ooxml_common.locales.Locale
    wb.cell("Stock!F2")                  # Value(sheet='Stock', row=2, column=6, v=12.5, t='number')
    for row in wb.read("Stock!A:C"):     # whole columns, clipped to the used range; lazy
        ...
    wb.used_range("Stock")               # Area A1:H40, computed from the cells
    wb.problems                          # [Problem(code='no_cached_value', address='Stock!F7', count=1)]

and its tables, defined names and regions (:mod:`xlsx_agent.read.regions`), which a ``t:``
address resolves against::

    wb.tables                            # [Region(kind='table', name='Products', ...)]
    wb.names                             # [DefinedName(name='Prices', scope=None, ...)]
    wb.resolve("t:Products/SKU-0042/Price", key="SKU").range    # Stock!C7
    wb.read("t:Products[Price]")         # the column's data cells, as any range is read
    wb.declare_region("Returns", "Returns!A3:F200", header_rows=2, headers=[...])

and what each column holds (:mod:`xlsx_agent.profile`), and the whole workbook at a
glance (:mod:`xlsx_agent.catalogue`)::

    wb.table("t:Products").column("Price").stats()     # ColumnStats(type='number', ...)
    wb.profile("t:Products")                           # every column's, in one pass
    wb.catalogue()                                     # what describe returns

and where some words are (:mod:`xlsx_agent.find`)::

    wb.find_text("Hardware", within="Stock!")          # Found(kind='cell', address='Stock!B7', ...)

and what changed between two versions of it (:mod:`xlsx_agent.compare`)::

    Workbook.compare(old, new, tables=["t:Products"], key="SKU")   # Diff

and whether it holds what a declared schema says (:mod:`xlsx_agent.schema`)::

    wb.validate_data({"tables": [{"table": "t:Products", "key": ["SKU"], "columns": [
        {"name": "Price", "type": "number", "min": 0}]}]})     # ValidationReport

Transitional and strict packages (``purl.oclc.org`` namespaces) are read alike.
"""

from __future__ import annotations

import json
import os
import re
from collections import OrderedDict
from dataclasses import dataclass, field, replace
from typing import Any, BinaryIO, Iterator, Mapping, Sequence

from lxml import etree
from ooxml_edit.charts.workbook import Area
from ooxml_edit.opc import OpcPackage
from ooxml_edit.tools import Dataset
from ooxml_edit.tools.limits import Limits
from ooxml_edit.xml import Element

from .address import (NameNotFound, RangeAddress, SheetRef, TableAddress, nearest,
                      parse_address, quote_sheet)
from .catalogue import DEFAULT_TOKENS, catalogue
from .compare import Diff, compare
from .derive import Derived, derive
from .find import Found, find_text
from .package import OpenError, open_package, read_source, workbook_part
from .profile import Column, ColumnStats, profile
from .read.names import DefinedName, read_names
from .schema import ValidationReport, validate
from .read.regions import Region, Regions, Selection, UnclearRegion
from .read.comments import Comment, read_comments
from .read.stream import Hyperlink, SheetReader, SheetScan, read_area, scan
from .read.strings import SharedStrings
from .read.styles import NumberFormat, Styles
from .read.tables import TablePart, read_table
from .read.types import Problem, Value, format_language, reading_locale
from .xml import iterparse_xml

#: SpreadsheetML's main namespace, transitional and strict.
NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")
#: The relationships namespace of ``r:id``, transitional and strict.
NS_R = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships")

#: ``ST_SheetState`` (ECMA-376 Part 1, 18.18.68).
SHEET_STATES = ("visible", "hidden", "veryHidden")

#: The last segment of a sheet's relationship type -> :attr:`Sheet.type`.
SHEET_TYPES = {"worksheet": "worksheet", "chartsheet": "chartsheet", "dialogsheet": "dialogsheet",
               "xlMacrosheet": "macrosheet", "xlIntlMacrosheet": "macrosheet"}

_PROTECTION = tuple(f"{{{ns}}}sheetProtection" for ns in NS_MAIN)
_ROWS = tuple(f"{{{ns}}}row" for ns in NS_MAIN)


@dataclass(frozen=True)
class Sheet:
    """One sheet, as the workbook part lists it.

    ``id`` is its ``sheetId`` (what ``#3!F2`` names; it survives a rename), ``None`` when
    the workbook gives none that is a number.  ``state`` is ``"visible"``, ``"hidden"`` or
    ``"veryHidden"`` (as written, when the file says something else).  ``protected``: the
    sheet's ``sheetProtection`` turns protection on (``sheet="1"``; a chart sheet's
    ``content="1"``).  ``type`` is ``"worksheet"``, ``"chartsheet"``, ``"dialogsheet"`` or
    ``"macrosheet"``, from its relationship; ``part`` the package part, ``None`` when the
    relationship leads nowhere."""

    name: str
    id: int | None
    state: str = "visible"
    protected: bool = False
    type: str | None = "worksheet"
    part: str | None = field(default=None, repr=False)

    @property
    def visible(self) -> bool:
        return self.state == "visible"


class Workbook:
    """An open workbook.  Made by :meth:`open`."""

    def __init__(self, package: OpcPackage, *, part: str, kind: str, data: bytes | None = None,
                 locale: Any = None, grouping: bool = True, local_dates: bool = False) -> None:
        #: The ooxml-edit package: every part as the bytes that were read.
        self.package = package
        #: The workbook part (``xl/workbook.xml`` in every file Excel writes).
        self.part = part
        #: ``"xlsx"``, ``"xlsm"``, ``"xltx"`` or ``"xltm"``: what the workbook part declares.
        self.kind = kind
        self._locale = reading_locale(locale)
        #: Whether a text number's whole part may be grouped with the locale's thousands
        #: separator (``1,234.5``), and whether text dates other than ISO 8601 are read in
        #: the locale's date order: reader settings, set at :meth:`open`
        #: (:mod:`xlsx_agent.read.types`).
        self.grouping = bool(grouping)
        self.local_dates = bool(local_dates)
        self._data = data
        self._sheets: list[Sheet] | None = None
        self._strings: SharedStrings | None = None
        self._styles: Styles | None = None
        self._scans: dict[str, SheetScan] = {}
        self._comments: dict[str, list[Comment]] = {}
        self._table_parts_read: list[TablePart] | None = None
        self._catalogue: tuple | None = None   # what the catalogue is made from, gathered
        #: ``xlsx_query``'s last few matches, by query: its next page needs them again.
        self._queries: OrderedDict = OrderedDict()
        #: ``find_text``'s last few counts (total and kinds), by search: its next page
        #: stops once it is full.
        self._finds: OrderedDict = OrderedDict()
        self.regions = Regions(self)

    # -- lifecycle -------------------------------------------------------------------------

    @classmethod
    def open(cls, source: "str | os.PathLike[str] | bytes | BinaryIO | OpcPackage", *,
             limits: Limits | None = None, locale: Any = None, grouping: bool = True,
             local_dates: bool = False) -> "Workbook":
        """Open a ``.xlsx``, ``.xlsm``, ``.xltx`` or ``.xltm`` from a path, bytes, a binary
        file or an :class:`~ooxml_edit.opc.OpcPackage`.

        Bytes are checked first against ``limits`` (ooxml-edit's
        :class:`~ooxml_edit.tools.limits.Limits`; its defaults when ``None``): a package
        over a size, part-count or expansion limit is refused before it is unpacked.  A
        ``.xls``, a ``.xlsb``, an encrypted workbook or any other file is refused too:
        :class:`~xlsx_agent.package.OpenError`, whose ``code`` says which.  A path whose
        extension names another kind than the package's opens with a
        :class:`~xlsx_agent.package.KindMismatchWarning`.

        The workbook part and the sheet list are read here, so a part the safe parser
        refuses (an entity bomb, XML that is not well-formed) refuses the workbook now,
        with the part's name (``code`` ``"damaged"``), rather than at a later read.

        ``locale`` is how text is read as numbers and dates, and the language the
        language-dependent built-in number formats are read in: an
        :class:`ooxml_common.locales.Locale`, or any object with its five string fields;
        en-US (:data:`~ooxml_common.locales.EN_US`) when ``None``.  ``grouping`` (on)
        recognises a text number whose whole part is grouped with the locale's thousands
        separator; ``local_dates`` (off: ISO 8601 dates only) recognises text dates in the
        locale's date order too -- Excel's own reading, measured
        (:mod:`xlsx_agent.read.types`).  All three are the caller's settings, never a
        model's."""
        if isinstance(source, OpcPackage):
            part, kind = workbook_part(source)
            workbook = cls(source, part=part, kind=kind, locale=locale, grouping=grouping,
                           local_dates=local_dates)
        else:
            limits = limits or Limits()
            data = read_source(source, limits)
            package = open_package(data, limits)
            part, kind = workbook_part(package, source)
            workbook = cls(package, part=part, kind=kind, data=data, locale=locale,
                           grouping=grouping, local_dates=local_dates)
        workbook._sheets = workbook._read_sheets()
        return workbook

    @property
    def locale(self) -> Any:
        """How text is read as numbers and dates, and the language of the
        language-dependent built-in number formats: the caller's
        (:class:`ooxml_common.locales.Locale`; en-US unless given).  Setting another one
        (as the tool layer does with its session's) forgets what was read in the old one:
        the number formats, the sheets' scans, the detected regions, the catalogue."""
        return self._locale

    @locale.setter
    def locale(self, locale: Any) -> None:
        locale = reading_locale(locale)
        if locale == self._locale:
            return
        self._locale = locale
        self._styles = None
        self._scans.clear()
        self._catalogue = None
        self._queries.clear()
        self._finds.clear()
        self.regions.forget_cells()

    def to_bytes(self) -> bytes:
        """The workbook as a file.  While nothing has changed it is the bytes that were
        opened, byte for byte; a workbook opened from a package is that package's
        :meth:`~ooxml_edit.opc.OpcPackage.to_bytes`."""
        if self._data is not None and not self.package.changed_parts():
            return self._data
        return self.package.to_bytes()

    # -- the workbook part -----------------------------------------------------------------

    @property
    def _root(self) -> Element:
        return _parsed(self.part, lambda: self.package.tree(self.part))

    def _child(self, local: str) -> Element | None:
        for ns in NS_MAIN:
            node = self._root.find(f"{{{ns}}}{local}")
            if node is not None:
                return node
        return None

    @property
    def sheets(self) -> list[Sheet]:
        """Every sheet, in the workbook's order (the order of its tabs)."""
        if self._sheets is None:
            self._sheets = self._read_sheets()
        return list(self._sheets)

    def sheet(self, name: str) -> Sheet:
        """The sheet called ``name``, matched as Excel matches names (case aside);
        :class:`~xlsx_agent.address.NameNotFound` (a ``KeyError``) naming the sheets there
        are, with the nearest as ``valid_options``, as an address writes them."""
        folded = name.casefold()
        for sheet in self.sheets:
            if sheet.name.casefold() == folded:
                return sheet
        names = [sheet.name for sheet in self.sheets]
        written = [quote_sheet(sheet.name) for sheet in self.sheets if sheet.name]
        raise NameNotFound(f"no sheet {name!r}; the sheets are {names}", what="sheet",
                           name=name, candidates=written,
                           valid_options=nearest(quote_sheet(name) if name else "", written))

    def sheet_by_id(self, sheet_id: int) -> Sheet:
        """The sheet whose ``sheetId`` is ``sheet_id`` (what ``#3!F2`` names);
        :class:`~xlsx_agent.address.NameNotFound` with the ids there are, as ``#3``."""
        for sheet in self.sheets:
            if sheet.id is not None and sheet.id == sheet_id:
                return sheet
        ids = [f"#{sheet.id}" for sheet in self.sheets if sheet.id is not None]
        raise NameNotFound(f"no sheet with id {sheet_id}; the ids are {ids}", what="sheet id",
                           name=sheet_id, candidates=ids, valid_options=ids[:50])

    def sheet_of(self, address: "str | SheetRef | RangeAddress | TableAddress") -> Sheet:
        """The sheet an address names: a :class:`~xlsx_agent.address.SheetRef`, a cell or
        range (``Stock!F2``, ``#3!A:A``), a detected region (``t:Stock!A1``), or such an
        address as text.  By name, matched case aside; by id, its ``sheetId``.  A table,
        dataset or ref names no sheet by itself: ``ValueError``."""
        if isinstance(address, str):
            address = parse_address(address)
        if isinstance(address, TableAddress) and address.region is not None:
            address = address.region
        if isinstance(address, RangeAddress):
            address = address.sheet
        if not isinstance(address, SheetRef):
            raise ValueError(f"a {getattr(address, 'kind', type(address).__name__)} address "
                             "names no sheet by itself")
        if address.id is not None:
            return self.sheet_by_id(address.id)
        return self.sheet(address.name)

    @staticmethod
    def sheet_refs(sheet: Sheet) -> tuple[SheetRef | None, SheetRef | None]:
        """Both ways to name ``sheet``: by name and by id, each ``None`` when the workbook
        gives none (an empty name, a ``sheetId`` that is not a number).  A result gives
        both the first time it names a sheet."""
        return (SheetRef(name=sheet.name) if sheet.name else None,
                SheetRef(id=sheet.id) if sheet.id is not None and sheet.id >= 0 else None)

    @property
    def protected(self) -> bool:
        """The workbook's structure or windows are locked (``workbookProtection``)."""
        node = self._child("workbookProtection")
        if node is None:
            return False
        return any(_true(node.get(attribute)) for attribute in ("lockStructure", "lockWindows"))

    @property
    def date1904(self) -> bool:
        """The workbook uses the 1904 date system (``workbookPr date1904``); else 1900."""
        node = self._child("workbookPr")
        return node is not None and _true(node.get("date1904"))

    @property
    def calculation(self) -> list[str]:
        """Why cached values may be stale: ``"full_calc_on_load"`` (``calcPr
        fullCalcOnLoad``: the file asks for a full recalculation when it opens) and
        ``"manual_calculation"`` (saved with ``calcMode="manual"``)."""
        node = self._child("calcPr")
        if node is None:
            return []
        out = []
        if _true(node.get("fullCalcOnLoad")):
            out.append("full_calc_on_load")
        if (node.get("calcMode") or "").strip() == "manual":
            out.append("manual_calculation")
        return out

    @property
    def problems(self) -> list[Problem]:
        """What is odd about the workbook, by code and address, never by content: the
        :attr:`calculation` flags once each, then, per worksheet, each issue its cells
        have (``no_cached_value``, ``number_as_text``, ``at_limit``, ``volatile``...) with
        its first address and count.  Every worksheet is scanned once (streamed) the first
        time this is asked."""
        out = [Problem(code) for code in self.calculation]
        out.extend(self.regions.problems())
        for sheet in self.sheets:
            if sheet.part is not None and sheet.type in _CELL_SHEETS:
                out.extend(self.scan(sheet).problems)
        return out

    # -- tables, names and regions ---------------------------------------------------------

    @property
    def tables(self) -> list[Region]:
        """The Excel tables (ListObjects), sheet by sheet in the workbook's order, each in
        the order its sheet relates to it; a table whose ``ref`` is not a range is left out
        and reported in :attr:`problems` (``bad_ref``, ``bad_table``)."""
        return self.regions.tables

    @property
    def names(self) -> list[DefinedName]:
        """The defined names, in the file's order, hidden and built-in ones too; those that
        are not a range say why (:mod:`xlsx_agent.read.names`)."""
        return self.regions.names

    def region(self, address: "str | TableAddress") -> Region:
        """The table, name or region a ``t:`` address names (its column or row aside):
        :class:`~xlsx_agent.address.NameNotFound` for an unknown name, with the nearest as
        ``valid_options``; :class:`~xlsx_agent.read.regions.RegionError` for a name that is
        not a range, an ambiguous one, or a header that is not detected."""
        return self.regions.region(address)

    def table(self, address: "str | TableAddress", *, key: str | None = None) -> Region:
        """As :meth:`region`, with ``key`` (a column name) as its key column."""
        region = self.regions.region(address)
        if key is not None:
            region = replace(region, key=region.columns[region.column_index(key)])
        return region

    def resolve(self, address: "str | TableAddress", *, key: str | None = None) -> Selection:
        """What a ``t:`` address names, as cells: ``t:Products`` (its data rows),
        ``t:Products[Price]``, ``t:Products/#5``, ``t:Products/SKU-0042``,
        ``t:Products/SKU-0042/Price``, ``t:Stock!Prices`` and ``t:Stock!A1`` (a detected
        region).  ``key`` names the key column for a row by key; see
        :mod:`xlsx_agent.read.regions` for how keys, columns and names are matched."""
        return self.regions.resolve(address, key=key)

    def declare_region(self, name: str, where: "str | RangeAddress", *, header_rows: int = 1,
                       headers: "list[str] | None" = None, key: str | None = None) -> Region:
        """Declare a region over a range (``Returns!A3:F200``), usable as ``t:<name>`` from
        then on: ``header_rows`` header rows (0 for none), named by ``headers`` (one per
        column, unique case aside; required with two or more header rows, since a
        multi-row or merged header is never read as names), else by the header row's text
        (one header row) or the columns' letters (none).  ``key`` names the key column.
        ``name`` follows ECMA-376's rule for names (:func:`~xlsx_agent.address.is_name`)
        and is no table's or workbook-wide name's; declaring a region's name again replaces
        it."""
        return self.regions.declare(name, where, header_rows=header_rows, headers=headers,
                                    key=key)

    def detect_region(self, where: "str | RangeAddress | TableAddress") -> Region:
        """The region whose single-row header starts at a cell (``t:Stock!A1``); see
        :meth:`~xlsx_agent.read.regions.Regions.detect`."""
        return self.regions.detect(where)

    def find_regions(self, sheet: "Sheet | str | SheetRef") -> "list[Region | UnclearRegion]":
        """The blocks of a sheet outside its tables, each detected or flagged
        ``header_unclear``; see :meth:`~xlsx_agent.read.regions.Regions.find`."""
        return self.regions.find(sheet)

    def column(self, address: "str | TableAddress", *, key: str | None = None) -> Column:
        """The column a ``t:`` column address names (``t:Products[Price]``), whose
        :meth:`~xlsx_agent.profile.Column.stats` is its profile; ``ValueError`` for an
        address that names no column, or a row or value of one."""
        if isinstance(address, str):
            address = TableAddress.parse(address)
        if not isinstance(address, TableAddress) or address.column is None \
                or address.key is not None or address.position is not None:
            raise ValueError("a column address names one column of a table, as t:Products[Price]")
        region = self.table(address.whole, key=key)
        return Column(self, region, address.column)

    def profile(self, address: "str | TableAddress | Region",
                columns: "list[str] | None" = None) -> list[ColumnStats]:
        """The profile of every column of a table, name or region (or of ``columns``, by
        name), in one streamed pass: what ``xlsx_profile`` returns
        (:mod:`xlsx_agent.profile`)."""
        region = address if isinstance(address, Region) else self.regions.region(address)
        return profile(self, region, columns)

    def key_columns(self, address: "str | TableAddress | Region") -> list[str]:
        """The columns that can be a region's key: unique, never blank, text or whole
        numbers."""
        region = address if isinstance(address, Region) else self.regions.region(address)
        return self.regions.key_columns(region)

    # -- derived figures -------------------------------------------------------------------

    def derive(self, from_: "str | TableAddress | Region | Derived | Dataset", *, where=None,
               group_by=None, measures=None, bins=None, pivot=None, sort=None,
               top: int | None = None, round: "int | dict[str, int] | None" = None, coerce=None,
               id: str | None = None) -> Derived:
        """Figures from a table, name or region (``t:Products``), computed here, each with
        the cells it came from: ``where`` (:mod:`xlsx_agent.query`'s filters), ``group_by``
        (column names), ``measures`` (``[{"fn": "mean", "column": "Price", "as":
        "average"}]``: count, sum, mean, median, min, max, share, distinct), ``bins``
        (``{"column", "edges", "as"}``), ``pivot`` (a column whose values become columns),
        ``sort``, ``top``, ``round`` (decimals, half away from zero, on the reported values
        only), ``coerce`` (``["number_as_text"]``) and ``id``.  A measure its column's type
        does not take is refused before anything is computed; see
        :mod:`xlsx_agent.derive`.  ``from_`` may be a dataset instead -- an earlier
        :class:`Derived`, or an ooxml-edit ``Dataset`` -- whose sources carry through to
        the original cells (:mod:`xlsx_agent.dataset`)."""
        return derive(self, from_, where=where, group_by=group_by, measures=measures,
                      bins=bins, pivot=pivot, sort=sort, top=top, round=round,
                      coerce=coerce, id=id)

    # -- comparing -------------------------------------------------------------------------

    @staticmethod
    def compare(a: "Workbook", b: "Workbook", *,
                tables: "str | Sequence[str | tuple[str, str]] | None" = None,
                key: "str | Mapping[str, str] | None" = None,
                columns: "Sequence[str] | None" = None) -> Diff:
        """``a`` (the older workbook) against ``b`` (the newer): sheets by id, tables and
        regions by name (``tables``: some of them, or pairs; every table and declared region
        when ``None``), rows by key (``key``: a column name, or one per table; else the
        declared key, or the one key column both find), and cells by type and value, each
        change with both cells typed.  A key two rows share is reported, never paired; see
        :mod:`xlsx_agent.compare`."""
        return compare(a, b, tables=tables, key=key, columns=columns)

    # -- checking against a schema ---------------------------------------------------------

    def validate_data(self, schema: Mapping[str, Any]) -> ValidationReport:
        """The workbook checked against a declared ``schema`` (plain data: required
        ``sheets``; ``tables``, each a ``t:`` address or a declared region spec, with its
        ``key`` and its ``columns``' ``type``, ``required``, ``unique``, ``min``, ``max``,
        ``step``, ``enum``, ``case``, ``validation_list`` and ``pattern``).  The report
        lists problems by code and address, never a cell's content, and counts them by
        code; a schema that is not one is refused (:class:`~xlsx_agent.schema.SchemaError`)
        before anything is read.  See :mod:`xlsx_agent.schema`."""
        return validate(self, schema)

    # -- finding text ----------------------------------------------------------------------

    def find_text(self, needle: "str | re.Pattern", *,
                  within: "str | RangeAddress | TableAddress | None" = None) -> Iterator[Found]:
        """Where ``needle`` (text, case counting, or a compiled pattern) is: text cells,
        table and region names, table headers, comments and defined names, each match by
        address and kind, sheet by sheet in reading order, streamed; ``within`` a sheet
        (``Stock!``), range or table.  Formula text and non-text values are not searched;
        see :mod:`xlsx_agent.find`."""
        return find_text(self, needle, within)

    # -- the catalogue ---------------------------------------------------------------------

    def catalogue(self, *, detail: str = "catalogue", max_tokens: int = DEFAULT_TOKENS) -> dict:
        """What the workbook holds, as data: sheets, regions with their key columns and
        each column's profile summarised, sample rows, flags and problems, within
        ``max_tokens`` as the siblings' offline proxy counts them (``detail="full"``:
        every column's whole profile, nothing cut).  What the shared ``describe`` tool
        returns; see :mod:`xlsx_agent.catalogue`."""
        return catalogue(self, detail=detail, max_tokens=max_tokens)

    def describe(self, *, max_tokens: int = DEFAULT_TOKENS) -> str:
        """The catalogue as text: :meth:`catalogue`'s data as compact JSON."""
        return json.dumps(self.catalogue(max_tokens=max_tokens), ensure_ascii=False,
                          separators=(",", ":"), default=str)

    def _table_parts(self) -> list[TablePart]:
        if self._table_parts_read is None:
            out = []
            for sheet in self.sheets:
                if sheet.part is None or sheet.type != "worksheet":
                    continue
                written = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
                for rel in self.package.relationships(sheet.part).values():
                    if rel.is_external or rel.type.rpartition("/")[2] != "table" \
                            or rel.target_part is None \
                            or not self.package.has_part(rel.target_part):
                        continue
                    part = rel.target_part
                    root = _parsed(part, lambda: self.package.tree(part))
                    table = read_table(root, sheet=sheet.name, written_sheet=written, part=part)
                    if table is not None:
                        out.append(table)
            self._table_parts_read = out
        return list(self._table_parts_read)

    def _defined_names(self) -> list[DefinedName]:
        return read_names(self._child("definedNames"), self.sheets)

    # -- cells -----------------------------------------------------------------------------

    def _related(self, kind: str) -> str | None:
        """The workbook's part related by a relationship whose type ends in ``kind``
        (``"sharedStrings"``, ``"styles"``): transitional and strict alike."""
        for rel in self.package.relationships(self.part).values():
            if not rel.is_external and rel.type.rpartition("/")[2] == kind \
                    and rel.target_part is not None and self.package.has_part(rel.target_part):
                return rel.target_part
        return None

    @property
    def strings(self) -> SharedStrings:
        """The shared-string table, read (streamed) the first time a cell needs it."""
        if self._strings is None:
            part = self._related("sharedStrings")
            self._strings = SharedStrings(self.package.read(part) if part else None)
        return self._strings

    @property
    def styles(self) -> Styles:
        """``styles.xml``'s number formats, read once."""
        if self._styles is None:
            part = self._related("styles")
            root = _parsed(part, lambda: self.package.tree(part)) if part else None
            self._styles = Styles(root, format_language(self.locale.language))
        return self._styles

    def number_format(self, style: int) -> NumberFormat:
        """The number format of a cell whose ``s`` is ``style`` (a ``cellXfs`` index)."""
        return self.styles.format(style)

    def _cell_sheet(self, sheet: "Sheet | str | SheetRef") -> Sheet:
        if isinstance(sheet, Sheet):
            found = sheet
        elif isinstance(sheet, SheetRef):
            found = self.sheet_of(sheet)
        else:
            found = self.sheet(sheet)
        if found.type not in _CELL_SHEETS:
            raise ValueError(f"sheet {quote_sheet(found.name) if found.name else found.id} "
                             f"is a {found.type or 'sheet of no known type'} and holds no cells")
        return found

    def _reader(self, sheet: Sheet, *, formats: bool = False) -> SheetReader:
        data = self.package.read(sheet.part) if sheet.part is not None else None
        name = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
        return SheetReader(data or b"", sheet=name, hidden=not sheet.visible,
                           strings=self.strings, styles=self.styles, date1904=self.date1904,
                           locale=self.locale, formats=formats, grouping=self.grouping,
                           local_dates=self.local_dates)

    def _stream(self, sheet: Sheet, run):
        """``run()`` over a sheet's part, with the parser's refusal turned into
        :class:`OpenError`.  ``run`` may be a generator function: its items are passed on."""
        return _guarded(sheet.part or f"sheet #{sheet.id}", run)

    def scan(self, sheet: "Sheet | str | SheetRef") -> SheetScan:
        """One full streamed pass over a sheet, done once and kept: the used range
        computed from the cells, merges, ordering and the issues met
        (:class:`~xlsx_agent.read.stream.SheetScan`)."""
        sheet = self._cell_sheet(sheet)
        key = sheet.part or f"#{sheet.id}"
        if key not in self._scans:
            self._scans[key] = _parsed(key, lambda: scan(self._reader(sheet)))
        return self._scans[key]

    def used_range(self, sheet: "Sheet | str | SheetRef") -> Area | None:
        """The used range: the smallest area holding every cell with a value or a formula,
        computed from the cells (a ``dimension`` is never trusted); ``None`` when there
        are none."""
        return self.scan(sheet).used

    def merges(self, sheet: "Sheet | str | SheetRef") -> list[Area]:
        """The sheet's merged areas, in the order the file gives them."""
        return list(self.scan(sheet).merges)

    def arrays(self, sheet: "Sheet | str | SheetRef") -> list[Area]:
        """The blocks the sheet's array formulas fill (a dynamic array's spill among them),
        each as its formula gives it, in reading order."""
        return list(self.scan(sheet).arrays)

    def hyperlinks(self, sheet: "Sheet | str | SheetRef") -> list[Hyperlink]:
        """The sheet's hyperlinks, in the file's order, each with the cells it covers and
        its ``url`` (an external target, from the sheet's relationships) or ``location`` (a
        place in the workbook).  Content: never followed, never in a message."""
        sheet = self._cell_sheet(sheet)
        links = self.scan(sheet).hyperlinks
        if not links:
            return []
        rels = self.package.relationships(sheet.part)
        out = []
        for link in links:
            rel = rels.get(link.rel_id) if link.rel_id is not None else None
            out.append(replace(link, url=rel.target) if rel is not None else link)
        return out

    def comments(self, sheet: "Sheet | str | SheetRef") -> list[Comment]:
        """The sheet's comments (its comments part, :mod:`xlsx_agent.read.comments`), in
        the file's order; read once."""
        sheet = self._cell_sheet(sheet)
        key = sheet.part or f"#{sheet.id}"
        if key not in self._comments:
            found: list[Comment] = []
            if sheet.part is not None:
                for rel in self.package.relationships(sheet.part).values():
                    if rel.is_external or rel.type.rpartition("/")[2] != "comments" \
                            or rel.target_part is None \
                            or not self.package.has_part(rel.target_part):
                        continue
                    part = rel.target_part
                    found = read_comments(_parsed(part, lambda: self.package.tree(part)))
                    break
            self._comments[key] = found
        return list(self._comments[key])

    def cells(self, sheet: "Sheet | str | SheetRef", *, formats: bool = False) -> Iterator[Value]:
        """Every cell the sheet stores, typed, in the order of the file, streamed (a cell
        with only a style comes as ``blank``).  ``formats`` puts each value's format code on
        it."""
        sheet = self._cell_sheet(sheet)
        return self._stream(sheet, lambda: self._reader(sheet, formats=formats).values())

    def read(self, address: "str | RangeAddress | TableAddress", *,
             formats: bool = False) -> Iterator[list[Value]]:
        """The rows of a cell or range (``Stock!A1:H40``, ``#3!B2``), each a list of
        values from the range's first column to its last, blanks filled in; whole columns
        and rows (``Stock!A:A``, ``Stock!2:5``) clipped to the used range.  A ``t:``
        address is read as the cells :meth:`resolve` gives it (a table's data rows, a
        column's, a row, a value; nothing when it has no data rows).

        Lazy: rows are read as they are asked for, and the pass stops after the range's
        last row.  The sheet is scanned once first (:meth:`scan`: the used range, and
        whether its rows are in order; a sheet whose rows are not has the range's cells
        gathered before the first row is given).  A dataset or ref address is not a range:
        ``ValueError``."""
        return self._read(address, formats=formats)

    def _read(self, address: "str | RangeAddress | TableAddress", *, formats: bool = False,
              columns: "set[int] | None" = None) -> Iterator[list[Value]]:
        """:meth:`read`, typing the cells of ``columns`` (sheet columns) only, for a pass
        that looks at no other (:func:`~xlsx_agent.read.stream.read_area`)."""
        if isinstance(address, str):
            address = parse_address(address)
        if isinstance(address, TableAddress):
            selected = self.resolve(address).range
            if selected is None:
                return iter(())
            address = selected
        if not isinstance(address, RangeAddress):
            raise ValueError(f"a {address.kind} address is read from a later milestone; "
                             "give a cell or range, as Stock!A1:H40")
        sheet = self._cell_sheet(self.sheet_of(address))
        found = self.scan(sheet)
        area = address.clipped(found.used)
        if area is None:
            return iter(())
        return self._stream(sheet, lambda: read_area(self._reader(sheet, formats=formats),
                                                     area, found.ordered, found.index,
                                                     columns))

    def _rows_apart(self, address: RangeAddress) -> bool:
        """Whether rows of the range's sheet far apart are read as well by a read for each
        as by one through them all: its reads start near their rows (it has a
        :class:`~xlsx_agent.read.stream.RowIndex`), and no cell read depends on one passed
        on the way (it has no shared formula, whose group's first cell the others name)."""
        index = self.scan(self._cell_sheet(self.sheet_of(address))).index
        return index is not None and not index.shared_formulas

    def cell(self, address: "str | RangeAddress | TableAddress", *,
             formats: bool = False) -> Value:
        """One cell's value (``Stock!F2``, or a ``t:`` value such as
        ``t:Products/SKU-0042/Price``); ``blank`` when the sheet stores nothing there."""
        if isinstance(address, str):
            address = parse_address(address)
        if isinstance(address, TableAddress):
            address = self.resolve(address).range
        if not isinstance(address, RangeAddress) or address.kind != "cell":
            raise ValueError("a cell address names one cell, as Stock!F2")
        return next(iter(self.read(address, formats=formats)))[0]

    def _read_sheets(self) -> list[Sheet]:
        holder = self._child("sheets")
        if holder is None:
            return []
        relationships = self.package.relationships(self.part)
        out = []
        for node in holder:
            if not isinstance(node.tag, str) or node.tag.rpartition("}")[2] != "sheet":
                continue
            rel_id = next((node.get(f"{{{ns}}}id") for ns in NS_R if node.get(f"{{{ns}}}id")), None)
            rel = relationships.get(rel_id) if rel_id else None
            part = rel.target_part if rel is not None and not rel.is_external else None
            if part is not None and not self.package.has_part(part):
                part = None
            kind = rel.type.rpartition("/")[2] if rel is not None else None
            kind = SHEET_TYPES.get(kind, kind)
            out.append(Sheet(name=node.get("name") or "", id=_int(node.get("sheetId")),
                             state=node.get("state") or "visible",
                             protected=part is not None and self._protected(part, kind), type=kind, part=part))
        return out

    def _protected(self, part: str, kind: str | None) -> bool:
        """Whether the sheet part's ``sheetProtection`` turns protection on, streamed: the
        element follows the cells, and a sheet is never parsed whole."""
        data = self.package.read(part) or b""
        if data[:2] not in (b"\xff\xfe", b"\xfe\xff") and b"\x00" not in data[:64] \
                and b"sheetProtection" not in data:
            return False  # the element's name is spelt out in any 8-bit encoding
        attribute = "content" if kind == "chartsheet" else "sheet"
        return _parsed(part, lambda: _scan_protection(data, attribute))


#: The sheet types that hold cells (a chart sheet and a dialog sheet hold none).
_CELL_SHEETS = ("worksheet", "macrosheet")


def _guarded(part: str, run) -> Iterator:
    try:
        yield from run()
    except etree.XMLSyntaxError:
        raise OpenError("damaged", f"{part} cannot be parsed as safe, well-formed XML") from None


def _scan_protection(data: bytes, attribute: str) -> bool:
    for _, node in iterparse_xml(data, tag=_PROTECTION + _ROWS):
        if node.tag in _PROTECTION:
            return _true(node.get(attribute))
        node.clear()
        while node.getprevious() is not None:
            del node.getparent()[0]
    return False


def _parsed(part: str, parse):
    """``parse()``, with the parser's refusal turned into :class:`OpenError`: the part's
    name only, since lxml's message may quote the part."""
    try:
        return parse()
    except etree.XMLSyntaxError:
        raise OpenError("damaged", f"{part} cannot be parsed as safe, well-formed XML") from None


def _true(value: str | None) -> bool:
    """An ``xsd:boolean``."""
    return value is not None and value.strip() in ("1", "true")


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None
