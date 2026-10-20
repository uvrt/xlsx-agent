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
    wb.expand("t:Products/*/Price")      # ['t:Products/SKU-0001/Price', ...]: by key
    wb.declare_region("Returns", "Returns!A3:F200", header_rows=2, headers=[...])
    wb.declare_tables([{"name": "checks", "sheets": "Area *", "header_row": 3,
                        "key": "Ref"}])   # the caller's tables (xlsx_agent.declare)

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

and, from 0.2, edited (:mod:`xlsx_agent.edit`): one history, atomic steps, a checked save::

    wb.set_cells({"Stock!D10": 12, "Stock!H10": {"date": "2026-10-20"},
                  "Stock!J10": {"formula": "=D10*2"}})   # typed, guarded, one step (E3)
    with wb.batch() as edit:             # one step, all or nothing
        wb.set_cells({...}); wb.set_cells({...})
    wb.undo(); wb.redo(); wb.version     # undo to the start: the bytes opened
    wb.validate()                        # the structure check, by code and address
    wb.save("stock.xlsx")                # reopened and read back first

Transitional and strict packages (``purl.oclc.org`` namespaces) are read alike; a strict one
is not written.
"""

from __future__ import annotations

import json
import os
import re
from collections import OrderedDict
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, BinaryIO, Iterator, Mapping, Sequence

from lxml import etree
from ooxml_edit.sheets import Area, Cell
from ooxml_edit.opc import OpcPackage
from ooxml_edit.tools import Dataset
from ooxml_edit.tools.limits import Limits
from ooxml_edit.xml import Element

from .address import (NameNotFound, RangeAddress, SheetRef, TableAddress, nearest,
                      parse_address, quote_sheet)
from .catalogue import DEFAULT_TOKENS, catalogue
from .compare import Diff, compare
from .dataset import table_dataset
from .derive import Derived, derive
from .find import Found, find_text
from .package import OpenError, open_package, read_source, workbook_part
from .profile import Column, ColumnStats, profile
from .read.names import DefinedName, read_names
from .schema import ValidationReport, validate
from .read.regions import Region, Regions, Selection, UnclearRegion, _overlaps
from .read.comments import Comment, read_comments
from .read.stream import Hyperlink, SheetReader, SheetScan, _area, read_area, scan
from .read.strings import SharedStrings
from .read.styles import NumberFormat, Styles
from .read.tables import TablePart, read_table
from .read.types import Problem, Value, format_language, reading_locale
from .redact import Redaction
from .xml import iterparse_xml

if TYPE_CHECKING:
    from .edit import Editor, EditReport, SaveReport
    from .edit.cells import Written

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
    relationship leads nowhere.  Under a redaction (:meth:`Workbook.redact`) whose rules
    match its name, ``name`` is the name redacted -- what results show and addresses
    take -- and :attr:`stored_name` the name as the file stores it."""

    name: str
    id: int | None
    state: str = "visible"
    protected: bool = False
    type: str | None = "worksheet"
    part: str | None = field(default=None, repr=False)
    stored: str | None = field(default=None, repr=False, compare=False)

    @property
    def visible(self) -> bool:
        return self.state == "visible"

    @property
    def stored_name(self) -> str:
        """The name as the file stores it (``name`` unless a redaction changed it): for
        the program, and for the caller's declarations, which name sheets as stored."""
        return self.stored if self.stored is not None else self.name


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
        #: The sheets withheld from the model (:meth:`withhold_sheets`), by name folded
        #: (as shown and as stored).
        self._withheld_sheets: frozenset[str] = frozenset()
        #: The terms redacted from every text read (:meth:`redact`), or ``None``.
        self._redaction: Redaction | None = None
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
        #: A sheet's problems without the cells declared tables hide, by sheet and areas.
        self._withheld_problems: dict = {}
        #: The history and the cell writer (:attr:`editor`), made when first needed.
        self._editor: "Editor | None" = None
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

    def to_bytes(self, kind: str | None = None) -> bytes:
        """The workbook as a file.  While nothing has changed -- or every change has been
        undone -- it is the bytes that were opened, byte for byte; a workbook opened from a
        package is that package's :meth:`~ooxml_edit.opc.OpcPackage.to_bytes`.  ``kind``
        (``"xlsx"``, ``"xlsm"``, ``"xltx"``, ``"xltm"``) writes it as that kind instead
        (:func:`xlsx_agent.edit.serialise`: a kind without macros is refused while the
        workbook holds a VBA project).  Unchecked: :meth:`save` reads the file back first."""
        if kind is not None and kind != self.kind:
            from .edit import serialise

            return serialise(self, kind)
        if self._data is not None and not self.package.changed_parts():
            return self._data
        return self.package.to_bytes()

    def save(self, target: "str | os.PathLike[str] | BinaryIO", *,
             kind: str | None = None) -> "SaveReport":
        """Write the workbook to ``target`` (a path or a binary file), checked first: the
        bytes are reopened and every cell written since the workbook was opened is read
        back with its type and value (:func:`xlsx_agent.edit.verify`; ``EditError``
        ``readback`` and nothing written when one is not).  ``kind`` is the path's
        extension's when it names a workbook kind, else the workbook's own; a ``.xlsx``
        (or ``.xltx``) is refused while the workbook holds a VBA project
        (``vba_project``).  Returns the :class:`~xlsx_agent.edit.SaveReport`: the kind,
        the size, the cells read back and the structure issues of what was saved."""
        from ooxml_common.kinds import kind_for

        from .edit import serialise, verify

        if kind is None:
            kind = kind_for(target, "excel") or self.kind
        data = serialise(self, kind)
        report = verify(self, data, kind)
        if isinstance(target, (str, os.PathLike)):
            with open(os.fspath(target), "wb") as handle:
                handle.write(data)
        else:
            target.write(data)
        return report

    # -- editing (0.2) ---------------------------------------------------------------------

    @property
    def editor(self) -> "Editor":
        """The workbook's :class:`~xlsx_agent.edit.Editor`: its history and cell writer,
        made the first time it is needed (:mod:`xlsx_agent.edit`)."""
        if self._editor is None:
            from .edit import Editor

            self._editor = Editor(self)
        return self._editor

    @property
    def history(self) -> Any:
        """ooxml-edit's :class:`~ooxml_edit.history.History` of the workbook's edits."""
        return self.editor.history

    @property
    def version(self) -> int:
        """The workbook's state number: 0 as opened, a new number for every step, and the
        number of the state undo or redo brings back (equal versions, equal content).  A
        failed step keeps it.  What a caller holds as the revision it last read."""
        return self.editor.history.version

    #: The same as :attr:`version`.
    revision = version

    def batch(self):
        """``with wb.batch() as edit:`` -- one step, all or nothing; ``edit.report`` the
        :class:`~xlsx_agent.edit.EditReport` once it closes (:mod:`xlsx_agent.edit`)."""
        return self.editor.batch()

    def savepoint(self):
        """Inside a batch, roll back what runs inside alone if it raises
        (:meth:`Editor.savepoint <xlsx_agent.edit.Editor.savepoint>`)."""
        return self.editor.savepoint()

    def undo(self) -> bool:
        """Back one step; whether there was one.  Undoing every step gives the bytes that
        were opened."""
        return self.editor.undo()

    def redo(self) -> bool:
        """Forward one undone step; whether there was one."""
        return self.editor.redo()

    @property
    def last_edit(self) -> "EditReport | None":
        """The last step's :class:`~xlsx_agent.edit.EditReport`; ``None`` before any."""
        return self._editor.last if self._editor is not None else None

    def validate(self, *, sheets: "Sequence[str] | None" = None) -> list[Problem]:
        """The structure check (:mod:`xlsx_agent.edit.structure`): what is out of order
        among the parts -- cell order, shared strings and their counts, styles, the
        dimension, tables, names, relationships and content types, the calculation chain --
        each by code with its first address and count, never content.  ``sheets`` (names)
        limits the cell checks to those sheets."""
        from .edit.structure import check_structure

        parts = None
        if sheets is not None:
            parts = [self._cell_sheet(name).part for name in sheets]
        return check_structure(self, sheets=parts)

    @property
    def calculation_status(self) -> str:
        """Whether the workbook's cached formula values are as last saved:
        ``"invalidated"`` when the workbook asks for a full recalculation on load
        (``fullCalcOnLoad``, which every edit sets) or was saved with manual calculation --
        its cached values may be stale -- and ``"cached_unverified"`` otherwise: as last
        saved, with no known invalidation, never checked to be current.  Read from the
        package, so a rollback or an undo restores it with the rest."""
        return "invalidated" if self.calculation else "cached_unverified"

    @property
    def _external_names(self) -> frozenset[str]:
        """The defined names (case folded) that refer to another workbook."""
        return frozenset(name.name.casefold() for name in self.regions._every_name
                         if name.reason == "external")

    @property
    def external_links(self) -> int:
        """How many external-link parts the workbook relates to (``externalLink``): other
        workbooks its formulas or names may refer to, never followed."""
        return sum(1 for rel in self.package.relationships(self.part).values()
                   if rel.type.rpartition("/")[2] == "externalLink")

    @property
    def strict(self) -> bool:
        """The workbook part is in strict OOXML's namespace (``purl.oclc.org``)."""
        return self._root.tag.startswith(f"{{{NS_MAIN[1]}}}")

    def set_cells(self, cells: "Mapping[str, Any] | Sequence[tuple[str, Any]]", *,
                  replace_formulas: bool = False) -> "Written":
        """Write cells by address, typed and guarded (:mod:`xlsx_agent.edit.cells`):
        ``{"Stock!D10": 12, "Stock!H10": {"date": "2026-10-20"}, "Stock!J10":
        {"formula": "=D10*2"}, "Stock!B4": None}`` -- text always literal, numbers,
        booleans, dates in the workbook's date system into a cell that shows one (or with
        ``"format"``), blanks that keep the cell's format, formulas only given as one.
        Every value and guard is checked before anything changes, and the writes are one
        step (or part of the open batch): :class:`~xlsx_agent.edit.EditError` with the
        cells refused, and nothing written, for any refusal.  A cell holding a formula is
        written over only with ``replace_formulas``.  Returns the
        :class:`~xlsx_agent.edit.cells.Written`: the cells, their old types, the shared
        formula groups handed on, the calculation status (``invalidated`` after any
        write)."""
        from .edit.cells import set_cells

        return set_cells(self, cells, replace_formulas=replace_formulas)

    def _forget(self) -> None:
        """Forget everything read from the package: it changed (an edit, a rollback, an
        undo or redo).  The application's settings stay: declared tables and regions, the
        redaction, the sheets withheld, the locale."""
        self._sheets = None
        self._strings = None
        self._styles = None
        self._scans.clear()
        self._comments.clear()
        self._table_parts_read = None
        self._catalogue = None
        self._queries.clear()
        self._finds.clear()
        self._withheld_problems.clear()
        self.regions.forget_package()

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
        """Every sheet, in the workbook's order (the order of its tabs); a sheet withheld
        from the model (:meth:`withhold_sheets`) is not one of them."""
        withheld = self._withheld_sheets
        return [s for s in self._every_sheet if s.name.casefold() not in withheld]

    @property
    def _every_sheet(self) -> list[Sheet]:
        """Every sheet the workbook part lists, withheld ones too: for what reads the file
        as stored (a defined name's ``localSheetId``, a formula's run of sheets), never
        for what a result shows."""
        if self._sheets is None:
            self._sheets = self._read_sheets()
        return list(self._sheets)

    @property
    def withheld_sheets(self) -> list[str]:
        """The names of the sheets withheld from the model (:meth:`withhold_sheets`), in
        the workbook's order: for the program that withheld them; no tool gives them."""
        withheld = self._withheld_sheets
        return [s.name for s in self._every_sheet if s.name.casefold() in withheld]

    @property
    def redaction(self) -> "Redaction | None":
        """The terms redacted from every text this workbook gives (:meth:`redact`), or
        ``None``; its :attr:`~xlsx_agent.redact.Redaction.redacted` counts the texts
        changed so far, for the program."""
        return self._redaction

    def redact(self, redaction: "Redaction | Sequence[Mapping]") -> None:
        """Redact terms from every text this workbook gives from now on (see
        :mod:`xlsx_agent.redact`): ``redaction`` a :class:`~xlsx_agent.redact.Redaction`
        or its declarations (``[{"redact": ["Acme Corp"], "with": "[redacted]"}]``).  The
        caller's setting, declared with the tables (:mod:`xlsx_agent.declare`), and set
        before anything is declared or kept: ``ValueError`` once a table, region, key or
        hidden sheet is, unless the rules are the ones set already.  Everything read
        before is forgotten."""
        from .redact import parse_rules
        if not isinstance(redaction, Redaction):
            redaction = parse_rules(redaction)
        if self._redaction is not None and self._redaction.spec == redaction.spec:
            return
        regions = self.regions
        if regions.declared or regions.configured or regions.declared_keys \
                or self._withheld_sheets or (self._redaction is not None):
            raise ValueError("a redaction is set before tables, regions, keys or hidden "
                             "sheets are declared, and once")
        self._set_redaction(redaction if redaction.rules else None)

    def _set_redaction(self, redaction: "Redaction | None") -> None:
        """Set the redaction, forgetting everything read with another."""
        self._redaction = redaction
        self._sheets = None
        self._strings = None
        self._styles = None
        self._scans.clear()
        self._comments.clear()
        self._table_parts_read = None
        self._catalogue = None
        self._queries.clear()
        self._finds.clear()
        self._withheld_problems.clear()
        self.regions = Regions(self)
        self._sheets = self._read_sheets()

    def _redacted(self, text: str) -> str:
        """``text`` as this workbook gives it: redacted (:meth:`redact`), or as it is."""
        return self._redaction.text(text) if self._redaction is not None and text else text

    def withhold_sheets(self, names: "Sequence[str]") -> None:
        """Withhold the sheets called ``names`` (exact names, case aside; more are added to
        those withheld already, and none is ever given back) from everything this
        workbook shows: they are no :attr:`sheets`, no address reaches their cells, and
        their tables and the defined names on them or local to them are not the
        workbook's; a formula elsewhere that reaches them keeps its value, not its text.
        The caller's setting, declared with the tables (``{"sheets": "Notes", "hidden":
        true}``, :mod:`xlsx_agent.declare`), which checks it; ``KeyError`` here for a name
        that is no sheet."""
        wanted = {name.casefold() for name in names}
        found = set()
        for sheet in self._every_sheet:
            folded = {sheet.name.casefold(), sheet.stored_name.casefold()}
            if sheet.name and folded & wanted:
                found |= folded
                wanted -= folded
        if wanted:
            raise KeyError("no such sheet")
        self._set_withheld_sheets(self._withheld_sheets | found)

    def _set_withheld_sheets(self, withheld: "frozenset[str]") -> None:
        """Set the sheets withheld (by name folded), forgetting what was gathered with
        another set: the catalogue, finds, queries and the regions' names and tables."""
        if withheld == self._withheld_sheets:
            return
        self._withheld_sheets = frozenset(withheld)
        self._catalogue = None
        self._queries.clear()
        self._finds.clear()
        self._withheld_problems.clear()
        self.regions.forget_withheld()

    def _sheet_withheld(self, name: str | None) -> bool:
        """Whether the sheet called ``name`` is withheld from the model."""
        return bool(name) and name.casefold() in self._withheld_sheets

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
        time this is asked.  The cells of a column a declared table hides from the model
        (:mod:`xlsx_agent.declare`) are left out: a problem met only there is not listed,
        and one met there too is counted and placed at its other cells only (the sheet is
        read once more for it, while the declarations stay as they are)."""
        out = [Problem(code) for code in self.calculation]
        out.extend(self.regions.problems())
        for sheet in self.sheets:
            if sheet.part is not None and sheet.type in _CELL_SHEETS:
                out.extend(self._sheet_problems(sheet))
        return out

    def _sheet_problems(self, sheet: Sheet) -> list[Problem]:
        found = self.scan(sheet)
        hidden = self.regions.withheld(sheet.name) if sheet.name else []
        if not hidden or not found.issues:
            return found.problems
        key = (sheet.part, tuple(hidden), self.locale)
        if key not in self._withheld_problems:
            counted: dict[str, list] = {}

            def count(code: str, value: Value) -> None:
                entry = counted.get(code)
                if entry is None:
                    counted[code] = [1, value.a]
                else:
                    entry[0] += 1

            for value in self.cells(sheet):
                if any(a.first.row <= value.row <= a.last.row
                       and a.first.column <= value.column <= a.last.column for a in hidden):
                    continue
                if value.issue is not None:
                    count(value.issue, value)
                if value.at_limit:
                    count("at_limit", value)
                if value.volatile:
                    count("volatile", value)
            self._withheld_problems[key] = [Problem(code, address, n) for code, (n, address)
                                            in sorted(counted.items())]
        return list(self._withheld_problems[key])

    # -- tables, names and regions ---------------------------------------------------------

    @property
    def tables(self) -> list[Region]:
        """The Excel tables (ListObjects), sheet by sheet in the workbook's order, each in
        the order its sheet relates to it; a table whose ``ref`` is not a range is left out
        and reported in :attr:`problems` (``bad_ref``, ``bad_table``), and so is one a
        declared table covers, which replaces it (:mod:`xlsx_agent.declare`)."""
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

    def expand(self, address: "str | TableAddress", *, key: str | None = None) -> list[str]:
        """What a ``t:`` address covers, one address by key per cell (per row, without a
        column), in the table's row order: ``t:Products/*/Price`` gives
        ``t:Products/SKU-0001/Price``, ... -- a coverage selector, for checking that a set
        of addresses (a derived value's sources, say, which name cells by key) holds every
        one::

            required = set(wb.expand("t:Products/*/Price"))
            used = {a for s in sources for a in wb.expand(s)}
            missing = required - used

        ``t:Products/*/*`` gives every value, ``t:Products/SKU-0042/*`` each of one row's,
        a key set each of its own, ``t:Products`` and ``t:Products/*`` every row.  The keys
        are the key column's (``key``, else the region's own); a row without a key, or with
        one another row shares, is refused (``no_key``, ``duplicate_key``), never named by
        a key that would not name it alone.  See
        :meth:`~xlsx_agent.read.regions.Regions.expand`."""
        return [a.text for a in self.regions.expand(address, key=key)]

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

    def declare_key(self, address: "str | TableAddress", column: str) -> Region:
        """Name ``column`` the key column of a table, name or region (``t:Products``) from
        then on: rows by key look their keys up in it, :meth:`expand` names cells by it,
        and a derived value's sources are written by it (``t:Products/{SKU-0003,
        SKU-0007}/Price``) -- the caller's setting, for a table with several columns that
        could be its key, or none that is.  A region declared with ``key=`` keeps its own.
        See :meth:`~xlsx_agent.read.regions.Regions.declare_key`."""
        return self.regions.declare_key(address, column)

    def declare_tables(self, declarations: Any) -> list[Region]:
        """Declare tables from the caller's configuration (plain data: a mapping, a list of
        them, or ``{"tables": [...]}``), usable as ``t:<name>`` from then on: sheets by
        name or glob, a table over several sheets read as one, the header row (or the
        row found by its text, ``{"match": [...]}``), the data rows' end, the key column,
        column ids and descriptions, columns hidden from the model, key-value tables; and
        whole sheets hidden from the model (``{"sheets": "Notes", "hidden": true}``,
        :meth:`withhold_sheets`).  Every declaration is checked against the workbook before
        any is kept (:class:`~xlsx_agent.declare.DeclarationError`, naming the entry and
        what is wrong); see :mod:`xlsx_agent.declare`.  The tables declared, in order."""
        from .declare import declare_tables
        return declare_tables(self, declarations)

    def declare_table(self, name: str, sheets: "str | Sequence[str]", **declaration: Any
                      ) -> Region:
        """One table, as :meth:`declare_tables` declares it: ``wb.declare_table("checks",
        "Area *", header_row=3, key="Ref", fields=[...])``."""
        found = self.declare_tables([{"name": name, "sheets": sheets, **declaration}])
        if not found:   # optional, and its sheets are not there
            raise KeyError(f"no sheet matches the sheets of t:{name}")
        return found[0]

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
        if not isinstance(address, TableAddress) or address.kind != "column":
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

    def dataset(self, address: "str | TableAddress", *, max_cells: int | None = None) -> Dataset:
        """The rows a ``t:`` address names as an ooxml-edit
        :class:`~ooxml_edit.tools.Dataset` -- columns, rows, each value's cell by key (or
        its cell address without a key) -- as a deck or document tool takes it with
        ``from: "t:Products"``: ``t:Products``, ``t:Products[Price]``,
        ``t:Products/{SKU-0003,SKU-0007}``; nothing coerced.  See
        :func:`~xlsx_agent.dataset.table_dataset`."""
        return table_dataset(self, address, max_cells=max_cells)

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
            for sheet in self._every_sheet:     # withheld ones too: a formula may name them
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
                    table = read_table(root, sheet=sheet.name, written_sheet=written, part=part,
                                       redact=self._redaction.text
                                       if self._redaction is not None else None)
                    if table is not None:
                        out.append(table)
            self._table_parts_read = out
        return list(self._table_parts_read)

    def _defined_names(self) -> list[DefinedName]:
        return read_names(self._child("definedNames"), self._every_sheet)

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
        """The shared-string table, read (streamed) the first time a cell needs it; each
        item redacted (:meth:`redact`)."""
        if self._strings is None:
            part = self._related("sharedStrings")
            self._strings = SharedStrings(self.package.read(part) if part else None,
                                          redact=self._redaction.text
                                          if self._redaction is not None else None)
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
            if self._sheet_withheld(found.name):
                found = self.sheet(found.name)   # refused as a sheet the workbook has not
        elif isinstance(sheet, SheetRef):
            found = self.sheet_of(sheet)
        else:
            found = self.sheet(sheet)
        if found.type not in _CELL_SHEETS:
            raise ValueError(f"sheet {quote_sheet(found.name) if found.name else found.id} "
                             f"is a {found.type or 'sheet of no known type'} and holds no cells")
        return found

    def _reader(self, sheet: Sheet, *, formats: bool = False,
                markdown: bool = False) -> SheetReader:
        data = self.package.read(sheet.part) if sheet.part is not None else None
        name = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
        return SheetReader(data or b"", sheet=name, hidden=not sheet.visible,
                           strings=self.strings, styles=self.styles, date1904=self.date1904,
                           locale=self.locale, formats=formats, grouping=self.grouping,
                           local_dates=self.local_dates, redaction=self._redaction,
                           cache=self.calculation_status,
                           external_names=self._external_names, markdown=markdown)

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
        place in the workbook).  Content: never followed, never in a message.  A link to a
        place on a sheet withheld from the model (its ``location`` names the sheet, or a
        defined name of it) is left out: it would say the sheet's name."""
        sheet = self._cell_sheet(sheet)
        links = self.scan(sheet).hyperlinks
        if links and self._withheld_sheets:
            links = [link for link in links if link.location is None
                     or not self.regions._text_names_withheld_sheet(link.location, sheet.name,
                                                                    set())]
        if not links:
            return []
        rels = self.package.relationships(sheet.part)
        out = []
        for link in links:
            rel = rels.get(link.rel_id) if link.rel_id is not None else None
            link = replace(link, url=rel.target) if rel is not None else link
            if self._redaction is not None:
                link = replace(link, url=self._redacted(link.url) if link.url else link.url,
                               location=self._redacted(link.location) if link.location
                               else link.location)
            out.append(link)
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
                    if self._redaction is not None:
                        found = [replace(c, text=self._redacted(c.text),
                                         author=self._redacted(c.author) if c.author
                                         else c.author) for c in found]
                    break
            self._comments[key] = found
        return list(self._comments[key])

    def cells(self, sheet: "Sheet | str | SheetRef", *, formats: bool = False,
              markdown: bool = False) -> Iterator[Value]:
        """Every cell the sheet stores, typed, in the order of the file, streamed (a cell
        with only a style comes as ``blank``).  ``formats`` puts each value's format code on
        it; ``markdown`` gives text as Markdown (as :meth:`read`)."""
        sheet = self._cell_sheet(sheet)
        values = self._stream(sheet, lambda: self._reader(sheet, formats=formats,
                                                          markdown=markdown).values())
        return (_as_markdown(value) for value in values) if markdown else values

    def read(self, address: "str | RangeAddress | TableAddress", *,
             formats: bool = False, markdown: bool = False) -> Iterator[list[Value]]:
        """The rows of a cell or range (``Stock!A1:H40``, ``#3!B2``), each a list of
        values from the range's first column to its last, blanks filled in; whole columns
        and rows (``Stock!A:A``, ``Stock!2:5``) clipped to the used range.  A ``t:``
        address is read as the cells :meth:`resolve` gives it (a table's data rows, a
        column's, a row, a value; nothing when it has no data rows); a key set's rows
        in the set's order (``t:Products/{SKU-0042,SKU-0007}/Price``).

        Lazy: rows are read as they are asked for, and the pass stops after the range's
        last row.  The sheet is scanned once first (:meth:`scan`: the used range, and
        whether its rows are in order; a sheet whose rows are not has the range's cells
        gathered before the first row is given).  A cell of a column a declared table hides
        from the model (:mod:`xlsx_agent.declare`) is given blank: every read of the
        sheet's cells goes through here.  Text comes with the workbook's redacted terms
        replaced (:meth:`redact`), as from :meth:`cells`.  A dataset or ref address is not a range:
        ``ValueError``.

        ``markdown``: every text value is given as Markdown (:mod:`xlsx_agent.markdown`):
        rich text's bold, italic and strikethrough runs as ``**``, ``*`` and ``~~``, lines
        of bullets and numbers as list items, every other character escaped (``*`` as
        ``\\*``) -- what ``set_cells``' ``{"markdown": ...}`` writes back as the same cell.
        A value a declared column's rule cleaned stays as cleaned."""
        rows = self._read(address, formats=formats, markdown=markdown)
        return _markdown_rows(rows) if markdown else rows

    def _read(self, address: "str | RangeAddress | TableAddress", *, formats: bool = False,
              columns: "set[int] | None" = None,
              markdown: bool = False) -> Iterator[list[Value]]:
        """:meth:`read`, typing the cells of ``columns`` (sheet columns) only, for a pass
        that looks at no other (:func:`~xlsx_agent.read.stream.read_area`)."""
        if isinstance(address, str):
            address = parse_address(address)
        if isinstance(address, TableAddress):
            selection = self.resolve(address)
            if selection.positions:
                return self._read_rows(selection.areas, formats=formats, columns=columns,
                                       markdown=markdown)
            selected = selection.range
            if selected is None:
                return iter(())
            address = selected
        if not isinstance(address, RangeAddress):
            raise ValueError(f"a {address.kind} address is read from a later milestone; "
                             "give a cell or range, as Stock!A1:H40")
        span = self.regions.span(address.sheet)
        if span is not None:   # a declared table read as one: its own rows and columns
            return span.read(self, address.area, formats=formats, columns=columns,
                             markdown=markdown)
        sheet = self._cell_sheet(self.sheet_of(address))
        found = self.scan(sheet)
        area = address.clipped(found.used)
        if area is None:
            return iter(())
        readers: list[SheetReader] = []

        def run() -> Iterator[list[Value]]:
            readers.append(self._reader(sheet, formats=formats, markdown=markdown))
            return read_area(readers[0], area, found.ordered, found.index, columns)

        rows = self._stream(sheet, run)
        withheld = [a for a in self.regions.withheld(sheet.name) if _overlaps(area, a)] \
            if sheet.name else []
        if withheld:
            rows = _withheld(rows, withheld)
        if sheet.name and self.regions.withholds:
            rows = self._formulas_withheld(rows, sheet.name, readers)
        return rows

    def _formulas_withheld(self, rows: Iterator[list[Value]], sheet: str,
                           readers: "list[SheetReader]") -> Iterator[list[Value]]:
        """``rows`` with the text of each formula that reaches a cell hidden from the
        model left out (:attr:`Value.formula_withheld
        <xlsx_agent.read.types.Value.formula_withheld>`; :meth:`Regions.formula_withheld
        <xlsx_agent.read.regions.Regions.formula_withheld>`).  A shared group is judged
        whole, by its first cell's text and the range it fills."""
        regions = self.regions
        for row in rows:
            groups = readers[0].shared_groups if readers else {}
            for value in row:
                if not value.formula:
                    continue
                anchor = value.shared_from or value.a
                text, fill = groups.get(anchor, (value.formula_text, None))
                if not text:
                    continue
                fill = _area(fill)
                if fill is None and value.shared_from:   # a group without its range
                    first = _area(anchor.rpartition("!")[2])
                    if first is not None:
                        fill = Area(first.first, Cell(max(first.first.row, value.row),
                                                      max(first.first.column, value.column)))
                origin = fill.first if fill is not None else Cell(value.row, value.column)
                if regions.formula_withheld(text, sheet, origin, fill):
                    value.formula_text = None
                    value.formula_withheld = True
            yield row

    def _read_rows(self, areas: "list[RangeAddress]", *, formats: bool = False,
                   columns: "set[int] | None" = None,
                   markdown: bool = False) -> Iterator[list[Value]]:
        """The rows of ``areas`` (one row each, on one sheet, alike in columns), in their
        order: one read through the rows they span, keeping theirs."""
        wanted = {a.area.first.row for a in areas}
        first, last = areas[0].area.first, areas[0].area.last
        span = RangeAddress(areas[0].sheet, Area(Cell(min(wanted), first.column),
                                                 Cell(max(wanted), last.column)))
        found = {}
        for row in self._read(span, formats=formats, columns=columns, markdown=markdown):
            if row and row[0].row in wanted:
                found[row[0].row] = row
        return iter([found[a.area.first.row] for a in areas])

    def _rows_apart(self, address: RangeAddress) -> bool:
        """Whether rows of the range's sheet far apart are read as well by a read for each
        as by one through them all: its reads start near their rows (it has a
        :class:`~xlsx_agent.read.stream.RowIndex`), and no cell read depends on one passed
        on the way (it has no shared formula, whose group's first cell the others name)."""
        if self.regions.span(address.sheet) is not None:
            return False
        index = self.scan(self._cell_sheet(self.sheet_of(address))).index
        return index is not None and not index.shared_formulas

    def cell(self, address: "str | RangeAddress | TableAddress", *,
             formats: bool = False, markdown: bool = False) -> Value:
        """One cell's value (``Stock!F2``, or a ``t:`` value such as
        ``t:Products/SKU-0042/Price``); ``blank`` when the sheet stores nothing there."""
        if isinstance(address, str):
            address = parse_address(address)
        if isinstance(address, TableAddress):
            address = self.resolve(address).range
        if not isinstance(address, RangeAddress) or address.kind != "cell":
            raise ValueError("a cell address names one cell, as Stock!F2")
        return next(iter(self.read(address, formats=formats, markdown=markdown)))[0]

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
        if self._redaction is not None:
            out = _shown_names(out, self._redaction)
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


def _as_markdown(value: Value) -> Value:
    """``value`` with its Markdown as ``v`` (a text value no rule cleaned)."""
    if value.t == "text" and value.markdown is not None and not value.cleaned:
        value.v = value.markdown
    return value


def _markdown_rows(rows: Iterator[list[Value]]) -> Iterator[list[Value]]:
    for row in rows:
        yield [_as_markdown(value) for value in row]


#: The sheet types that hold cells (a chart sheet and a dialog sheet hold none).
_CELL_SHEETS = ("worksheet", "macrosheet")


def _shown_names(sheets: list[Sheet], redaction: "Redaction") -> list[Sheet]:
    """``sheets`` with each name a redaction changes shown redacted, made unique (case
    aside) among every sheet's name, shown or stored, by `` (2)``, `` (3)``..."""
    taken = {s.name.casefold() for s in sheets}
    out = []
    for sheet in sheets:
        shown = redaction.text(sheet.name) if sheet.name else sheet.name
        if shown == sheet.name:
            out.append(sheet)
            continue
        name, n = shown, 1
        while name.casefold() in taken:
            n += 1
            name = f"{shown} ({n})"
        taken.add(name.casefold())
        out.append(replace(sheet, name=name, stored=sheet.name))
    return out


def _withheld(rows: Iterator[list[Value]], areas: list[Area]) -> Iterator[list[Value]]:
    """``rows`` with each cell of ``areas`` (columns hidden from the model by a declared
    table) given blank: its place, nothing of what it holds."""
    for row in rows:
        for n, value in enumerate(row):
            if any(a.first.row <= value.row <= a.last.row
                   and a.first.column <= value.column <= a.last.column for a in areas):
                row[n] = Value(value.sheet, value.row, value.column)
        yield row


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
