"""The open workbook: its package, its kind, its sheets.

``Workbook.open`` takes bytes, a path, a binary file or an ooxml-edit
:class:`~ooxml_edit.opc.OpcPackage` (a chart's embedded workbook, say), checks it as
:mod:`xlsx_agent.package` says, and reads the workbook part and each sheet's protection
(streamed: a sheet is never parsed whole)::

    wb = Workbook.open("inventory.xlsx")
    wb.sheets        # [Sheet(name='Stock', id=3, state='visible', protected=False, type='worksheet')]
    wb.to_bytes()    # the bytes that were opened, while nothing has changed

Transitional and strict packages (``purl.oclc.org`` namespaces) are read alike.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import BinaryIO

from lxml import etree
from ooxml_edit.opc import OpcPackage
from ooxml_edit.tools.limits import Limits
from ooxml_edit.xml import Element

from .package import OpenError, open_package, read_source, workbook_part
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

    def __init__(self, package: OpcPackage, *, part: str, kind: str, data: bytes | None = None) -> None:
        #: The ooxml-edit package: every part as the bytes that were read.
        self.package = package
        #: The workbook part (``xl/workbook.xml`` in every file Excel writes).
        self.part = part
        #: ``"xlsx"``, ``"xlsm"``, ``"xltx"`` or ``"xltm"``: what the workbook part declares.
        self.kind = kind
        self._data = data
        self._sheets: list[Sheet] | None = None

    # -- lifecycle -------------------------------------------------------------------------

    @classmethod
    def open(cls, source: "str | os.PathLike[str] | bytes | BinaryIO | OpcPackage", *,
             limits: Limits | None = None) -> "Workbook":
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
        with the part's name (``code`` ``"damaged"``), rather than at a later read."""
        if isinstance(source, OpcPackage):
            part, kind = workbook_part(source)
            workbook = cls(source, part=part, kind=kind)
        else:
            limits = limits or Limits()
            data = read_source(source, limits)
            package = open_package(data, limits)
            part, kind = workbook_part(package, source)
            workbook = cls(package, part=part, kind=kind, data=data)
        workbook._sheets = workbook._read_sheets()
        return workbook

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
        ``KeyError`` naming the sheets there are."""
        folded = name.casefold()
        for sheet in self.sheets:
            if sheet.name.casefold() == folded:
                return sheet
        raise KeyError(f"no sheet {name!r}; the sheets are {[sheet.name for sheet in self.sheets]}")

    @property
    def protected(self) -> bool:
        """The workbook's structure or windows are locked (``workbookProtection``)."""
        node = self._child("workbookProtection")
        if node is None:
            return False
        return any(_true(node.get(attribute)) for attribute in ("lockStructure", "lockWindows"))

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
