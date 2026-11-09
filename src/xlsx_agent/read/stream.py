"""The streaming pass over a sheet: every cell typed, the sheet never parsed whole.

A worksheet may hold a million rows (ROADMAP.md, "Design rules", 6), so it is read with
:func:`xlsx_agent.xml.iterparse_xml` (``parse_xml``'s safe options): each cell is typed as
its row ends, and the row is then dropped, so what is held at any moment is one row.  What
the pass meets on the way -- hidden columns (``cols`` comes before the cells), hidden rows,
array formulas' blocks, merges, data-validation rules and hyperlinks (the last three come
after the cells), how the rows are ordered -- is collected as it goes.

**Hyperlinks** (ECMA-376 Part 1, §18.3.1.47 ``hyperlink``): the cells each covers
(``ref``), its relationship id (``r:id``, an external target in the sheet's relationships)
and its ``location`` (a place in the workbook), as written.  Content: returned by a read
when asked for, never in a message.

**Data validation** (ECMA-376 Part 1, §18.3.1.32 ``dataValidation``): each rule's type
(ST_DataValidationType), the cells it covers (``sqref``) and its first formula, as text,
for a program only (formulas hold literals: never in a message).  A rule Excel keeps in
the worksheet's ``extLst`` (the ``x14:dataValidation`` of Microsoft's [MS-XLSX]
extensions, whose ``xm:sqref`` and ``x14:formula1/xm:f`` carry the same) is read alike.

Positions are computed as Excel computes them (ROADMAP.md, "Messy workbooks"): a row
without ``r`` is the row after the one before it, a cell without ``r`` the column after
the one before it in its row (the first column in a row's first cell).  A ``dimension``
is never trusted: :class:`SheetScan` computes the used range from the cells.

How a cell's value is typed is :mod:`xlsx_agent.read.types`'s; the cell types are
ECMA-376 Part 1, §18.18.11 (ST_CellType): ``b``, ``d``, ``e``, ``inlineStr``, ``n``,
``s``, ``str``.
"""

from __future__ import annotations

import io
import math
import re
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, field
from typing import Any, Callable, Iterator

from ooxml_common.zipparts import ZipPart
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell

from ..formula import is_external
from lxml import etree

from ..xml import SAFE_OPTIONS, iterparse_xml
from ooxml_edit.markdown import Run

from .strings import SharedStrings, decode_xstring, item_runs, item_text, text_markdown
from .styles import DATE_KINDS, NumberFormat, Styles

from ..address import _letters, format_cell
from .types import (TEXT_LIMIT, Problem, Value, looks_like_date, looks_like_number,
                    serial_text)

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")


def _tags(local: str) -> frozenset[str]:
    return frozenset(f"{{{ns}}}{local}" for ns in NS_MAIN)


_ROW, _C, _COL, _MERGE = _tags("row"), _tags("c"), _tags("col"), _tags("mergeCell")
_DIMENSION, _HYPERLINK = _tags("dimension"), _tags("hyperlink")
#: ``r:id``, transitional and strict.
_R_ID = tuple(f"{{{ns}}}id" for ns in (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships"))
_VALIDATION, _FORMULA1, _FORMULA2 = _tags("dataValidation"), _tags("formula1"), _tags("formula2")
#: [MS-XLSX]'s namespaces of the x14 data-validation extension and of its ``xm:`` children.
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"
_X14_VALIDATION = frozenset({f"{{{NS_X14}}}dataValidation"})
_X14_FORMULA1, _XM_F, _XM_SQREF = f"{{{NS_X14}}}formula1", f"{{{NS_XM}}}f", f"{{{NS_XM}}}sqref"
_X14_FORMULA2 = f"{{{NS_X14}}}formula2"
#: Conditional formats (ECMA-376 Part 1, §18.3.1.18 ``conditionalFormatting``), and the
#: ``x14:conditionalFormatting`` of [MS-XLSX] in the worksheet's ``extLst``.
_CONDITIONAL = _tags("conditionalFormatting")
_X14_CONDITIONAL = frozenset({f"{{{NS_X14}}}conditionalFormatting"})
_V, _F, _IS = _tags("v"), _tags("f"), _tags("is")
_STREAMED = tuple(_ROW | _COL | _MERGE | _DIMENSION | _VALIDATION | _X14_VALIDATION
                  | _HYPERLINK | _CONDITIONAL | _X14_CONDITIONAL)

_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})")
_INTEGER = re.compile(r"[+-]?[0-9]+")
_TIME = r"\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?"
_ISO_DATETIME = re.compile(rf"(?:(\d{{4}}-\d{{2}}-\d{{2}})(?:T({_TIME}))?|({_TIME}))"
                           r"(?:Z|[+-]\d{2}:\d{2})?")

#: The functions whose value changes at every recalculation (ROADMAP.md, "Values and
#: types"): a cached value of a formula that calls one is as of the last save.
VOLATILE_FUNCTIONS = ("NOW", "TODAY", "RAND", "OFFSET", "INDIRECT")
_VOLATILE = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?(?:" + "|".join(VOLATILE_FUNCTIONS)
                       + r")\s*\(", re.IGNORECASE)
_QUOTED = re.compile(r'"[^"]*"|\'[^\']*\'')

_COLUMNS: dict[str, int] = {}
_MISSING = object()
_CACHE_LIMIT = 1024

#: A sheet's :class:`RowIndex` notes where every this many rows start.
CHECKPOINT_ROWS = 128
#: A ``row`` start tag and the ``sheetData`` start tag, as bytes, with any prefix.
_ROW_TAG = re.compile(rb"<(?:[A-Za-z_][-\w.]*:)?row[\s/>]")
_SHEET_DATA_TAG = re.compile(rb"<([A-Za-z_][-\w.]*:)?sheetData(?:\s[^>]*)?>")

#: Whether a pass reads the rows that hold only styles from the bytes (:class:`_Plan`):
#: off, every row is parsed (the tests compare the two).
SKIM = True
#: How many bytes a pass looks ahead for such rows at a time: a read that stops early
#: never searches the whole part.
_WINDOW = 1 << 20
#: How many bytes the parser is given at a time (as ``iterparse`` reads them).
_PIECE = 1 << 15
#: How many bytes of a part kept compressed (:class:`SheetBytes`) are inflated at a time,
#: and how many a pass holds of it: its window is refilled to :data:`_FILL` bytes once
#: fewer than :data:`_AHEAD` are left ahead of the pass.  A pass searches only what it
#: holds, and a row cut by the window's end is parsed (the same cells as read from the
#: bytes), so these change what a pass holds, never what it reads.
_READ = 1 << 16
_FILL = 1 << 16
_AHEAD = 1 << 14
#: The most bytes before ``sheetData`` a streamed part's :class:`_Facts` keep as its header;
#: a part with more is read without a :class:`_Plan` or a :class:`RowIndex` (as it is).
_HEADER_LIMIT = 1 << 24
#: How many bytes of a streamed part are kept around each place its :class:`_Facts` look
#: at: a closing tag's prefix and spaces, a tag cut by a piece's end.
_CONTEXT = 1 << 12
_S = rb"[ \t\r\n]"
#: An attribute as a row of :class:`_Plan` may have it: a name (with a prefix or not) and
#: a value of printable ASCII, no reference in it (so as written is as read).
_ATTR = (rb"[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?" + _S + rb"*=" + _S
         + rb"*(?:\"[ !#-%'-;=-~]*\"|'[ -%(-;=-~]*')")
_ATTRS = re.compile(rb"([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)" + _S + rb"*=" + _S
                    + rb"*(?:\"([^\"]*)\"|'([^']*)')")
_ENCODING = re.compile(rb"<\?xml[^>]*?encoding" + _S + rb"*=" + _S + rb"*[\"']([^\"']*)[\"']")
_PATTERNS: dict[bytes, tuple["re.Pattern[bytes]", ...]] = {}
_LETTERS: dict[bytes, int] = {}


def _skim_patterns(prefix: bytes) -> "tuple[re.Pattern[bytes], ...]":
    """A row that holds only styles, written with ``prefix`` (``b""`` or ``b"x:"``): the
    row (group ``row``, 1), its attributes (2) and cells (3, ``None`` for an empty row),
    each cell ``c``
    with ``r`` (column letters and row number, as Excel writes them) and at most ``s``;
    the same with whitespace before it; and one such cell's letters, digits and style."""
    found = _PATTERNS.get(prefix)
    if found is None:
        p = re.escape(prefix)
        cell = (rb"<" + p + rb"c" + _S + rb"+r" + _S + rb"*=" + _S
                + rb"*(?:\"[A-Z]{1,3}[0-9]{1,7}\"|'[A-Z]{1,3}[0-9]{1,7}')(?:" + _S + rb"+s" + _S
                + rb"*=" + _S + rb"*(?:\"[0-9]{1,9}\"|'[0-9]{1,9}'))?" + _S + rb"*/>")
        written = (rb"(?P<row><" + p + rb"row((?:" + _S + rb"+" + _ATTR + rb")*)" + _S
                   + rb"*(?:/>|>((?:" + _S + rb"*" + cell + rb")*)" + _S + rb"*</" + p
                   + rb"row" + _S + rb"*>))")
        one = re.compile(rb"<" + p + rb"c" + _S + rb"+r" + _S + rb"*=" + _S
                         + rb"*[\"']([A-Z]{1,3})([0-9]{1,7})[\"'](?:" + _S + rb"+s" + _S
                         + rb"*=" + _S + rb"*[\"']([0-9]{1,9})[\"'])?")
        found = _PATTERNS[prefix] = (re.compile(written), re.compile(_S + rb"*" + written),
                                     one)
    return found


@dataclass(frozen=True)
class _Plan:
    """Where a part's rows lie, for a pass that reads a row holding only styles (each
    cell ``<c r="B9" s="3"/>``, or no cell: the rows a sheet formatted to its last row has
    by the hundred thousand) from the bytes instead of parsing it.  ``begin`` is where the
    rows start (after the ``sheetData`` start tag), ``end`` where its end tag starts;
    ``row``, ``following`` and ``cell`` match such a row, one after whitespace, and its
    cells (:func:`_skim_patterns`);
    ``prefixes`` those bound in the header, which a row's attribute may use.

    Made only for a part where every ``<`` is a tag (no comment, CDATA section, document
    type or processing instruction), in UTF-8, with one ``sheetData`` of the main
    namespace and no namespace bound inside it: a row matched there is a row element,
    well-formed, read as the parser reads it (a row's attributes are checked for repeats
    and unbound prefixes, which the parser refuses, by :func:`_row_attributes`)."""

    begin: int
    end: int
    row: "re.Pattern[bytes]"
    following: "re.Pattern[bytes]"
    cell: "re.Pattern[bytes]"
    prefixes: frozenset[bytes]
    #: A row's start with its number first (``<row r="9"``), and its end tag.
    head: "re.Pattern[bytes]"
    end_tag: bytes
    #: Every byte up to ``begin``: the namespaces, ``cols``, ``dimension``.
    header: bytes
    #: Whether the part has a row that holds only styles: a part with none is parsed
    #: whole, as before (``iterparse``); its plan still serves :func:`merges_only`.
    skims: bool = True

    @classmethod
    def of(cls, facts: "_Facts") -> "_Plan | None":
        if not facts.plain or facts.utf16:
            return None
        declared = facts.declared
        if declared is not None and declared.lower() not in (b"utf-8", b"utf8",
                                                             b"us-ascii", b"ascii"):
            return None
        header = facts.header
        if header is None or facts.end is None or facts.xmlns:
            return None
        prefix = facts.prefix
        bound = re.findall(rb"xmlns(?::([A-Za-z_][\w.-]*))?" + _S + rb"*=" + _S
                           + rb"*[\"']([^\"']*)[\"']", header)
        main = [uri for name, uri in bound if name + (b":" if name else b"") == prefix]
        if len(main) != 1 or main[0].decode("ascii", "replace") not in NS_MAIN:
            return None
        row, following, cell = _skim_patterns(prefix)
        return cls(len(header), facts.end, row, following, cell,
                   frozenset(name for name, _ in bound if name) | {b"xml"},
                   re.compile(rb"<" + re.escape(prefix) + rb'row r="([1-9][0-9]{0,6})"'),
                   b"</" + prefix + b"row>", header, facts.skims)


@dataclass(frozen=True)
class _Facts:
    """What a pass needs to know of a whole part before it starts, found once per part:

    ``plain`` -- every ``<`` is a tag (no ``<!``: no comment, CDATA section or document
    type), at most one ``<?`` (the declaration, at the start or after a BOM) and
    ``sheetData`` written twice; ``utf16`` -- a UTF-16 BOM begins it; ``declared`` -- the
    declaration's encoding; ``header`` -- every byte through the ``sheetData`` start tag
    (with ``prefix``, the tag's), ``None`` without one; ``end`` -- where the closing tag
    starts; ``xmlns`` and ``merges`` -- whether ``xmlns`` or ``mergeCell`` may stand among
    the rows (``True`` when not known: a part's bytes held whole are searched, a streamed
    one's from its first ``sheetData`` to its second).  A streamed part whose header is
    over :data:`_HEADER_LIMIT`, or whose tags' prefixes or spaces are longer than
    :data:`_CONTEXT`, gets no header or end: it is then parsed, as any part without
    :class:`_Plan`, with the same result."""

    plain: bool
    utf16: bool
    declared: bytes | None
    header: bytes | None
    prefix: bytes
    end: int | None
    xmlns: bool
    merges: bool
    #: Whether the part may hold ``sheetProtection``: UTF-16, a NUL in its first 64 bytes,
    #: or the name in its bytes (spelt out so in any 8-bit encoding).
    protection: bool
    #: Whether a row that holds only styles (:func:`_skim_patterns`) stands among the
    #: rows.  A streamed part's pieces are searched each, with :data:`_CONTEXT` bytes of
    #: the one before: a longer such row cut by a piece's end may go unseen (``False``:
    #: the part is parsed whole, the same cells).
    skims: bool
    #: The part's :class:`_Plan`, made of these.
    plan: "_Plan | None" = field(default=None, compare=False)

    @classmethod
    def of(cls, source: "SheetBytes") -> "_Facts":
        facts = cls._of_bytes(source.data) if source.entry is None \
            else cls._of_stream(source.chunks())
        object.__setattr__(facts, "plan", _Plan.of(facts))
        return facts

    @classmethod
    def _of_bytes(cls, data: bytes) -> "_Facts":
        plain = not (b"<!" in data or data.count(b"<?") > 1
                     or data.find(b"<?") not in (-1, 0, 3) or data.count(b"sheetData") != 2)
        declared = _ENCODING.match(data, 3 if data.startswith(b"\xef\xbb\xbf") else 0)
        header = end = None
        prefix = b""
        xmlns = merges = True
        skims = False
        start = _SHEET_DATA_TAG.search(data)
        if start is not None:
            header, prefix = data[:start.end()], start.group(1) or b""
            closing = re.compile(rb"</" + re.escape(prefix) + rb"sheetData" + _S + rb"*>")
            found = closing.search(data, start.end())
            if found is not None:
                end = found.start()
                xmlns = data.find(b"xmlns", start.end(), end) >= 0
                merges = data.find(b"mergeCell", start.end(), end) >= 0
                skims = _skim_patterns(prefix)[0].search(data, start.end(), end) is not None
        protection = data[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in data[:64] \
            or _PROTECTION_NAME in data
        return cls(plain, data.startswith((b"\xfe\xff", b"\xff\xfe")),
                   declared.group(1) if declared is not None else None,
                   header, prefix, end, xmlns, merges, protection, skims)

    @classmethod
    def _of_stream(cls, chunks: Iterator[bytes]) -> "_Facts":
        bang = False
        pis = 0
        pi_first = -1
        places: list[int] = []      # where "sheetData" stands, the first three
        xmlns = merges = False
        head: list[bytes] | None = []   # the bytes before the header is known
        head_size = 0
        header: bytes | None = None
        prefix = b""
        end: int | None = None
        closing: bytes | None = None    # what follows the second "sheetData", until ">"
        first = b""
        head64 = b""
        protection = False
        skims = False
        tail = b""
        offset = 0                      # where ``buffer`` starts in the part
        for chunk in chunks:
            if len(head64) < 64:
                head64 += chunk[:64 - len(head64)]
                if head64[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in head64:
                    protection = True
            if len(first) < 4:
                first = (first + chunk)[:4]
            buffer = tail + chunk if tail else chunk
            new = len(tail)
            if not protection and buffer.find(_PROTECTION_NAME,
                                              max(new - len(_PROTECTION_NAME) + 1, 0)) >= 0:
                protection = True

            if not bang and buffer.find(b"<!", max(new - 1, 0)) >= 0:
                bang = True
            if pis < 2:
                at = buffer.find(b"<?", max(new - 1, 0))
                while at >= 0 and pis < 2:
                    if pis == 0:
                        pi_first = offset + at
                    pis += 1
                    at = buffer.find(b"<?", at + 2)
            first_here = len(places)
            if len(places) < 3:
                at = buffer.find(b"sheetData", max(new - 8, 0))
                while at >= 0 and len(places) < 3:
                    places.append(offset + at)
                    at = buffer.find(b"sheetData", at + 9)

            if head is not None:
                head.append(chunk)
                head_size += len(chunk)
                if places:
                    joined = b"".join(head)
                    close = joined.find(b">", places[0] + 9)
                    if close >= 0:
                        start = _SHEET_DATA_TAG.search(joined, 0, close + 1)
                        if start is not None and start.end() == close + 1:
                            header, prefix = joined[:close + 1], start.group(1) or b""
                        head = None
                    elif head_size > _HEADER_LIMIT:
                        head = None
                elif head_size > _HEADER_LIMIT:
                    head = None

            # Among the rows: between the first "sheetData" and the second.
            if places and not (xmlns and merges):
                low = max(places[0] - offset, max(new - 8, 0))
                high = places[1] - offset if len(places) > 1 else len(buffer)
                if low < high:
                    if not xmlns and buffer.find(b"xmlns", low, high + 4) >= 0:
                        xmlns = True
                    if not merges and buffer.find(b"mergeCell", low, high + 8) >= 0:
                        merges = True

            if first_here < 2 <= len(places):
                # The second "sheetData" came in this piece: the closing tag's name?
                at = places[1] - offset
                opened = at - 2 - len(prefix)
                if header is not None and opened >= 0 and buffer[opened:at] == b"</" + prefix:
                    closing, end = buffer[at + 9:at + 9 + _CONTEXT], offset + opened
            elif closing is not None and b">" not in closing and len(closing) < _CONTEXT:
                closing += chunk[:_CONTEXT - len(closing)]

            if header is not None and not skims and (end is None or offset < end):
                low = max(len(header) - offset, 0)
                high = end - offset if end is not None else len(buffer)
                if low < high and _skim_patterns(prefix)[0].search(buffer, low, high):
                    skims = True

            tail = buffer[-_CONTEXT:]
            offset += len(buffer) - len(tail)
        # The closing tag ends at its first ">", after spaces only.
        if closing is None or re.match(_S + rb"*>", closing) is None:
            end = None
        plain = not bang and pis <= 1 and pi_first in (-1, 0, 3) and len(places) == 2
        declared = None
        if header is not None:
            found = _ENCODING.match(header, 3 if header.startswith(b"\xef\xbb\xbf") else 0)
            declared = found.group(1) if found is not None else None
        if header is None or end is None:
            xmlns = merges = True
        return cls(plain, first.startswith((b"\xfe\xff", b"\xff\xfe")), declared, header,
                   prefix, end, xmlns, merges, protection, skims and end is not None)


class SheetBytes:
    """A sheet part's bytes as a pass reads them: ``bytes`` held whole (a part an edit
    wrote), or the part's compressed ZIP entry (:class:`ooxml_common.zipparts.ZipPart`),
    inflated a piece at a time by each pass -- a sheet of hundreds of megabytes is then
    never held whole.  ``size`` is the part's size either way."""

    __slots__ = ("data", "entry", "size")

    def __init__(self, source: "bytes | ZipPart | None") -> None:
        if isinstance(source, ZipPart):
            self.data: bytes = b""
            self.entry: ZipPart | None = source
            self.size = source.size
        else:
            self.data, self.entry = source or b"", None
            self.size = len(self.data)

    def __bool__(self) -> bool:
        return self.size > 0

    def stream(self, offset: int = 0):
        """A binary file of the bytes from ``offset`` on."""
        if self.entry is not None:
            return self.entry.open(offset)
        stream = io.BytesIO(self.data)
        stream.seek(offset)
        return stream

    def chunks(self, offset: int = 0, tap: "_Tags | None" = None) -> Iterator[bytes]:
        """The bytes from ``offset`` on: a held part's at once, a streamed part's
        :data:`_READ` at a time; each streamed piece shown to ``tap`` first."""
        if self.entry is None:
            if offset < len(self.data):
                yield self.data[offset:] if offset else self.data
            return
        with self.entry.open(offset) as reader:
            while True:
                piece = reader.read(_READ)
                if not piece:
                    if tap is not None:
                        tap.done = True
                    return
                if tap is not None:
                    tap.feed(piece)
                yield piece


class _Window:
    """A part's bytes from an offset on, for a pass to search: all of them for a part held
    whole (``base`` 0, never refilled); for a streamed one, the bytes from ``base`` on, read
    on by :meth:`cover` as the pass moves (those before the pass are let go)."""

    def __init__(self, source: SheetBytes, offset: int, tap: "_Tags | None" = None) -> None:
        self._chunks: "Iterator[bytes] | None"
        if source.entry is None:
            self.data, self.base, self._chunks = source.data, 0, None
        else:
            self.data, self.base, self._chunks = b"", offset, source.chunks(offset, tap)

    def cover(self, pos: int) -> None:
        """Hold the bytes from ``pos`` on: refilled to :data:`_FILL` of them (fewer at the
        part's end) once fewer than :data:`_AHEAD` are held."""
        chunks = self._chunks
        if chunks is None or self.base + len(self.data) - pos >= _AHEAD:
            return
        kept = self.data[pos - self.base:]
        pieces, total = [kept], len(kept)
        while total < _FILL:
            piece = next(chunks, None)
            if piece is None:
                self._chunks = None
                break
            pieces.append(piece)
            total += len(piece)
        self.data, self.base = b"".join(pieces), pos

    def rest(self, pos: int) -> Iterator[bytes]:
        """The bytes from ``pos`` to the part's end, in pieces of :data:`_PIECE` or more."""
        data = self.data[pos - self.base:]
        if self._chunks is None:
            if data:
                yield data
            return
        if data:
            yield data
        yield from self._chunks


def _row_attributes(written: bytes, prefixes: frozenset[bytes]) -> "dict[bytes, bytes] | None":
    """A matched row's attributes by name, as the parser reads them; ``None`` when it would
    refuse them (a name given twice, a prefix not bound) or one binds a namespace."""
    out: dict[bytes, bytes] = {}
    count = 0
    for name, double, single in _ATTRS.findall(written):
        count += 1
        prefix, colon, _ = name.partition(b":")
        if prefix == b"xmlns" or (colon and prefix not in prefixes):
            return None
        out[name] = double or single
    return out if len(out) == count else None


#: How many :class:`_Template` a reader keeps (a sheet has a few kinds of row).
_TEMPLATES = 64


@dataclass(frozen=True)
class _Template:
    """A row read from the bytes, kept to read the rows alike but for their number
    (:meth:`SheetReader._known`): whether it is ``hidden``, its ``cells`` on the grid as
    ``(column, style-only cell)`` in the file's order, and whether their columns
    increase."""

    hidden: bool
    cells: tuple
    increasing: bool


class _StyleOnly:
    """A ``c`` element with only a style, as a row read from the bytes gives it: what a
    cell's reader asks of it (its ``t``, its ``s``, its children: none)."""

    __slots__ = ("_style",)
    tag = f"{{{NS_MAIN[0]}}}c"

    def __init__(self, style: str | None) -> None:
        self._style = style

    def get(self, name: str, default: Any = None) -> Any:
        return self._style if name == "s" and self._style is not None else default

    def __iter__(self) -> Iterator:
        return iter(())

    def __len__(self) -> int:
        return 0


def _column(letters: str) -> int:
    index = _COLUMNS.get(letters)
    if index is None:
        index = 0
        for char in letters.upper():
            index = index * 26 + ord(char) - 64
        _COLUMNS[letters] = index
    return index


def _letters_column(letters: bytes) -> int:
    index = _LETTERS[letters] = _column(letters.decode())
    return index


def is_volatile(formula: str) -> bool:
    """Whether a formula's text calls one of :data:`VOLATILE_FUNCTIONS` (outside its
    string literals and quoted sheet names)."""
    return bool(_VOLATILE.search(_QUOTED.sub("", formula)))


@dataclass(frozen=True)
class Validation:
    """One data-validation rule: ``type`` its ST_DataValidationType as written (``"list"``,
    ``"whole"``, ``"decimal"``...; ``"none"`` when the file gives none), ``areas`` the
    cells it covers (``sqref``; a reference that is not one is left out), ``formula1`` its
    first formula as text (a list's ``"Red,Green"`` literal, a range or a name), for a
    program only.  ``origin`` is ``"standard"`` (the worksheet's ``dataValidations``) or
    ``"x14"`` (its ``extLst``); ``operator`` the ``operator`` attribute as written
    (``None`` when absent: ECMA-376's default is ``between``), ``allow_blank`` its
    ``allowBlank``, ``formula2`` its second formula (a ``between`` rule's upper bound), for
    a program only; ``sqref`` the cells as written."""

    type: str
    areas: tuple[Area, ...]
    formula1: str | None = field(default=None, repr=False)
    origin: str = "standard"
    operator: str | None = None
    allow_blank: bool = False
    formula2: str | None = field(default=None, repr=False)
    sqref: str = ""


@dataclass(frozen=True)
class ConditionalFormat:
    """One conditional format (ECMA-376 Part 1, §18.3.1.18): ``sqref`` the cells as written,
    ``origin`` ``"standard"`` or ``"x14"`` (the worksheet's ``extLst``), ``rules`` its
    ``cfRule`` elements in priority order, each as ``(local name, attributes, text,
    children)`` (:func:`_canonical`), for a program only: formulas are content."""

    sqref: str
    origin: str = "standard"
    rules: tuple = field(default=(), repr=False)


@dataclass(frozen=True)
class Hyperlink:
    """One hyperlink as the sheet writes it: ``area`` the cells it covers, ``rel_id`` its
    relationship (an external target), ``location`` a place in the workbook (``Sheet2!A1``,
    a defined name), and ``url`` the relationship's target, which
    :meth:`Workbook.hyperlinks <xlsx_agent.workbook.Workbook.hyperlinks>` fills in; each
    ``None`` when not given.  Content, for a result only: kept off the repr."""

    area: Area
    rel_id: str | None = field(default=None, repr=False)
    location: str | None = field(default=None, repr=False)
    url: str | None = field(default=None, repr=False)


@dataclass
class SheetScan:
    """What one full pass over a sheet found.

    ``used`` is the used range computed from the cells that hold a value or a formula
    (``None`` for a sheet with none; a cell with only a style is blank and does not count);
    ``dimension`` the ``dimension`` as written, never trusted; ``cells`` how many cells
    hold a value or a formula; ``merges`` the merged areas; ``hidden_rows`` how many rows
    are hidden; ``ordered`` whether rows and cells come in the order of the sheet;
    ``issues`` each value issue (``no_cached_value``, ``number_as_text``...), ``at_limit``,
    ``volatile`` and ``external_formula`` (a formula that refers to another workbook) by
    code, as ``(count, first address)``; ``validations`` the sheet's
    data-validation rules, in the file's order; ``arrays`` the blocks the array formulas
    (and dynamic arrays' spills) fill, each its formula's ``ref``, in reading order;
    ``hyperlinks`` the sheet's hyperlinks, in the file's order."""

    sheet: str
    used: Area | None = None
    dimension: str | None = None
    cells: int = 0
    merges: list[Area] = field(default_factory=list)
    hidden_rows: int = 0
    ordered: bool = True
    issues: dict[str, tuple[int, str]] = field(default_factory=dict)
    validations: list[Validation] = field(default_factory=list)
    arrays: list[Area] = field(default_factory=list)
    hyperlinks: list[Hyperlink] = field(default_factory=list)
    #: Where the rows start, for a later read to start near them (an ordered sheet's;
    #: ``None`` otherwise).
    index: "RowIndex | None" = field(default=None, repr=False, compare=False)

    @property
    def problems(self) -> list[Problem]:
        """The issues as :class:`~xlsx_agent.read.types.Problem`, by code."""
        return [Problem(code, address, count) for code, (count, address) in sorted(self.issues.items())]

    def _count(self, code: str, value: Value) -> None:
        seen = self.issues.get(code)
        self.issues[code] = (1, value.a) if seen is None else (seen[0] + 1, seen[1])


@dataclass(frozen=True)
class Checkpoint:
    """Where a row element starts in a sheet part (``offset``, in bytes) and what a pass
    from the top knows there: the row number before it (``row``: a row without ``r`` is
    the next), the last cell's row and column, and how many hidden rows it has met."""

    offset: int
    row: int
    last_row: int
    last_column: int
    hidden: int


class RowIndex:
    """Where a sheet's rows start: a :class:`Checkpoint` every :data:`CHECKPOINT_ROWS`
    rows, so that a read of later rows starts at the last checkpoint before them rather
    than at the top (an ordered sheet only: its cells come in reading order, so none after
    a checkpoint is above it).  The part is parsed from there as its ``header`` (every byte
    up to and including the ``sheetData`` start tag: the namespaces, ``cols``,
    ``dimension``) followed by the bytes from the checkpoint on: the same document without
    the rows before.  Made by :func:`scan` from the marks its pass noted; ``None`` there
    for a part whose bytes it cannot match to the rows the parser met."""

    def __init__(self, header: bytes, checkpoints: list[Checkpoint], hidden: list[int],
                 shared_formulas: bool = False) -> None:
        self.header = header
        self.checkpoints = checkpoints
        #: The hidden rows in the order the pass met them.
        self.hidden = hidden
        #: Whether the sheet has a shared formula group: a read through rows notes each
        #: group's first cell for the cells after it, which a read started below misses.
        self.shared_formulas = shared_formulas
        self._last_rows = [c.last_row for c in checkpoints]

    def before(self, row: int) -> Checkpoint | None:
        """The last checkpoint with every cell before it above ``row``; ``None`` when the
        first is not (start at the top)."""
        k = bisect_left(self._last_rows, row) - 1
        return self.checkpoints[k] if k > 0 else None

    @classmethod
    def build(cls, source: "SheetBytes", reader: "SheetReader") -> "RowIndex | None":
        """The index of a part from what a whole pass over it noted (``reader.marks``):
        each checkpoint's row element found by its start tag in the bytes, the
        ``n``-th tag being the ``n``-th row the parser met.  ``None`` when that cannot be
        relied on: a comment, CDATA section, document type or processing instruction
        (whose text may hold a tag), ``sheetData`` written other than once as a start tag
        and once as an end tag, or as many tags as rows not found."""
        marks = reader.marks
        if not marks or len(marks) < 2:
            return None
        facts = reader.facts()
        if not facts.plain or facts.header is None:
            return None
        wanted = {mark[0]: mark for mark in marks}
        checkpoints = []
        count = 0
        tags = reader._tags
        if tags is not None and tags.done and not tags.failed \
                and all(n % CHECKPOINT_ROWS == 0 for n in wanted):
            if tags.count != reader.row_elements:
                return None
            for k, at in enumerate(tags.every):
                mark = wanted.get(k * CHECKPOINT_ROWS)
                if mark is not None:
                    checkpoints.append(Checkpoint(at, *mark[1:]))
            if len(checkpoints) != len(marks):
                return None
            return cls(facts.header, checkpoints, list(reader.hidden_order),
                       bool(reader._shared))
        if tags is not None and tags.failed:
            return None
        for at in _row_tags(source, len(facts.header)):
            if at < 0:
                return None
            mark = wanted.get(count)
            if mark is not None:
                checkpoints.append(Checkpoint(at, *mark[1:]))
            count += 1
        if count != reader.row_elements or len(checkpoints) != len(marks):
            return None
        return cls(facts.header, checkpoints, list(reader.hidden_order),
                   bool(reader._shared))


def _row_tags(source: "SheetBytes", begin: int) -> Iterator[int]:
    """Where each row start tag (:data:`_ROW_TAG`) stands from ``begin`` on, as
    ``finditer`` over the whole part finds them; a streamed part's pieces each searched,
    with a tag cut by a piece's end carried into the next (a tag holds one ``<``, its
    first byte).  ``-1`` then stops it: a name longer than :data:`_CONTEXT` was cut."""
    if source.entry is None:
        for match in _ROW_TAG.finditer(source.data, begin):
            yield match.start()
        return
    carried = b""
    offset = begin
    for chunk in source.chunks(begin):
        buffer = carried + chunk if carried else chunk
        last = 0
        for match in _ROW_TAG.finditer(buffer):
            yield offset + match.start()
            last = match.end()
        opened = buffer.rfind(b"<", last)
        carried = buffer[opened:] if opened >= 0 and _NAME_SO_FAR.fullmatch(buffer, opened) \
            else b""
        if len(carried) > _CONTEXT:
            yield -1
            return
        offset += len(buffer) - len(carried)


#: A tag's ``<`` and as much of a name as a piece may end in.
_NAME_SO_FAR = re.compile(rb"<[-\w.:]*")
_PROTECTION_NAME = b"sheetProtection"


class _Tags:
    """The row start tags (:data:`_ROW_TAG`) a pass from the top meets in a streamed part,
    counted as its pieces go by (:meth:`SheetBytes.chunks`'s ``tap``) from ``begin`` on --
    what :func:`_row_tags` would find, without inflating the part again for
    :meth:`RowIndex.build`: ``every``, where each :data:`CHECKPOINT_ROWS`-th stands, and
    ``count``.  ``done`` once the last piece went by; ``failed`` as :func:`_row_tags`
    stops (a name longer than :data:`_CONTEXT` cut by a piece's end)."""

    def __init__(self, begin: int, offset: int) -> None:
        self.begin = begin
        self.offset = offset
        self.carried = b""
        self.count = 0
        self.every: list[int] = []
        self.done = self.failed = False

    def feed(self, piece: bytes) -> None:
        start = self.offset
        self.offset += len(piece)
        if self.failed or self.offset <= self.begin:
            return
        if start < self.begin:
            piece, start = piece[self.begin - start:], self.begin
        carried = self.carried
        buffer = carried + piece if carried else piece
        base = start - len(carried)
        last = 0
        every = self.every
        for match in _ROW_TAG.finditer(buffer):
            if self.count % CHECKPOINT_ROWS == 0:
                every.append(base + match.start())
            self.count += 1
            last = match.end()
        opened = buffer.rfind(b"<", last)
        self.carried = buffer[opened:] if opened >= 0 \
            and _NAME_SO_FAR.fullmatch(buffer, opened) else b""
        if len(self.carried) > _CONTEXT:
            self.failed = True


class _Pieces:
    """An iterator of byte pieces as a binary file, for the parser."""

    def __init__(self, pieces: Iterator[bytes]) -> None:
        self._pieces = pieces
        self._piece = b""
        self._at = 0

    def read(self, size: int = -1) -> bytes:
        if self._at >= len(self._piece):
            self._piece, self._at = next(self._pieces, b""), 0
            if not self._piece:
                return b""
        if size is None or size < 0:
            size = len(self._piece)
        out = self._piece[self._at:self._at + size]
        self._at += len(out)
        return out


class _Resumed:
    """A part's bytes as a binary file read from a checkpoint: ``header``, then the part
    (:class:`SheetBytes`) from ``offset`` on, read in pieces (never held whole)."""

    def __init__(self, header: bytes, source: "SheetBytes", offset: int) -> None:
        self._header = memoryview(header)
        self._at = 0
        self._rest = source.stream(offset)

    def read(self, size: int = -1) -> bytes:
        header = self._header
        if self._at < len(header):
            end = len(header) if size is None or size < 0 else min(len(header), self._at + size)
            out = bytes(header[self._at:end])
            self._at = end
            if size is None or size < 0:
                return out + self._rest.read()
            return out
        return self._rest.read(size)


class SheetReader:
    """One sheet's cells, streamed from the part's bytes.

    ``sheet`` is the sheet as an address writes it; ``hidden`` whether the sheet is hidden
    (every value from it is then marked ``hidden``); ``strings``, ``styles``, ``date1904``
    and ``locale`` (with its reader settings ``grouping`` and ``local_dates``,
    :mod:`xlsx_agent.read.types`) are the workbook's.  ``formats`` puts each value's format
    code on it.
    A reader makes one pass at a time; each pass starts from the top."""

    def __init__(self, data: "bytes | ZipPart | SheetBytes", *, sheet: str, hidden: bool,
                 strings: SharedStrings,
                 styles: Styles, date1904: bool, locale: Any, formats: bool = False,
                 grouping: bool = True, local_dates: bool = False,
                 redaction: Any = None, cache: str = "cached_unverified",
                 external_names: "frozenset[str]" = frozenset(),
                 markdown: bool = False) -> None:
        #: The part's bytes, held or streamed (:class:`SheetBytes`).
        self.source = data if isinstance(data, SheetBytes) else SheetBytes(data)
        self.sheet = sheet
        self.hidden = hidden
        self.strings = strings
        self.styles = styles
        self.date1904 = date1904
        self.locale = locale
        self.grouping = grouping
        self.local_dates = local_dates
        self.formats = formats
        #: Give each text value its Markdown too (:attr:`Value.markdown
        #: <xlsx_agent.read.types.Value.markdown>`, :mod:`xlsx_agent.markdown`).
        self.markdown = markdown
        #: The workbook's redaction (:mod:`xlsx_agent.redact`): applied to every text (the
        #: shared strings come redacted), a number format's code, and a formula's text,
        #: which is withheld when it holds a match.
        self.redaction = redaction
        #: The status of every cached formula value the workbook holds
        #: (:attr:`Workbook.calculation_status <xlsx_agent.Workbook.calculation_status>`):
        #: ``cached_unverified`` or ``invalidated``; a formula without one is ``absent``.
        self.cache = cache
        #: The defined names (case folded) that refer to another workbook: a formula using
        #: one is ``external`` as one with a workbook index is.
        self.external_names = external_names
        #: Hidden columns as sorted, non-overlapping ``(first, last)`` spans, known once
        #: the pass has passed ``cols``.
        self.hidden_columns: list[tuple[int, int]] = []
        #: Hidden rows met so far.
        self.hidden_rows: set[int] = set()
        self.merges: list[Area] = []
        self.validations: list[Validation] = []
        #: The conditional formats met, in the file's order (standard, then ``x14``).
        self.conditional_formats: list[ConditionalFormat] = []
        #: The blocks of the array formulas met so far, and the sheet's hyperlinks.
        self.arrays: list[Area] = []
        self.hyperlinks: list[Hyperlink] = []
        self.dimension: str | None = None
        self.ordered = True
        #: Set to a list by :func:`scan`: the pass then notes a checkpoint every
        #: :data:`CHECKPOINT_ROWS` rows, and the hidden rows in the order met, from which
        #: :class:`RowIndex` is made.
        self.marks: list[tuple[int, int, int, int, int]] | None = None
        self.hidden_order: list[int] = []
        #: How many row elements the pass met.
        self.row_elements = 0
        self._shared: dict[str, tuple[str, bool, bool]] = {}
        #: Each shared formula group met, by its first cell's address: the group's text and
        #: the range it fills (its ``ref``), for what its cells reach
        #: (:meth:`Regions.formula_withheld <xlsx_agent.read.regions.Regions.formula_withheld>`).
        self.shared_groups: dict[str, tuple[str, str | None]] = {}
        self._format_cache: dict[str | None, NumberFormat] = {}
        # Each shared string's text and what it looks like, once per pass (as many as the
        # table, which is held already); dates by serial, at most _CACHE_LIMIT at a time.
        self._string_cache: dict[int, tuple[str, str | None, bool]] = {}
        self._date_cache: dict[tuple, str | None] = {}
        #: The workbook's :class:`_Facts` by part and this part's key (:meth:`facts`), or
        #: ``None``.
        self.plans: "tuple[dict, str] | None" = None
        self._facts: "_Facts | None" = None
        #: The row tags the last pass from the top counted (:class:`_Tags`), or ``None``.
        self._tags: "_Tags | None" = None
        self._plan: Any = _MISSING
        #: The rows read from the bytes kept as templates (:meth:`_learn`), and a
        #: :class:`_StyleOnly` per style.
        self._templates: dict[bytes, _Template] = {}
        self._styled: dict[str | None, _StyleOnly] = {}

    # -- the pass ----------------------------------------------------------------------

    def values(self, *, blanks: bool = True) -> Iterator[Value]:
        """Every cell in the sheet, typed, in the order of the file (a cell with only a
        style comes as ``blank``).  ``blanks`` off: a cell with no child (no value, formula
        or inline text) is passed over untyped -- placed, never made a value -- unless its
        ``t`` is ``s`` (typed: a shared string with no index is ``missing_string``); so
        are the rows read from the bytes (:class:`_Plan`)."""
        value = self.value
        for row, column, row_hidden, node in self._walk(blanks=blanks):
            yield value(row, column, row_hidden, node)

    def facts(self) -> "_Facts":
        """The part's :class:`_Facts`, found once per reader, or taken from ``plans`` (the
        workbook's, by part) when given."""
        if self._facts is None:
            plans, key = self.plans if self.plans is not None else ({}, "")
            found = plans.get(key)
            if found is None:
                found = plans[key] = _Facts.of(self.source)
            self._facts = found
        return self._facts

    def plan(self) -> "_Plan | None":
        """The part's :class:`_Plan` (``None``: every row is parsed)."""
        if not SKIM or not self.source:
            return None
        if self._plan is _MISSING:
            self._plan = self.facts().plan
        return self._plan   # type: ignore[return-value]

    def _events(self, start: "tuple[RowIndex, Checkpoint] | None"):
        """In the order of the file: ``(element, None, None)`` for each element of
        :data:`_STREAMED` the parser ends; for each row read from the bytes instead,
        ``(None, number, template)`` when it is a row met before but for its number
        (:class:`_Template`), else ``(None, attributes, cells)`` (its attributes by name,
        its cells' bytes or ``None``)."""
        source = self.source
        plan = self.plan()
        # A pass from the top that will index the part (scan's) counts the row tags of a
        # streamed part as it goes, so the index need not inflate it again.
        parsed_whole = plan is None or not plan.skims \
            or (start is not None and len(start[0].header) != plan.begin)
        tap = None
        self._tags = None
        if start is None and self.marks is not None and source.entry is not None:
            facts = self.facts()
            if facts.plain and facts.header is not None:
                # The parser is fed the part from its start; the window from the rows'.
                tap = self._tags = _Tags(len(facts.header),
                                         0 if parsed_whole or plan is None else plan.begin)
        if parsed_whole:
            stream: Any
            if tap is not None:
                stream = _Pieces(source.chunks(0, tap))
            else:
                stream = source.stream() if start is None else \
                    _Resumed(start[0].header, source, start[1].offset)
            for _, node in iterparse_xml(stream, tag=_STREAMED):
                yield node, None, None
            return
        assert plan is not None   # not parsed whole: there is a plan
        parser = etree.XMLPullParser(events=("end",), tag=_STREAMED, **SAFE_OPTIONS)
        parser.feed(plan.header)
        for _, node in parser.read_events():
            yield node, None, None
        at, end = (plan.begin if start is None else start[1].offset), plan.end
        search, match, prefixes = plan.row.search, plan.following.match, plan.prefixes
        known = self._known
        window = _Window(source, at, tap)
        # Positions below are in ``data``, the window's bytes, which start at ``base`` in
        # the part: a part held whole is all there; a streamed one's window is refilled
        # as the pass moves on, holding what lies ahead (rows cut by its end are found
        # whole once it is refilled, or parsed).
        while at < end:
            window.cover(at)
            data, base = window.data, window.base
            pos, stop = at - base, min(end - base, len(data))
            if pos >= stop:   # the part ended before its closing tag: it was not read so
                raise ValueError("a sheet part shorter than when it was first read")
            reach = min(stop, pos + _WINDOW)
            found = search(data, pos, reach)
            fed = reach if found is None else found.start()
            while pos < fed:   # in pieces, each drained: a few rows are held at a time
                piece = min(fed, pos + _PIECE)
                parser.feed(data[pos:piece])
                for _, node in parser.read_events():
                    yield node, None, None
                pos = piece
            while found is not None:
                again = known(data, pos, stop, plan)
                while again is not None:   # rows met before: no pattern matched
                    pos = again[0]
                    yield None, again[1], again[2]
                    again = known(data, pos, stop, plan)
                if found.start("row") < pos:
                    found = match(data, pos, stop)
                    if found is None:
                        break
                attributes = _row_attributes(found.group(2), prefixes)
                if attributes is None:   # the parser's to refuse
                    parser.feed(data[pos:found.end()])
                    for _, node in parser.read_events():
                        yield node, None, None
                else:
                    self._learn(found, attributes, plan)
                    yield None, attributes, found.group(3)
                pos = found.end()
                found = match(data, pos, stop)
            at = base + pos
        for rest in window.rest(end):
            parser.feed(rest)
            for _, node in parser.read_events():
                yield node, None, None
        parser.close()
        for _, node in parser.read_events():
            yield node, None, None

    def _known(self, data: bytes, pos: int, end: int,
               plan: "_Plan") -> "tuple[int, int, _Template] | None":
        """The row starting at ``pos`` when it is a row met before but for its number:
        where it ends, its number and its :class:`_Template`; else ``None``.  The row is
        its bytes up to its first ``>`` (ending ``/>``) or its first end tag; with its
        number (``r``, first, as Excel writes it) put as ``#`` it is a template's key, and
        a template is made only of a row in which the number stands nowhere else and
        ``#`` nowhere: so the bytes are that row's with another number."""
        templates = self._templates
        if not templates:
            return None
        head = plan.head.match(data, pos, end)
        if head is None:
            return None
        close = data.find(b">", head.end(), end)
        if close < 0:
            return None
        if data[close - 1] == 0x2F:   # "/>"
            stop = close + 1
        else:
            stop = data.find(plan.end_tag, close, end)
            if stop < 0:
                return None
            stop += len(plan.end_tag)
        written = data[pos:stop]
        digits = head.group(1)
        if b"#" in written:
            return None
        template = templates.get(written.replace(digits + b'"', b'#"'))
        if template is None:
            return None
        number = int(digits)
        return (stop, number, template) if number <= MAX_ROW else None

    def _learn(self, found: "re.Match[bytes]", attributes: "dict[bytes, bytes]",
               plan: "_Plan") -> None:
        """Keep the row ``found`` matched as a :class:`_Template` when it may be one: its
        number written first, as Excel writes it (no zero before it, on the grid), every
        cell's reference in its row, the number nowhere else and ``#`` nowhere."""
        written = found.group("row")
        digits = attributes.get(b"r")
        head = plan.head.match(written)
        if digits is None or head is None or head.group(1) != digits or b"#" in written \
                or int(digits) > MAX_ROW:
            return
        cells = plan.cell.findall(found.group(3) or b"")
        if any(cell_digits != digits for _, cell_digits, _ in cells) \
                or written.count(digits + b'"') != len(cells) + 1:
            return
        hidden = attributes.get(b"hidden")
        placed = []
        for letters, _, style in cells:
            column = _LETTERS.get(letters) or _letters_column(letters)
            if column <= MAX_COLUMN:
                placed.append((column, self._style_only(style)))
        columns = [column for column, _ in placed]
        templates = self._templates
        if len(templates) >= _TEMPLATES:
            templates.clear()
        templates[written.replace(digits + b'"', b'#"')] = _Template(
            hidden is not None and _true(hidden.decode()), tuple(placed),
            all(a < b for a, b in zip(columns, columns[1:])))

    def _style_only(self, style: bytes) -> "_StyleOnly":
        key = style.decode() if style else None
        found = self._styled.get(key)
        if found is None:
            found = self._styled[key] = _StyleOnly(key)
        return found

    def _walk(self, start: "tuple[RowIndex, Checkpoint] | None" = None, *,
              blanks: bool = True):
        """``(row, column, row hidden, c element)`` for each cell.  The pass asks for the
        ``end`` of each row (and of ``col``, ``mergeCell``, ``dimension``), takes the row's
        cells, then drops the row: one row is held at a time.  ``start``: a sheet's
        :class:`RowIndex` and one of its checkpoints, where the pass starts instead of the
        top, as it would be there.

        A row that holds only styles is read from the bytes where the part allows
        (:class:`_Plan`), its cells given as :class:`_StyleOnly`, read as the parsed cell
        is.  ``blanks`` off: a cell with no child and no ``t="s"`` (blank, no issue) is
        placed (it counts for the order and the next cell's column) but not given."""
        if not self.source:
            return
        row = 0
        last_row = last_column = 0
        hidden_rows = self.hidden_rows
        hidden_spans: list[tuple[int, int]] = []
        match_ref = _REF.fullmatch
        columns = _COLUMNS
        cell_tags, row_tags = _C, _ROW
        marks, hidden_order = self.marks, self.hidden_order
        elements = 0
        if start is not None:
            index, checkpoint = start
            row, last_row, last_column = checkpoint.row, checkpoint.last_row, checkpoint.last_column
            hidden_rows.update(index.hidden[:checkpoint.hidden])
        for node, skimmed, written in self._events(start):
            if node is None and type(written) is _Template:
                # A row met before but for its number: its cells are in it, in the order
                # the first was found in.
                if marks is not None and elements % CHECKPOINT_ROWS == 0:
                    marks.append((elements, row, last_row, last_column, len(hidden_order)))
                elements += 1
                self.row_elements = elements
                row = skimmed
                row_hidden = written.hidden
                if row_hidden and row not in hidden_rows:
                    hidden_rows.add(row)
                    hidden_order.append(row)
                placed = written.cells
                if not placed:
                    continue
                if row < last_row or (row == last_row and placed[0][0] <= last_column) \
                        or not written.increasing:
                    self.ordered = False
                if blanks:
                    for column, blank in placed:
                        yield row, column, row_hidden, blank
                last_row, last_column = row, placed[-1][0]
                continue
            if node is None:
                # A row read from the bytes: placed as a parsed one is, its cells (each
                # with letters and digits, no ``$``) given as style-only cells.
                if marks is not None and elements % CHECKPOINT_ROWS == 0:
                    marks.append((elements, row, last_row, last_column, len(hidden_order)))
                elements += 1
                self.row_elements = elements
                attributes = skimmed
                written_r, written_hidden = attributes.get(b"r"), attributes.get(b"hidden")
                number = _int(written_r.decode()) if written_r is not None else None
                row = number if number is not None and 1 <= number <= MAX_ROW else row + 1
                row_hidden = written_hidden is not None and _true(written_hidden.decode())
                if row_hidden and row not in hidden_rows:
                    hidden_rows.add(row)
                    hidden_order.append(row)
                if not written:
                    continue
                row_digits = str(row).encode()
                for written_letters, cell_digits, style in self._cells(written):
                    column = _LETTERS.get(written_letters) or _letters_column(written_letters)
                    if cell_digits == row_digits:
                        cell_row, cell_hidden = row, row_hidden
                    else:
                        cell_row = int(cell_digits)
                        cell_hidden = row_hidden if cell_row == row else cell_row in hidden_rows
                    if not (1 <= cell_row <= MAX_ROW and 1 <= column <= MAX_COLUMN):
                        continue
                    if cell_row < last_row or (cell_row == last_row and column <= last_column):
                        self.ordered = False
                    last_row, last_column = cell_row, column
                    if blanks:
                        yield cell_row, column, cell_hidden, self._style_only(style)
                continue
            tag = node.tag
            if tag in row_tags:
                if marks is not None and elements % CHECKPOINT_ROWS == 0:
                    marks.append((elements, row, last_row, last_column, len(hidden_order)))
                elements += 1
                self.row_elements = elements
                number = _int(node.get("r"))
                row = number if number is not None and 1 <= number <= MAX_ROW else row + 1
                row_hidden = _true(node.get("hidden"))
                if row_hidden and row not in hidden_rows:
                    hidden_rows.add(row)
                    hidden_order.append(row)
                column = 0
                # A cell's ``r`` is nearly always its column's letters and this row's
                # number: those are looked up, the rest matched.
                digits = str(row)
                cut = -len(digits)
                for cell in node:
                    if cell.tag not in cell_tags:
                        continue
                    ref = cell.get("r")
                    if ref and ref.endswith(digits) and ref[:cut] in columns:
                        column, cell_row, cell_hidden = columns[ref[:cut]], row, row_hidden
                    elif (match := match_ref(ref) if ref else None) is not None:
                        letters = match.group(1)
                        column = columns.get(letters) or _column(letters)
                        cell_row = int(match.group(2))
                        cell_hidden = row_hidden if cell_row == row else cell_row in hidden_rows
                    else:
                        cell_row, column, cell_hidden = row, column + 1, row_hidden
                    if not (1 <= cell_row <= MAX_ROW and 1 <= column <= MAX_COLUMN):
                        continue
                    if cell_row < last_row or (cell_row == last_row and column <= last_column):
                        self.ordered = False
                    last_row, last_column = cell_row, column
                    if blanks or len(cell) or cell.get("t") == "s":
                        yield cell_row, column, cell_hidden, cell
                node.clear()
                parent = node.getparent()
                if parent is not None:
                    while node.getprevious() is not None:
                        del parent[0]
            elif tag in _COL:
                if _true(node.get("hidden")):
                    first, end = _int(node.get("min")), _int(node.get("max"))
                    if first is not None and end is not None and 1 <= first <= end:
                        hidden_spans.append((first, min(end, MAX_COLUMN)))
                        self.hidden_columns = _merge_spans(hidden_spans)
            elif tag in _MERGE:
                area = _area(node.get("ref"))
                if area is not None:
                    self.merges.append(area)
            elif tag in _DIMENSION:
                self.dimension = node.get("ref")
            elif tag in _HYPERLINK:
                area = _area(node.get("ref"))
                if area is not None:
                    rel_id = next((node.get(a) for a in _R_ID if node.get(a)), None)
                    self.hyperlinks.append(Hyperlink(area, rel_id, node.get("location") or None))
            elif tag in _VALIDATION:
                self.validations.append(_standard_validation(node))
            elif tag in _CONDITIONAL or tag in _X14_CONDITIONAL:
                self.conditional_formats.append(_conditional(node, tag in _X14_CONDITIONAL))
            elif tag in _X14_VALIDATION:
                self.validations.append(_x14_validation(node))

    def _cells(self, written: bytes) -> "list[tuple[bytes, bytes, bytes]]":
        """A skimmed row's cells: each one's column letters, row digits and style."""
        plan = self._plan
        assert isinstance(plan, _Plan)
        return plan.cell.findall(written)

    def column_hidden(self, column: int) -> bool:
        spans = self.hidden_columns
        if not spans:
            return False
        k = bisect_right(spans, (column, MAX_COLUMN + 1)) - 1
        return k >= 0 and spans[k][0] <= column <= spans[k][1]

    def _format(self, style: str | None) -> NumberFormat:
        number_format = self._format_cache.get(style)
        if number_format is None:
            number_format = self._format_cache[style] = self.styles.format(_int(style) or 0)
        return number_format

    # -- one cell ----------------------------------------------------------------------

    def value(self, row: int, column: int, row_hidden: bool, node) -> Value:
        """The typed value of a ``c`` element at ``row``, ``column``."""
        hidden = self.hidden or row_hidden or (bool(self.hidden_columns) and self.column_hidden(column))
        value = Value(self.sheet, row, column, hidden=hidden)
        kind = node.get("t") or "n"
        v_text = f_node = is_node = None
        for child in node:
            tag = child.tag
            if tag in _V:
                v_text = child.text or ""
            elif tag in _F:
                f_node = child
            elif tag in _IS:
                is_node = child
        if f_node is not None:
            self._formula(value, f_node)
        if self.formats:
            value.format = self._format(node.get("s")).code
            if self.redaction is not None and value.format:
                value.format = self.redaction.text(value.format)

        if kind == "n":
            self._number(value, v_text, node)
        elif kind == "s":
            self._shared_string(value, v_text)
        elif kind == "inlineStr":
            if is_node is not None:
                self._text(value, item_text(is_node),
                           item_runs(is_node) if self.markdown else None)
            elif v_text is not None:
                self._text(value, decode_xstring(v_text))
        elif kind == "str":
            if v_text is not None:     # an ST_Xstring: Excel writes q_x0001_r (measured)
                self._text(value, decode_xstring(v_text))
        elif kind == "b":
            flag = (v_text or "").strip()
            if flag in ("1", "true"):
                value.v, value.t = True, "bool"
            elif flag in ("0", "false"):
                value.v, value.t = False, "bool"
            elif v_text is not None and not value.formula:
                value.issue = "bad_value"
        elif kind == "e":
            if v_text is not None and v_text.strip():
                value.v, value.t = v_text.strip(), "error"
        elif kind == "d":
            self._iso(value, v_text)
        else:  # not an ST_CellType: read as a number, as the default type is
            self._number(value, v_text, node)
        if value.formula:
            if value.t == "blank" and value.issue is None:
                value.issue = "no_cached_value"
            value.cache = "absent" if value.issue == "no_cached_value" else self.cache
        return value

    @staticmethod
    def shared_index(node) -> int | None:
        """The shared string a ``t="s"`` cell's ``v`` names, as :meth:`value` reads it
        (``None`` for none or no number)."""
        v_text = None
        for child in node:
            if child.tag in _V:
                v_text = child.text or ""
        return _int(v_text)

    def _shared_string(self, value: Value, v_text: str | None) -> None:
        index = _int(v_text)
        cached = self._string_cache.get(index) if index is not None else None
        if cached is None:
            try:
                if index is None or index < 0:
                    raise IndexError
                text = self.strings[index]
            except IndexError:
                value.issue = "missing_string"
                return
            cached = self._string_cache[index] = (text, self._looks(text), len(text) >= TEXT_LIMIT)
        value.v, value.t, value.at_limit = cached[0], "text", cached[2]
        if not value.formula:
            value.issue = cached[1]
        if self.markdown and index is not None:   # it is: a string was found by it
            value.markdown = self.strings.markdown(index)

    def _looks(self, text: str) -> str | None:
        """``"number_as_text"`` or ``"date_as_text"`` when text looks like one, in the
        locale."""
        if looks_like_number(text, self.locale, grouping=self.grouping):
            return "number_as_text"
        if looks_like_date(text, self.locale, local_dates=self.local_dates):
            return "date_as_text"
        return None

    def _text(self, value: Value, text: str, runs: "tuple[Run, ...] | None" = None) -> None:
        if self.markdown:
            redact = self.redaction.text if self.redaction is not None else None
            value.markdown = text_markdown(runs if runs is not None else (Run(text),),
                                           redact) if text else ""
        if self.redaction is not None and text:
            text = self.redaction.text(text)
        value.v, value.t = text, "text"
        if len(text) >= TEXT_LIMIT:
            value.at_limit = True
        if not value.formula:
            value.issue = self._looks(text)

    def _number(self, value: Value, v_text: str | None, node) -> None:
        if not v_text:
            return
        text = v_text.strip()
        number: int | float
        if text.isdigit() and text.isascii():   # most numbers: a whole one, no sign
            number = int(text)
        else:
            try:
                if not text.isascii():
                    raise ValueError
                number = int(text) if _INTEGER.fullmatch(text) else float(text)
            except ValueError:
                value.issue = "bad_value"
                return
            if isinstance(number, float) and not math.isfinite(number):
                value.issue = "bad_value"
                return
        value.v, value.t = number, "number"
        style = node.get("s")
        number_format = self._format_cache.get(style) or self._format(style)
        kind = number_format.kind
        if kind == "general" or kind == "number" or kind == "text":
            return
        if kind == "percent":
            value.t = "percent"
        elif kind in DATE_KINDS:
            key = (number, kind, number_format.elapsed)
            iso = self._date_cache.get(key, _MISSING)
            if iso is _MISSING:
                if len(self._date_cache) >= _CACHE_LIMIT:
                    self._date_cache.clear()
                iso = self._date_cache[key] = serial_text(float(number), kind, self.date1904,
                                                          number_format.elapsed)
            if iso is None:
                value.issue = "date_out_of_range"
            else:
                value.v, value.t, value.serial = iso, kind, float(number)
        elif kind == "ambiguous":
            value.issue = "ambiguous_format"
        elif kind == "unknown":
            value.issue = "unknown_format"

    def _iso(self, value: Value, v_text: str | None) -> None:
        """A ``t="d"`` cell: an ISO 8601 date, time or both (§18.17.4.1), kept as given."""
        if v_text is None or not v_text.strip():
            return
        text = v_text.strip()
        match = _ISO_DATETIME.fullmatch(text)
        if match is None:
            value.issue = "bad_value"
            return
        value.v = text
        value.t = "time" if match.group(3) else "datetime" if match.group(2) else "date"

    def note_shared(self, row: int, column: int, node) -> None:
        """What typing a cell would note of its formula for the cells after it: the first
        cell of a shared formula group (whose other cells name it)."""
        f_node = None
        for child in node:
            if child.tag in _F:
                f_node = child   # the last, as value() takes it
        if f_node is not None and f_node.get("t") == "shared" and f_node.text:
            group = f_node.get("si")
            if group is not None:
                anchor = format_cell(self.sheet, row, column)
                self._shared[group] = (anchor, is_volatile(f_node.text),
                                       is_external(f_node.text, self.external_names))
                self.shared_groups[anchor] = (f_node.text, f_node.get("ref"))

    def _formula(self, value: Value, f_node) -> None:
        value.formula = True
        text = f_node.text
        kind = f_node.get("t")
        if text and self.redaction is not None and self.redaction.holds(text):
            self._formula_parts(value, f_node, text, kind)
            value.formula_text = None
            value.formula_withheld = True
            return
        self._formula_parts(value, f_node, text, kind)

    def _formula_parts(self, value: Value, f_node, text: str | None, kind: str | None) -> None:
        if kind == "shared":
            group = f_node.get("si")
            if text:
                value.formula_text = decode_xstring(text)
                value.volatile = is_volatile(text)
                value.external = is_external(text, self.external_names)
                if group is not None:
                    self._shared[group] = (value.a, value.volatile, value.external)
                    self.shared_groups[value.a] = (text, f_node.get("ref"))
            elif group in self._shared:
                value.shared_from, value.volatile, value.external = self._shared[group]
            return
        if kind == "array":
            written = f_node.get("ref") or value.ref
            block = _area(written)
            # As every address is written ($ and case aside); a ref that is no range as
            # stored.
            value.array = block.ref if block is not None else written
            if block is not None:
                self.arrays.append(block)
        if text:
            # An ST_Xstring too: Excel writes "_x0041_" in a formula as "_x005F_x0041_".
            value.formula_text = decode_xstring(text)
            value.volatile = is_volatile(text)
            value.external = is_external(text, self.external_names)


def scan(reader: SheetReader) -> SheetScan:
    """One full pass over ``reader``'s sheet: the used range, merges, ordering and the
    issues met (see :class:`SheetScan`).  Cells with nothing in them are placed, never
    typed (``blanks`` off: :meth:`SheetReader.values`)."""
    found: list[SheetScan] = []
    for _ in scanning(reader, done=found.append):
        pass
    return found[0]


def scanning(reader: SheetReader, *, blanks: bool = False,
             done: "Callable[[SheetScan], Any]") -> Iterator[Value]:
    """``reader.values(blanks=blanks)``, each value passed on as it comes, while the
    :class:`SheetScan` of :func:`scan` is made of them: ``done`` is called with it when
    the pass ends (never when it is stopped early).  A pass that gives every cell (a
    :meth:`Workbook.cells <xlsx_agent.workbook.Workbook.cells>`) scans the sheet too."""
    result = SheetScan(reader.sheet)
    reader.marks = []
    top = left = MAX_ROW + MAX_COLUMN
    bottom = right = 0
    for value in reader.values(blanks=blanks):
        if value.t == "blank" and not value.formula:
            if value.issue is not None:
                result._count(value.issue, value)
            yield value
            continue
        result.cells += 1
        row, column = value.row, value.column
        if row < top:
            top = row
        if row > bottom:
            bottom = row
        if column < left:
            left = column
        if column > right:
            right = column
        if value.issue is not None:
            result._count(value.issue, value)
        if value.at_limit:
            result._count("at_limit", value)
        if value.volatile:
            result._count("volatile", value)
        if value.external:
            result._count("external_formula", value)
        yield value
    if result.cells:
        result.used = Area(Cell(top, left), Cell(bottom, right))
    result.dimension = reader.dimension
    result.merges = list(reader.merges)
    result.validations = list(reader.validations)
    result.arrays = list(reader.arrays)
    result.hyperlinks = list(reader.hyperlinks)
    result.hidden_rows = len(reader.hidden_rows)
    result.ordered = reader.ordered
    if reader.ordered:
        result.index = RowIndex.build(reader.source, reader)
    reader.marks = None
    done(result)


def merges_only(reader: SheetReader) -> "list[Area] | None":
    """The sheet's merged areas, in the file's order, as :func:`scan` finds them, without
    a pass over the cells: ``mergeCells`` comes after ``sheetData`` (ECMA-376 Part 1,
    §18.3.1.55), so the part is parsed as its header and what follows the cells.  ``None``
    when the part has no :class:`_Plan` or a ``mergeCell`` among its rows (a pass is
    needed)."""
    plan = reader.plan() if reader.source else None
    if plan is None or reader.facts().merges:
        return None
    out: list[Area] = []
    for _, node in iterparse_xml(_Resumed(plan.header, reader.source, plan.end),
                                 tag=tuple(_MERGE)):
        area = _area(node.get("ref"))
        if area is not None:
            out.append(area)
    return out


@dataclass(frozen=True)
class SheetLayout:
    """What a sheet holds besides its values, as a full :func:`scan` finds it: its merged
    areas, the blocks its array formulas fill (in reading order) and its data-validation
    rules (in the file's order)."""

    merges: tuple[Area, ...] = ()
    arrays: tuple[Area, ...] = ()
    validations: tuple[Validation, ...] = ()

    @classmethod
    def of_scan(cls, found: SheetScan) -> "SheetLayout":
        return cls(tuple(found.merges), tuple(found.arrays), tuple(found.validations))


_LAYOUT_TAGS = tuple(_F | _MERGE | _VALIDATION | _X14_VALIDATION)


def layout(root) -> "SheetLayout | None":
    """A sheet's :class:`SheetLayout` from its parsed tree (an edit's, which the cell
    writer holds): only the formulas, merges and rules are visited, never a value typed.
    ``None`` when an array formula's cell has no reference to place it by, which only the
    full pass places (as a cell without ``r`` is the one after the cell before it)."""
    merges: list[Area] = []
    arrays: list[Area] = []
    validations: list[Validation] = []
    for node in root.iter(*_LAYOUT_TAGS):
        tag = node.tag
        if tag in _F:
            cell = node.getparent()
            row = cell.getparent() if cell is not None else None
            if cell is None or cell.tag not in _C or row is None or row.tag not in _ROW \
                    or node.get("t") != "array" or node.getnext() is not None \
                    and any(sibling.tag in _F for sibling in node.itersiblings()):
                continue    # not a cell's formula, not an array, or not the one it reads
            written = cell.get("r")
            match = _REF.fullmatch(written) if written else None
            if match is None:
                return None
            column, number = _column(match.group(1)), int(match.group(2))
            if not (1 <= number <= MAX_ROW and 1 <= column <= MAX_COLUMN):
                continue    # off the grid: the full pass skips the cell
            block = _area(node.get("ref") or f"{_letters(column)}{number}")
            if block is not None:
                arrays.append(block)
        elif tag in _MERGE:
            area = _area(node.get("ref"))
            if area is not None:
                merges.append(area)
        elif tag in _VALIDATION:
            validations.append(_standard_validation(node))
        else:
            validations.append(_x14_validation(node))
    return SheetLayout(tuple(merges), tuple(arrays), tuple(validations))


def read_area(reader: SheetReader, area: Area, ordered: bool,
              index: RowIndex | None = None,
              columns: "set[int] | None" = None) -> Iterator[list[Value]]:
    """The rows of ``area``, each a list of values from its first column to its last,
    blanks filled in.  With ``ordered`` (the sheet's rows and cells in order, as a scan
    found) the pass starts at ``index``'s last checkpoint before the area, if any, and
    stops after the area's last row; otherwise the area's cells are gathered first.

    ``columns`` (sheet columns), for a pass that looks at some of the area's columns only:
    the cells of the others are not typed, and come as blanks.  The set is read as each
    row is reached, so the pass may change it between rows."""
    top, left = area.first.row, area.first.column
    bottom, right = area.last.row, area.last.column
    width = right - left + 1
    pending: dict[int, dict[int, Value]] = {}
    if not ordered:
        for value in reader.values():
            if top <= value.row <= bottom and left <= value.column <= right \
                    and (columns is None or value.column in columns):
                pending.setdefault(value.row, {})[value.column] = value
        for row in range(top, bottom + 1):
            yield _row(reader, row, left, width, pending.pop(row, None))
        return
    next_row = top
    current: dict[int, Value] = {}
    current_row = None
    checkpoint = index.before(top) if index is not None else None
    start = (index, checkpoint) if index is not None and checkpoint is not None else None
    for row, column, row_hidden, node in reader._walk(start):
        if row < top:
            continue
        if row != current_row:
            if current_row is not None:
                yield _row(reader, current_row, left, width, current)
                next_row = current_row + 1
            if row > bottom:
                current_row = None
                break
            while next_row < row:
                yield _row(reader, next_row, left, width, None)
                next_row += 1
            current_row, current = row, {}
        if left <= column <= right:
            if columns is None or column in columns:
                current[column] = reader.value(row, column, row_hidden, node)
            else:
                reader.note_shared(row, column, node)
    if current_row is not None:
        yield _row(reader, current_row, left, width, current)
        next_row = current_row + 1
    while next_row <= bottom:
        yield _row(reader, next_row, left, width, None)
        next_row += 1


def _row(reader: SheetReader, row: int, left: int, width: int,
         cells: dict[int, Value] | None) -> list[Value]:
    out = []
    row_hidden = reader.hidden or row in reader.hidden_rows
    for column in range(left, left + width):
        value = cells.get(column) if cells else None
        if value is None:
            value = Value(reader.sheet, row, column,
                          hidden=row_hidden or reader.column_hidden(column))
        out.append(value)
    return out


def _standard_validation(node) -> Validation:
    """A ``dataValidation`` of the worksheet's ``dataValidations``."""
    formula = next((child.text for child in node if child.tag in _FORMULA1), None)
    second = next((child.text for child in node if child.tag in _FORMULA2), None)
    return _validation(node, node.get("sqref"), formula, second, "standard")


def _x14_validation(node) -> Validation:
    """An ``x14:dataValidation`` of the worksheet's ``extLst``."""
    formula = second = sqref = None
    for child in node:
        if child.tag == _X14_FORMULA1:
            formula = next((f.text for f in child if f.tag == _XM_F), None)
        elif child.tag == _X14_FORMULA2:
            second = next((f.text for f in child if f.tag == _XM_F), None)
        elif child.tag == _XM_SQREF:
            sqref = child.text
    return _validation(node, sqref, formula, second, "x14")


def _validation(node, sqref: str | None, formula: str | None, second: str | None,
                origin: str) -> Validation:
    areas = tuple(area for area in map(_area, (sqref or "").split()) if area is not None)
    operator = node.get("operator")
    return Validation((node.get("type") or "none").strip(), areas, formula, origin=origin,
                      operator=operator.strip() if operator else None,
                      allow_blank=_true(node.get("allowBlank")), formula2=second,
                      sqref=" ".join((sqref or "").split()))


#: What a conditional format's rule is compared without: its place among the sheet's rules
#: (``priority``, renumbered as rules come and go), the format it applies (``dxfId``, a
#: style, and an ``x14`` rule's own ``dxf``) and the ids linking a rule to its ``x14``
#: extension (``id``, ``extLst``).
_CF_LEFT_OUT = ("priority", "dxfId", "id")
_CF_CHILDREN_LEFT_OUT = ("extLst", "dxf")


def _conditional(node, x14: bool) -> "ConditionalFormat":
    sqref = node.get("sqref")
    rules = []
    for child in node:
        if not isinstance(child.tag, str):
            continue
        local = child.tag.rpartition("}")[2]
        if local == "sqref":
            sqref = child.text
        elif local == "cfRule":
            rules.append((_int(child.get("priority")), _canonical(child)))
    rules.sort(key=lambda rule: (rule[0] is None, rule[0] or 0))
    return ConditionalFormat(" ".join((sqref or "").split()), "x14" if x14 else "standard",
                             tuple(rule for _, rule in rules))


def _canonical(node) -> tuple:
    """An element as compared: ``(local name, attributes, text, children)``, the attributes
    sorted (less :data:`_CF_LEFT_OUT`), the text stripped; a formula's (``formula``,
    ``xm:f``) as written, for the comparer to normalise."""
    attrs = tuple(sorted((key.rpartition("}")[2], value) for key, value in node.attrib.items()
                         if key.rpartition("}")[2] not in _CF_LEFT_OUT))
    children = tuple(_canonical(child) for child in node if isinstance(child.tag, str)
                     and child.tag.rpartition("}")[2] not in _CF_CHILDREN_LEFT_OUT)
    return (node.tag.rpartition("}")[2], attrs, (node.text or "").strip(), children)


def _merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for first, last in sorted(spans):
        if out and first <= out[-1][1] + 1:
            out[-1] = (out[-1][0], max(out[-1][1], last))
        else:
            out.append((first, last))
    return out


def _area(ref: str | None) -> Area | None:
    if not ref:
        return None
    first, _, last = ref.partition(":")
    a, b = _REF.fullmatch(first), _REF.fullmatch(last or first)
    if a is None or b is None:
        return None
    one = Cell(int(a.group(2)), _column(a.group(1)))
    two = Cell(int(b.group(2)), _column(b.group(1)))
    if not (1 <= min(one.row, two.row) and max(one.row, two.row) <= MAX_ROW
            and 1 <= min(one.column, two.column) and max(one.column, two.column) <= MAX_COLUMN):
        return None
    return Area(Cell(min(one.row, two.row), min(one.column, two.column)),
                Cell(max(one.row, two.row), max(one.column, two.column)))


def _true(value: str | None) -> bool:
    return value is not None and value.strip() in ("1", "true")


def _int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None
