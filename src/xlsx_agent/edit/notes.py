"""Notes on cells (production R57): ``wb.set_comment("Stock!B3", "Check count")``,
``wb.set_comments({...})``, ``wb.remove_comment("Stock!B3")``.

::

    wb.set_comment("Stock!B3", "Check count", author="Stock team")
    wb.set_comments({"Stock!B3": "Check count", "Stock!B9": "Recount in May"})   # one step
    wb.remove_comment("Stock!B3")                 # or a range: every note in it
    wb.comments("Stock")                          # what the sheet holds (read)

A note is what Excel calls a *note* now, and a *comment* in its object model and in
ECMA-376 (Part 1, §18.7): a cell's text in the sheet's comments part, shown by a hidden
VML shape (a *legacy drawing*, §18.3.1.48) that pops up over the sheet when the cell is
pointed at.  Each call is one undoable step (or part of the open batch), and refused
before anything changes, every problem listed.

**What Excel writes, measured** (Excel for Mac 16.106, 2026-11-09, its object model's
``add comment`` on made-up sheets, saved; the probes are in the round's notes):

* *Parts.*  A sheet's first note makes a comments part (``xl/commentsN.xml``, an
  ``Override`` of its content type) and a VML drawing (``xl/drawings/vmlDrawingN.vml``,
  the ``vml`` extension's ``Default`` type), both related from the sheet, and a
  ``legacyDrawing`` naming the drawing in the sheet.  Deleting a sheet's last note took
  all of it away again.
* *The comments part* lists its ``authors`` (Excel's own user name for a note it made)
  and its ``comment`` elements in reading order (row by row, then by column), each with
  ``ref``, ``authorId`` and ``shapeId="0"``, its text one run in the note's font: bold,
  Tahoma 10, black (``<rPr><b/><sz val="10"/><color rgb="FF000000"/><rFont
  val="Tahoma"/><family val="2"/></rPr>``).  Excel adds an ``xr:uid`` (a revision id)
  too, which is not written here: it is optional and opened without one (oracle).
* *The VML drawing* has a shape layout claiming a block of shape ids (``o:idmap``: 1 for
  the first sheet, 2 for the second; a block holds 1,024 ids, the first shape of block
  ``n`` is ``_x0000_s{1024n+1}``), the note shape type (``_x0000_t202``) and one shape per
  note, ids rising (a new shape takes the highest id plus one, even after a deletion),
  hidden, of Excel's note fill (``#ffffe1``), 104 by 69 points, its ``x:ClientData``
  naming the cell (``x:Row``, ``x:Column``, from 0) and its ``x:Anchor`` (left column and
  offset, top row and offset, right column and offset, bottom row and offset, in pixels;
  on a Mac a pixel is a point).
* *Where the note sits* (measured on default, wide, narrow, tall, hidden rows and
  columns, and the sheet's edges): its left edge 15 points right of the cell's right edge,
  its top 10 points above the cell's top (2 points below the sheet's top for row 1); a
  hidden row or column takes no room, so the note starts in the next one shown.  For the
  last column the note goes to the left of the cell (its right edge 15 points left of the
  cell's left edge), and for the last row it is moved up to end 8 points above the sheet's
  bottom.  Column widths in pixels are the stored width times the default font's widest
  digit, as Excel stores widths (:mod:`.layout`: 6 px for the default Aptos Narrow 12 and
  for Calibri 11 on a Mac, measured).
* *Text* is an ST_Xstring: what XML cannot carry written ``_xHHHH_`` (Excel wrote the
  control character U+0001 ``_x0001_``), a line feed as it is, ``&`` and ``<`` as XML
  escapes; read back as given (:func:`~xlsx_agent.read.strings.encode_xstring`, as cells).
* *A comment another program wrote without a shape* (a comments part with no VML for
  it) was dropped by Excel when it saved the sheet: ``set_comments`` gives every such
  comment on the sheets it writes a shape, as Excel's own, so it survives (``shaped``).

**Replacing** a note's text keeps its shape and its place; ``author`` given replaces the
author, ``None`` keeps it (a new note's author is then empty: a library has no user name).
A cell's note set to the text it has is ``changed`` false.  **Threaded comments**
(Excel's newer comments, [MS-XLSX]'s ``threadedComments`` part, read as the plain
comment Excel writes beside each) are not written: a cell holding one is refused
(``threaded_comment``) by both calls, as a note written there would break the thread.
Inserting, deleting and moving rows and columns move notes, shapes and anchors with their
cells (:mod:`.restructure`).

The host's write policy must allow the cell (a note is written on it, as a value is); a
protected sheet refuses it.
"""

from __future__ import annotations

import bisect
import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping, Sequence

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell, column_letters

from ..address import AddressError, RangeAddress, format_cell, format_range, parse_address
from ..read.strings import encode_xstring, needs_xstring
from . import EditError
from .sheets import R, X, _X, find_sheet, placed

#: The most cells listed in a refusal or a result.
MAX_LISTED = 50
#: The longest note text written (a cell's limit, as Excel's object model takes).
MAX_TEXT = 32_767
#: The most notes one call writes or removes.
MAX_NOTES = 10_000

REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"
#: [MS-XLSX]: a sheet's threaded comments.
REL_THREADED = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"

V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
XL = "urn:schemas-microsoft-com:office:excel"
_VML_NS = {"v": V, "o": O, "x": XL}
#: A note shape's size and place, in points (measured: see the module).
NOTE_WIDTH, NOTE_HEIGHT = 104, 69
NOTE_RIGHT_OF, NOTE_ABOVE, NOTE_TOP_MIN, NOTE_BOTTOM_GAP = 15, 10, 2, 8
#: The note shape type Excel writes, and its id.
SHAPE_TYPE = "_x0000_t202"
#: Shape ids a VML drawing's id block holds.
BLOCK = 1024
#: A note's run properties, as Excel wrote them (measured).
_RUN_FONT: tuple[tuple[str, dict[str, str]], ...] = (("b", {}), ("sz", {"val": "10"}), ("color", {"rgb": "FF000000"}),
             ("rFont", {"val": "Tahoma"}), ("family", {"val": "2"}))
#: A sheet's default row height when its ``sheetFormatPr`` gives none, in points (Excel's
#: for Calibri 11; the default font's own is not measured).
DEFAULT_ROW_HEIGHT = 15.0
_SHAPE_ID = re.compile(r"^_x0000_s(\d+)$")


@dataclass(frozen=True)
class NoteChange:
    """What :func:`set_comments` or :func:`remove_comments` did, by cell address:
    ``added`` new notes, ``replaced`` notes whose text or author changed, ``unchanged``
    notes already so, ``removed`` notes taken away; ``shaped`` comments another program
    wrote without a shape, given one; ``changed`` whether anything changed (``note`` says
    why not)."""

    action: str
    added: tuple[str, ...] = ()
    replaced: tuple[str, ...] = ()
    unchanged: tuple[str, ...] = ()
    removed: tuple[str, ...] = ()
    shaped: tuple[str, ...] = ()
    note: str | None = None

    @property
    def changed(self) -> bool:
        return bool(self.added or self.replaced or self.removed or self.shaped)

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"action": self.action, "changed": self.changed}
        for name in ("added", "replaced", "unchanged", "removed", "shaped"):
            cells = getattr(self, name)
            if cells:
                out[name] = list(cells[:limit])
                if len(cells) > limit:
                    out[f"{name}_cut"] = len(cells) - limit
        if self.note:
            out["note"] = self.note
        return out


# -- what a call names ---------------------------------------------------------------------------


@dataclass
class _Problems:
    found: list[tuple[str, str, str]] = field(default_factory=list)

    def add(self, code: str, address: str, detail: str) -> None:
        self.found.append((code, address, detail))

    def raise_any(self) -> None:
        if not self.found:
            return
        code, _, detail = self.found[0]
        same = [address for c, address, _ in self.found if c == code]
        more = len(self.found) - len(same)
        raise EditError(code, detail + (f"; {more} other problem(s) in problems" if more
                                        else ""),
                        addresses=same, problems=[(c, a) for c, a, _ in self.found])


def _parsed(text: Any, problems: _Problems, *, ranges: bool) -> RangeAddress | None:
    shown = text if isinstance(text, str) else repr(text)
    want = "a cell or range with its sheet, as Stock!B3 or Stock!A1:D20" if ranges \
        else "a cell with its sheet, as Stock!B3"
    try:
        parsed = parse_address(text)
    except (AddressError, ValueError) as error:
        problems.add("bad_address", shown, f"{error}: give {want}")
        return None
    if not isinstance(parsed, RangeAddress) or parsed.sheet is None:
        problems.add("bad_address", shown, f"give {want}")
        return None
    if not ranges and (parsed.span != "cells" or parsed.area.first != parsed.area.last):
        problems.add("bad_address", shown, f"a note is on one cell: give {want}")
        return None
    return parsed


def _sheet_of(workbook: Any, parsed: RangeAddress, shown: str, problems: _Problems) -> Any:
    ref = parsed.sheet
    try:
        sheet = find_sheet(workbook, f"#{ref.id}" if ref.id is not None else ref.name)
    except EditError as error:
        problems.add(error.code, shown, str(error))
        return None
    if sheet.type != "worksheet" or sheet.part is None:
        problems.add("not_a_worksheet", shown, "notes are on a worksheet's cells")
        return None
    return sheet


# -- a sheet's notes -----------------------------------------------------------------------------


class _Notes:
    """One sheet's comments part, VML drawing and threaded comments, read once; written
    back by :meth:`write`."""

    def __init__(self, workbook: Any, sheet: Any) -> None:
        self.workbook = workbook
        self.package = workbook.package
        self.sheet = sheet
        self.label = sheet.stored_name
        rels = self.package.relationships(sheet.part)
        self.comments_part = next((rel.target_part for rel in rels.values()
                                   if not rel.is_external and rel.type == REL_COMMENTS
                                   and rel.target_part
                                   and self.package.has_part(rel.target_part)), None)
        tree = self.package.tree(sheet.part)
        legacy = tree.find(_X + "legacyDrawing")
        self.vml_part = None
        if legacy is not None:
            rel = rels.get(legacy.get(f"{{{R}}}id") or "")
            if rel is not None and not rel.is_external and rel.target_part \
                    and self.package.has_part(rel.target_part):
                self.vml_part = rel.target_part
        self.threaded: set[Cell] = set()
        for rel in rels.values():
            if rel.type == REL_THREADED and rel.target_part \
                    and self.package.has_part(rel.target_part):
                root = self.package.tree(rel.target_part)
                for node in (root if root is not None else []):
                    cell = _cell(node.get("ref"))
                    if cell is not None:
                        self.threaded.add(cell)
        self._vml: Any = None
        self.vml_problem: str | None = None
        if self.vml_part is not None:
            try:
                self._vml = self.package.tree(self.vml_part)
            except (etree.XMLSyntaxError, ValueError):
                self.vml_problem = "the sheet's VML drawing cannot be read as XML"

    # -- reading --

    def root(self) -> Any:
        return None if self.comments_part is None else self.package.tree(self.comments_part)

    def entries(self) -> dict[Cell, Any]:
        """The comment elements by their cell (the first of each)."""
        root = self.root()
        out: dict[Cell, Any] = {}
        holder = None if root is None else root.find(_X + "commentList")
        for node in (holder if holder is not None else []):
            if node.tag != _X + "comment":
                continue
            cell = _cell(node.get("ref"))
            if cell is not None and cell not in out:
                out[cell] = node
        return out

    def authors(self) -> list[str]:
        root = self.root()
        holder = None if root is None else root.find(_X + "authors")
        return [node.text or "" for node in (holder if holder is not None else [])
                if node.tag == _X + "author"]

    def shapes(self) -> dict[Cell, Any]:
        """The note shapes of the VML drawing by the cell they name."""
        out: dict[Cell, Any] = {}
        if self._vml is None:
            return out
        for data in self._vml.iter(f"{{{XL}}}ClientData"):
            if data.get("ObjectType") != "Note":
                continue
            row, column = data.find(f"{{{XL}}}Row"), data.find(f"{{{XL}}}Column")
            try:
                cell = Cell(int((row.text or "").strip()) + 1,
                            int((column.text or "").strip()) + 1)
            except (AttributeError, ValueError):
                continue
            out.setdefault(cell, data.getparent())
        return out

    # -- writing --

    def ensure_parts(self, position: int) -> None:
        """Make the comments part and the VML drawing the sheet lacks (see the module)."""
        package = self.package
        if self.comments_part is None:
            root = etree.Element(_X + "comments", nsmap={None: X})
            etree.SubElement(root, _X + "authors")
            etree.SubElement(root, _X + "commentList")
            part = package.unused_part_name("xl/comments{n}.xml")
            package.add_part(part, _bytes(root), CT_COMMENTS, override=True)
            package.add_relationship(self.sheet.part, REL_COMMENTS, part)
            self.comments_part = part
        if self._vml is None:
            root = etree.Element("xml", nsmap=_VML_NS)
            layout = etree.SubElement(root, f"{{{O}}}shapelayout")
            layout.set(f"{{{V}}}ext", "edit")
            idmap = etree.SubElement(layout, f"{{{O}}}idmap")
            idmap.set(f"{{{V}}}ext", "edit")
            idmap.set("data", str(_free_block(package, position)))
            _shape_type(root)
            part = package.unused_part_name("xl/drawings/vmlDrawing{n}.vml")
            package.add_part(part, _bytes(root, declaration=False), CT_VML)
            rel_id = package.add_relationship(self.sheet.part, REL_VML, part)
            tree = package.tree(self.sheet.part)
            legacy = etree.Element(_X + "legacyDrawing")
            legacy.set(f"{{{R}}}id", rel_id)
            placed(tree, legacy)
            package.mark_dirty(self.sheet.part)
            self.vml_part = part
            self._vml = package.tree(part)
        elif self._vml.find(f"{{{V}}}shapetype[@id='{SHAPE_TYPE}']") is None:
            node = _shape_type(self._vml)
            layout = self._vml.find(f"{{{O}}}shapelayout")
            if layout is not None:
                layout.addnext(node)
            else:
                self._vml.insert(0, node)

    def author_id(self, author: str) -> int:
        root = self.root()
        holder = root.find(_X + "authors")
        if holder is None:
            holder = etree.Element(_X + "authors")
            root.insert(0, holder)
        names = [node.text or "" for node in holder if node.tag == _X + "author"]
        if author in names:
            return names.index(author)
        etree.SubElement(holder, _X + "author").text = author or None
        return len(names)

    def put(self, cell: Cell, text: str, author: str | None, node: Any) -> None:
        """Write the note of ``cell``: a new comment element in reading order (``node``
        ``None``), or the text (and author, given) of ``node``."""
        root = self.root()
        if node is None:
            holder = root.find(_X + "commentList")
            if holder is None:
                holder = etree.SubElement(root, _X + "commentList")
            node = etree.Element(_X + "comment")
            node.set("ref", f"{column_letters(cell.column)}{cell.row}")
            node.set("authorId", str(self.author_id(author or "")))
            node.set("shapeId", "0")
            after = None
            for existing in holder:
                other = _cell(existing.get("ref")) if existing.tag == _X + "comment" else None
                if other is not None and (other.row, other.column) > (cell.row, cell.column):
                    after = existing
                    break
            if after is None:
                holder.append(node)
            else:
                after.addprevious(node)
        elif author is not None:
            node.set("authorId", str(self.author_id(author)))
        for child in list(node):
            if child.tag == _X + "text":
                node.remove(child)
        body = etree.SubElement(node, _X + "text")
        run = etree.SubElement(body, _X + "r")
        properties = etree.SubElement(run, _X + "rPr")
        for tag, attrs in _RUN_FONT:
            etree.SubElement(properties, _X + tag, attrs)
        stored = encode_xstring(text) if needs_xstring(text) else text
        t = etree.SubElement(run, _X + "t")
        t.text = stored or None
        if stored != stored.strip(" \t\n"):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self.package.mark_dirty(self.comments_part)

    def drop(self, node: Any) -> None:
        node.getparent().remove(node)
        self.package.mark_dirty(self.comments_part)

    def add_shape(self, cell: Cell, grid: "_Grid") -> None:
        vml = self._vml
        numbers, zs = [], [0]
        for shape in vml.iter(f"{{{V}}}shape"):
            found = _SHAPE_ID.match(shape.get("id") or "")
            if found:
                numbers.append(int(found.group(1)))
            z = re.search(r"z-index:\s*(\d+)", shape.get("style") or "")
            if z:
                zs.append(int(z.group(1)))
        blocks = _blocks(vml)
        if not blocks:
            blocks = [_free_block(self.package, 1)]
            _set_blocks(vml, blocks)
        number = max(numbers + [blocks[0] * BLOCK]) + 1
        while number > max(blocks) * BLOCK + BLOCK - 1:
            blocks.append(_free_block(self.package, max(blocks) + 1))
            _set_blocks(vml, blocks)
            number = max(number, blocks[-1] * BLOCK + 1)
        left, top = grid.place(cell)
        anchor = grid.anchor(left, top)
        _shape(vml, number, max(zs) + 1, left, top, anchor, cell)
        self.package.mark_dirty(self.vml_part)

    def drop_shape(self, shape: Any) -> None:
        shape.getparent().remove(shape)
        self.package.mark_dirty(self.vml_part)

    def tidy(self) -> None:
        """Authors no comment names dropped (the rest renumbered); the parts, the
        relationships and the ``legacyDrawing`` removed when no note is left, as Excel
        does (a drawing holding other shapes stays)."""
        root = self.root()
        if root is None:
            return
        holder = root.find(_X + "commentList")
        entries = [n for n in (holder if holder is not None else [])
                   if n.tag == _X + "comment"]
        names = self.authors()
        used = sorted({int(n.get("authorId", "0")) for n in entries
                       if (n.get("authorId") or "0").isdigit()})
        if used != list(range(len(names))):
            keep = [i for i in used if i < len(names)]
            renumber = {old: new for new, old in enumerate(keep)}
            names_node = root.find(_X + "authors")
            for index, node in enumerate([n for n in names_node if n.tag == _X + "author"]):
                if index not in renumber:
                    names_node.remove(node)
            for node in entries:
                old = int(node.get("authorId", "0")) if (node.get("authorId") or "0").isdigit() \
                    else 0
                node.set("authorId", str(renumber.get(old, 0)))
            self.package.mark_dirty(self.comments_part)
        if entries:
            return
        package, part = self.package, self.sheet.part
        release = [rel.id for rel in package.relationships(part).values()
                   if rel.type == REL_COMMENTS and rel.target_part == self.comments_part]
        vml = self._vml
        others = vml is not None and any(
            shape.find(f"{{{XL}}}ClientData") is None
            or shape.find(f"{{{XL}}}ClientData").get("ObjectType") != "Note"
            for shape in vml.iter(f"{{{V}}}shape"))
        if vml is not None and not others:
            tree = package.tree(part)
            legacy = tree.find(_X + "legacyDrawing")
            if legacy is not None:
                release.append(legacy.get(f"{{{R}}}id") or "")
                tree.remove(legacy)
                package.mark_dirty(part)
        package.release(part, [r for r in release if r])
        if self.comments_part is not None and package.has_part(self.comments_part):
            package.remove_part(self.comments_part)
        self.comments_part = None


# -- the grid, in pixels --------------------------------------------------------------------------


class _Grid:
    """A sheet's columns and rows in pixels (a point each, on a Mac), hidden ones none."""

    def __init__(self, workbook: Any, sheet: Any) -> None:
        from .layout import Columns, default_face, digit_width

        writer = workbook.editor.cells.sheet(sheet.stored_name)
        columns = Columns(workbook, writer)
        digit = digit_width(default_face(workbook))
        default = _pixels(columns.default_width(), digit)
        widths = [default] * (MAX_COLUMN + 1)
        widths[0] = 0
        for low, high, attrs in columns.runs:
            hidden = (attrs.get("hidden") or "").strip().lower() in ("1", "true")
            try:
                width = 0 if hidden else _pixels(float(attrs["width"]), digit) \
                    if "width" in attrs else default
            except ValueError:
                width = default
            for number in range(max(low, 1), min(high, MAX_COLUMN) + 1):
                widths[number] = width
        self.lefts = [0] * (MAX_COLUMN + 2)
        for number in range(1, MAX_COLUMN + 1):
            self.lefts[number + 1] = self.lefts[number] + widths[number]
        self.widths = widths
        tree = workbook.package.tree(sheet.part)
        fmt = tree.find(_X + "sheetFormatPr")
        height = DEFAULT_ROW_HEIGHT
        zero = False
        if fmt is not None:
            try:
                height = float(fmt.get("defaultRowHeight", height))
            except ValueError:
                pass
            zero = (fmt.get("zeroHeight") or "").strip().lower() in ("1", "true")
        self.default_height = 0 if zero else int(round(height))
        custom: dict[int, int] = {}
        data = tree.find(_X + "sheetData")
        number = 0
        for row in (data.iterchildren(_X + "row") if data is not None else []):
            given = row.get("r")
            number = int(given) if given and given.isdigit() else number + 1
            hidden = (row.get("hidden") or "").strip().lower() in ("1", "true")
            ht = row.get("ht")
            if hidden:
                custom[number] = 0
            elif ht:
                try:
                    custom[number] = int(round(float(ht)))
                except ValueError:
                    pass
        self.rows = sorted(r for r, h in custom.items() if h != self.default_height)
        self.heights = custom
        self.extra = [0]
        for r in self.rows:
            self.extra.append(self.extra[-1] + custom[r] - self.default_height)

    def height(self, row: int) -> int:
        return self.heights.get(row, self.default_height)

    def top(self, row: int) -> int:
        """The top of ``row`` (from 1)."""
        index = bisect.bisect_left(self.rows, row)
        return (row - 1) * self.default_height + self.extra[index]

    def total_height(self) -> int:
        return self.top(MAX_ROW + 1)

    def place(self, cell: Cell) -> tuple[int, int]:
        """The note's left and top edges for ``cell`` (see the module)."""
        left = self.lefts[cell.column + 1] + NOTE_RIGHT_OF
        if left + NOTE_WIDTH > self.lefts[MAX_COLUMN + 1]:
            left = self.lefts[cell.column] - NOTE_RIGHT_OF - NOTE_WIDTH
        top = max(self.top(cell.row) - NOTE_ABOVE, NOTE_TOP_MIN)
        bottom = self.total_height() - NOTE_BOTTOM_GAP
        if top + NOTE_HEIGHT > bottom:
            top = bottom - NOTE_HEIGHT
        return max(left, 0), max(top, 0)

    def column_at(self, x: int) -> tuple[int, int]:
        """The column shown at ``x`` (from 0) and the offset into it."""
        number = bisect.bisect_right(self.lefts, x, 1, MAX_COLUMN + 1) - 1
        number = min(max(number, 1), MAX_COLUMN)
        return number - 1, x - self.lefts[number]

    def row_at(self, y: int) -> tuple[int, int]:
        low, high = 1, MAX_ROW
        while low < high:
            middle = (low + high + 1) // 2
            if self.top(middle) <= y:
                low = middle
            else:
                high = middle - 1
        return low - 1, y - self.top(low)

    def anchor(self, left: int, top: int) -> tuple[int, ...]:
        c1, x1 = self.column_at(left)
        r1, y1 = self.row_at(top)
        c2, x2 = self.column_at(left + NOTE_WIDTH)
        r2, y2 = self.row_at(top + NOTE_HEIGHT)
        return c1, x1, r1, y1, c2, x2, r2, y2


def _pixels(width: float, digit: int) -> int:
    """A stored column width in pixels (ECMA-376 Part 1, 18.3.1.13)."""
    return int(((256 * width + int(128 / digit)) / 256) * digit)


# -- VML -------------------------------------------------------------------------------------------


def _shape_type(parent: Any) -> Any:
    node = etree.SubElement(parent, f"{{{V}}}shapetype")
    node.set("id", SHAPE_TYPE)
    node.set("coordsize", "21600,21600")
    node.set(f"{{{O}}}spt", "202")
    node.set("path", "m,l,21600r21600,l21600,xe")
    etree.SubElement(node, f"{{{V}}}stroke", joinstyle="miter")
    path = etree.SubElement(node, f"{{{V}}}path", gradientshapeok="t")
    path.set(f"{{{O}}}connecttype", "rect")
    return node


def _shape(parent: Any, number: int, z: int, left: int, top: int, anchor: tuple[int, ...],
           cell: Cell) -> Any:
    shape = etree.SubElement(parent, f"{{{V}}}shape")
    shape.set("id", f"_x0000_s{number}")
    shape.set("type", f"#{SHAPE_TYPE}")
    shape.set("style", f"position:absolute;margin-left:{left}pt;margin-top:{top}pt;"
              f"width:{NOTE_WIDTH}pt;height:{NOTE_HEIGHT}pt;z-index:{z};visibility:hidden")
    shape.set("fillcolor", "#ffffe1 [80]")
    shape.set("strokecolor", "none [81]")
    shape.set(f"{{{O}}}insetmode", "auto")
    etree.SubElement(shape, f"{{{V}}}fill", color2="#ffffe1 [80]")
    etree.SubElement(shape, f"{{{V}}}shadow", color="none [81]", obscured="t")
    path = etree.SubElement(shape, f"{{{V}}}path")
    path.set(f"{{{O}}}connecttype", "none")
    box = etree.SubElement(shape, f"{{{V}}}textbox", style="mso-direction-alt:auto")
    etree.SubElement(box, "div", style="text-align:left")
    data = etree.SubElement(shape, f"{{{XL}}}ClientData", ObjectType="Note")
    etree.SubElement(data, f"{{{XL}}}MoveWithCells")
    etree.SubElement(data, f"{{{XL}}}SizeWithCells")
    etree.SubElement(data, f"{{{XL}}}Anchor").text = ", ".join(str(n) for n in anchor)
    etree.SubElement(data, f"{{{XL}}}AutoFill").text = "False"
    etree.SubElement(data, f"{{{XL}}}Row").text = str(cell.row - 1)
    etree.SubElement(data, f"{{{XL}}}Column").text = str(cell.column - 1)
    return shape


def _blocks(vml: Any) -> list[int]:
    idmap = vml.find(f"{{{O}}}shapelayout/{{{O}}}idmap")
    if idmap is None:
        return []
    return [int(n) for n in re.findall(r"\d+", idmap.get("data") or "")]


def _set_blocks(vml: Any, blocks: list[int]) -> None:
    layout = vml.find(f"{{{O}}}shapelayout")
    if layout is None:
        layout = etree.Element(f"{{{O}}}shapelayout")
        layout.set(f"{{{V}}}ext", "edit")
        vml.insert(0, layout)
    idmap = layout.find(f"{{{O}}}idmap")
    if idmap is None:
        idmap = etree.SubElement(layout, f"{{{O}}}idmap")
        idmap.set(f"{{{V}}}ext", "edit")
    idmap.set("data", ",".join(str(b) for b in blocks))


def _free_block(package: Any, wanted: int) -> int:
    """``wanted`` when no VML drawing of the package claims it, else the lowest that none
    does (from 1)."""
    claimed: set[int] = set()
    for name in package.part_names:
        if not name.lower().endswith(".vml"):
            continue
        data = package.read(name) or b""
        for found in re.finditer(rb"idmap[^>]*?data=\"([^\"]*)\"", data):
            claimed.update(int(n) for n in re.findall(rb"\d+", found.group(1)))
    if wanted >= 1 and wanted not in claimed:
        return wanted
    number = 1
    while number in claimed:
        number += 1
    return number


def _bytes(root: Any, *, declaration: bool = True) -> bytes:
    body = etree.tostring(root, encoding="UTF-8")
    if declaration:
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body
    return body


def _cell(ref: str | None) -> Cell | None:
    try:
        area = Area.parse(ref or "", lenient=True)
    except ValueError:
        return None
    return area.first


# -- the calls -----------------------------------------------------------------------------------


def _guard(workbook: Any, sheet: Any, cells: Iterable[Cell], problems: _Problems) -> None:
    label = sheet.stored_name
    if sheet.protected:
        problems.add("protected", format_cell(label, *_first(cells)), "the sheet is "
                     "protected: unprotect it to change its notes")
        return
    allowed = workbook._allowed()
    if allowed is not None:
        for cell in cells:
            if not allowed.allows(sheet.part, cell):
                problems.add("not_writable", format_cell(label, cell.row, cell.column),
                             "the host's write policy does not let this cell be written")


def _first(cells: Iterable[Cell]) -> tuple[int, int]:
    for cell in cells:
        return cell.row, cell.column
    return 1, 1


def _position(workbook: Any, sheet: Any) -> int:
    for index, other in enumerate(workbook._every_sheet, 1):
        if other.part == sheet.part:
            return index
    return 1


def set_comments(workbook: Any, notes: "Mapping[str, Any] | Sequence[tuple[str, Any]]", *,
                 author: str | None = None) -> NoteChange:
    """Give each cell named its note (see the module): ``notes`` maps a cell address
    (``Stock!B3``) to its text; ``author`` names the author of every note written
    (``None``: a note's own author is kept, a new one's is empty)."""
    if workbook.strict:
        raise EditError("strict")
    items = list(notes.items()) if isinstance(notes, Mapping) else list(notes)
    problems = _Problems()
    if not items:
        return NoteChange("set", note="no notes given")
    if len(items) > MAX_NOTES:
        raise EditError("too_many_cells", f"{len(items):,} notes; at most {MAX_NOTES:,} in "
                        "one call")
    if author is not None and (not isinstance(author, str) or len(author) > 255
                               or "\n" in author or "\r" in author):
        problems.add("bad_value", "author", "an author is text of at most 255 characters on "
                     "one line, or None")
    planned: dict[str, tuple[Any, dict[Cell, str]]] = {}
    seen: set[tuple[str, Cell]] = set()
    for item in items:
        if not (isinstance(item, (tuple, list)) and len(item) == 2):
            problems.add("bad_value", repr(item)[:40], "give {cell: text} or (cell, text) pairs")
            continue
        address, text = item
        shown = address if isinstance(address, str) else repr(address)
        parsed = _parsed(address, problems, ranges=False)
        if parsed is None:
            continue
        sheet = _sheet_of(workbook, parsed, shown, problems)
        if sheet is None:
            continue
        cell = parsed.area.first
        label = format_cell(sheet.stored_name, cell.row, cell.column)
        if (sheet.part, cell) in seen:
            problems.add("duplicate_cell", label, "a cell's note is given twice")
            continue
        seen.add((sheet.part, cell))
        if not isinstance(text, str):
            problems.add("bad_value", label, f"a note's text is text, not "
                         f"{type(text).__name__}: remove_comment takes a note away")
            continue
        if len(text) > MAX_TEXT:
            problems.add("text_too_long", label, f"{len(text):,} characters; a note holds at "
                         f"most {MAX_TEXT:,}")
            continue
        planned.setdefault(sheet.part, (sheet, {}))[1][cell] = text
    jobs: list[tuple[_Notes, dict[Cell, str]]] = []
    for sheet, cells in planned.values():
        _guard(workbook, sheet, cells, problems)
        notes_of = _Notes(workbook, sheet)
        if notes_of.vml_problem:
            problems.add("not_measured", format_range(sheet.stored_name, None),
                         notes_of.vml_problem)
            continue
        for cell in cells:
            if cell in notes_of.threaded:
                problems.add("threaded_comment", format_cell(sheet.stored_name, cell.row,
                                                             cell.column),
                             "the cell holds a threaded comment, which is not written here: "
                             "reply to it or delete it in Excel")
        jobs.append((notes_of, cells))
    problems.raise_any()
    added: list[str] = []
    replaced: list[str] = []
    unchanged: list[str] = []
    shaped: list[str] = []
    editor = workbook.editor
    with editor.batch():
        for notes_of, cells in jobs:
            sheet = notes_of.sheet
            label = sheet.stored_name
            entries = notes_of.entries()
            authors = notes_of.authors()
            todo: dict[Cell, str] = {}
            for cell, text in cells.items():
                address = format_cell(label, cell.row, cell.column)
                node = entries.get(cell)
                if node is not None and _same(node, text, author, authors):
                    unchanged.append(address)
                    continue
                todo[cell] = text
                (replaced if node is not None else added).append(address)
            if not todo:
                continue
            notes_of.ensure_parts(_position(workbook, sheet))
            entries = notes_of.entries()
            for cell, text in todo.items():
                notes_of.put(cell, text, author, entries.get(cell))
            shapes = notes_of.shapes()
            grid: _Grid | None = None
            for cell in sorted(notes_of.entries(), key=lambda c: (c.row, c.column)):
                if cell in shapes:
                    continue
                grid = grid or _Grid(workbook, sheet)
                notes_of.add_shape(cell, grid)
                if cell not in todo:
                    shaped.append(format_cell(label, cell.row, cell.column))
            notes_of.tidy()
            editor.touched(sheet, [])
        workbook._forget()
    note = None if added or replaced or shaped else "every note given is already so"
    return NoteChange("set", added=tuple(added), replaced=tuple(replaced),
                      unchanged=tuple(unchanged), shaped=tuple(shaped), note=note)


def _same(node: Any, text: str, author: str | None, authors: list[str]) -> bool:
    from ..read.strings import item_text

    body = node.find(_X + "text")
    if body is None or item_text(body) != text:
        return False
    if author is None:
        return True
    index = node.get("authorId") or ""
    return index.isdigit() and int(index) < len(authors) and authors[int(index)] == author


def remove_comments(workbook: Any, targets: "str | Sequence[str]") -> NoteChange:
    """Take away the note of each cell named, or every note in each range named
    (``Stock!A1:D20``, ``Stock!B:B``, ``Stock!``); none there: ``changed`` false."""
    if workbook.strict:
        raise EditError("strict")
    texts = [targets] if isinstance(targets, str) else list(targets)
    problems = _Problems()
    if not texts:
        return NoteChange("remove", note="no cells given")
    areas: dict[str, tuple[Any, list[Area]]] = {}
    for text in texts:
        shown = text if isinstance(text, str) else repr(text)
        parsed = _parsed(text, problems, ranges=True)
        if parsed is None:
            continue
        sheet = _sheet_of(workbook, parsed, shown, problems)
        if sheet is None:
            continue
        areas.setdefault(sheet.part, (sheet, []))[1].append(parsed.area)
    jobs: list[tuple[_Notes, list[tuple[Cell, Any]]]] = []
    for sheet, found_areas in areas.values():
        notes_of = _Notes(workbook, sheet)
        entries = notes_of.entries()
        hit = [(cell, node) for cell, node in entries.items()
               if any(area.contains(cell) for area in found_areas)]
        threaded = sorted(cell for cell in notes_of.threaded
                          if any(area.contains(cell) for area in found_areas))
        for cell in threaded:
            problems.add("threaded_comment", format_cell(sheet.stored_name, cell.row,
                                                         cell.column),
                         "the cell holds a threaded comment, which is not removed here: "
                         "delete it in Excel, or name the cells around it")
        if notes_of.vml_problem and hit:
            problems.add("not_measured", format_range(sheet.stored_name, None),
                         notes_of.vml_problem)
            continue
        if hit:
            _guard(workbook, sheet, [cell for cell, _ in hit], problems)
            jobs.append((notes_of, hit))
    problems.raise_any()
    if not jobs:
        return NoteChange("remove", note="no note there")
    removed: list[str] = []
    editor = workbook.editor
    with editor.batch():
        for notes_of, hit in jobs:
            shapes = notes_of.shapes()
            for cell, node in sorted(hit, key=lambda pair: (pair[0].row, pair[0].column)):
                notes_of.drop(node)
                if cell in shapes:
                    notes_of.drop_shape(shapes[cell])
                removed.append(format_cell(notes_of.sheet.stored_name, cell.row, cell.column))
            notes_of.tidy()
            editor.touched(notes_of.sheet, [])
        workbook._forget()
    return NoteChange("remove", removed=tuple(removed))


__all__ = ["MAX_NOTES", "MAX_TEXT", "NOTE_HEIGHT", "NOTE_WIDTH", "NoteChange",
           "remove_comments", "set_comments"]
