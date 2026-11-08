"""The structure check: what an application would repair, found by code and address.

:meth:`Workbook.validate <xlsx_agent.Workbook.validate>` reads the package as it is now --
edited or not -- and lists what is out of order in the structure the format's parts must
keep with each other (ECMA-376 Part 1, §18; the package rules of Part 2).  It is the check
an edit is held to (every edit report carries it, :mod:`xlsx_agent.edit`) and the one a
save repeats on the bytes it hands over.  Each issue is a
:class:`~xlsx_agent.read.types.Problem`: a code, the first address it was met at and how
many times; never a cell's text, a formula or a name's formula (ROADMAP.md, "Design rules",
7).  A worksheet is streamed, never parsed whole -- or read from its tree when an edit's
cell writer has parsed it already.

**A step pays for what it touched** (round 44, lane P).  The check each edit step closes
with reads, of a sheet whose cells alone were written, the cells of the rows written (and
of its tables' header rows): every row is still checked as a row (its number and order),
the ``dimension`` is held to the cells read and the calculation chain to their formulas.
On those cells it finds what the whole check finds (``tests/test_structure_steps.py``
compares them on random edits and on damage put where a step wrote); an issue in a cell
of a row the step did not write is :meth:`Workbook.validate`'s, which stays the whole
check.  A sheet changed otherwise (formatting, layout, rows or columns inserted, deleted or
moved) is checked whole.

The codes (:data:`CODES`):

* ``row_order`` -- a row numbered at or before the row before it (§18.3.1.73: rows in
  ascending order); ``cell_order`` -- a cell at or before the cell before it in its row;
  ``cell_row`` -- a cell whose reference names another row than its row's;
  ``bad_reference`` -- a row or cell reference that is not one, or off the grid.
* ``shared_string`` -- a ``t="s"`` cell whose index is not one of the shared-string
  table's items (or there is no table); ``shared_string_count`` -- the table's
  ``uniqueCount`` is not its number of items, or its ``count`` not the number of cells
  that use it (§18.4.9).
* ``style`` -- a cell whose ``s`` is not an index into ``cellXfs``.
* ``dimension`` -- the sheet's ``dimension`` is not a range, or does not hold every cell
  with a value or a formula (the reader never trusts it; an edit keeps it right).
* ``table_ref`` -- a table whose ``ref`` is not a range; ``table_overlap`` -- two tables
  on one sheet that overlap; ``table_columns`` -- a table with another number of columns
  than its ``ref`` is wide; ``table_name`` -- two tables with one name (``displayName``,
  case aside) or one ``id``; ``table_header`` -- a header cell holding text other than its
  column's name (§18.5.1.2: the header row shows the column names).
* ``name_duplicate`` -- two defined names alike (case aside) in one scope;
  ``name_scope`` -- a name whose ``localSheetId`` is no sheet.  The address is the name.
* ``missing_part`` -- an internal relationship whose target part is not in the package
  (the address is the relationships part and the id: ``xl/_rels/workbook.xml.rels#rId6``);
  ``no_content_type`` -- a part with neither a ``Default`` nor an ``Override``;
  ``stray_override`` -- an ``Override`` for a part the package does not hold.
* ``calc_chain`` -- the calculation chain names a cell that holds no formula.

A strict package (``purl.oclc.org`` namespaces) is checked alike.
"""

from __future__ import annotations

import re
from typing import Any, Iterable, Iterator, Mapping

from ooxml_edit.opc import rels_path_for
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell, column_index

from ..read.strings import item_text
from ..address import format_cell, format_range, format_sheet
from ..read.types import Problem
from ..xml import iterparse_xml

#: Every code :func:`check_structure` gives, in the order the module lists them.
CODES = ("row_order", "cell_order", "cell_row", "bad_reference", "shared_string",
         "shared_string_count", "style", "dimension", "table_ref", "table_overlap",
         "table_columns", "table_name", "table_header", "name_duplicate", "name_scope",
         "missing_part", "no_content_type", "stray_override", "calc_chain")

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")
_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
_SI = frozenset(f"{{{ns}}}si" for ns in NS_MAIN)
_CELL = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_RANGE = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$")


class _Found:
    """Issues counted by code and place: the first address of each, and how many."""

    def __init__(self) -> None:
        self._seen: dict[tuple[str, str], list] = {}

    def add(self, code: str, address: str | None, group: str = "") -> None:
        key = (code, group)
        entry = self._seen.get(key)
        if entry is None:
            self._seen[key] = [address, 1]
        else:
            entry[1] += 1

    def problems(self) -> list[Problem]:
        order = {code: index for index, code in enumerate(CODES)}
        items = sorted(self._seen.items(), key=lambda item: order[item[0][0]])
        return [Problem(code, address, count) for (code, _), (address, count) in items]


def check_structure(workbook: Any, *, sheets: "Iterable[str | None] | None" = None,
                    rows: "Mapping[str, Iterable[int]] | None" = None) -> list[Problem]:
    """The structure issues of ``workbook`` (an :class:`xlsx_agent.Workbook`), each by
    code with its first address and count; empty when there are none.  ``sheets`` limits
    the worksheet checks (cells, dimension, tables' headers) to those sheet parts; the
    workbook-wide checks always run.

    ``rows`` (sheet part: row numbers) narrows a sheet's cell checks further, to the cells
    of those rows and of its tables' header rows: what an edit step wrote (see the
    module).  Every row of such a sheet is still checked as a row (its number and
    order), its ``dimension`` is held to the cells checked, and the calculation chain to the
    formulas of those rows: on those cells it finds what the whole check finds
    (``tests/test_structure_steps.py`` compares them)."""
    found = _Found()
    package = workbook.package
    strings = _StringItems(workbook)
    styles = _style_count(workbook)
    every = [s for s in workbook._every_sheet if s.part is not None
             and s.type in ("worksheet", "macrosheet")]
    chosen = every if sheets is None else [s for s in every if s.part in set(sheets)]
    tables = _tables(workbook, found)
    headers: dict[str, dict[tuple[int, int], str]] = {}
    for table in tables:
        if table.area is not None and table.header_rows:
            row = table.area.first.row
            for offset, name in enumerate(table.columns):
                headers.setdefault(table.sheet, {})[(row, table.area.first.column + offset)] = name
    formulas: dict[str, set[tuple[int, int]]] = {}
    checked: dict[str, "set[int] | None"] = {}
    uses = 0
    for sheet in chosen:
        label = _label(sheet)
        own = headers.get(sheet.name, {})
        only = None
        if rows is not None and sheet.part in rows:
            only = set(rows[sheet.part]) | {row for row, _ in own}
        checked[sheet.part] = only
        cells, used = _sheet(_rows_of(package, sheet.part, only is not None), label, found,
                             strings, styles, own, only)
        formulas[sheet.part] = cells
        uses += used
    if strings.part is not None:
        _string_counts(workbook, found, uses if sheets is None else None)
    _names(workbook, found)
    _package(package, found)
    _calc_chain(workbook, found, formulas, every, checked)
    return found.problems()


# -- worksheets --------------------------------------------------------------------------------


def _label(sheet: Any) -> str:
    return format_sheet(sheet.stored_name, sheet.id)


_ROW = frozenset(f"{{{ns}}}row" for ns in NS_MAIN)
_C = frozenset(f"{{{ns}}}c" for ns in NS_MAIN)
_V = frozenset(f"{{{ns}}}v" for ns in NS_MAIN)
_F = frozenset(f"{{{ns}}}f" for ns in NS_MAIN)
_IS = frozenset(f"{{{ns}}}is" for ns in NS_MAIN)
_DIMENSION = frozenset(f"{{{ns}}}dimension" for ns in NS_MAIN)
_SHEET_DATA = frozenset(f"{{{ns}}}sheetData" for ns in NS_MAIN)


def _rows_of(package: Any, part: str, live: bool) -> "Iterator[tuple[str, Any]]":
    """A worksheet as the check reads it: ``("dimension", ref)`` for each ``dimension``,
    then each child of ``sheetData`` (a row, or a cell outside one) as ``("data", node)``,
    in the file's order.  From the part's tree when the package holds one (the cell writer
    parsed it) or ``live`` asks for it; streamed from its bytes otherwise, a row at a time
    (ROADMAP.md, "Design rules", 6)."""
    trees = getattr(package, "_trees", None)    # ooxml-edit's parsed parts: read, not parsed
    root = package.tree(part) if live else (trees.get(part) if isinstance(trees, dict)
                                              else None)
    if root is not None:
        yield from _tree_rows(root)
        return
    data = package.read(part) or b""
    if not data:
        return
    tags = tuple(_ROW | _DIMENSION | _SHEET_DATA)
    for _, node in iterparse_xml(data, tag=tags):
        tag = node.tag
        if tag in _DIMENSION:
            yield "dimension", node.get("ref")
            continue
        parent = node.getparent()
        if tag in _SHEET_DATA:
            # what is left: the cells after the last row, outside any
            for child in node:
                if child.tag in _C:
                    yield "data", child
            continue
        if parent is None or parent.tag not in _SHEET_DATA:
            continue
        # The cells outside a row before this one, then the row; then everything before it
        # dropped: one row is held at a time.
        for child in list(parent):
            if child is node:
                break
            if child.tag in _C:
                yield "data", child
        yield "data", node
        node.clear()
        while node.getprevious() is not None:
            del parent[0]
        del parent[0]


def _tree_rows(root: Any) -> "Iterator[tuple[str, Any]]":
    """:func:`_rows_of` from a parsed sheet."""
    for child in root:
        if child.tag in _DIMENSION:
            yield "dimension", child.get("ref")
            continue
        if child.tag not in _SHEET_DATA:
            continue
        for node in child:
            if node.tag in _ROW or node.tag in _C:
                yield "data", node


def _sheet(source: "Iterable[tuple[str, Any]]", label: str, found: _Found,
           strings: "_StringItems", styles: int | None, headers: dict[tuple[int, int], str],
           only: "set[int] | None" = None) -> tuple[set[tuple[int, int]], int]:
    """Check one worksheet: its rows' and cells' order and references, each cell's shared
    string and style, the dimension against the cells, the tables' header cells.  ``only``:
    the cells of those rows alone (every row is still checked as a row).  Returns the
    positions of the formula cells and how many cells use a shared string (of the cells
    checked).

    A sheet may hold a million rows, so the common cell -- a reference of its row, a
    style that is a number -- is checked in line; anything else goes the long way
    (:func:`_reference`, :func:`_number`), which gives the same answer."""
    formulas: set[tuple[int, int]] = set()
    uses = 0
    dimension: str | None = None
    top = left = MAX_ROW + MAX_COLUMN
    bottom = right = 0
    row_number = 0
    column = 0
    letters = _LETTERS
    cell_tags, row_tags, v_tags, f_tags, is_tags = _C, _ROW, _V, _F, _IS
    count = len(strings) if strings.part is not None else -1
    limit = styles if styles is not None else -1
    for kind, node in source:
        if kind == "dimension":
            dimension = node
            continue
        if node.tag in row_tags:
            written = node.get("r")
            if written is not None and written.isdigit() and written[0] != "0" \
                    and int(written) <= MAX_ROW:
                number: "int | None" = int(written)
            else:
                number = _number(written) if written is not None else row_number + 1
            if number is None or not 1 <= number <= MAX_ROW:
                found.add("bad_reference", f"{label}!{written}", label)
                number = row_number + 1
            elif number <= row_number:
                found.add("row_order", format_range(label, Area(Cell(number, 1),
                                                              Cell(number, MAX_COLUMN)),
                                                     "rows"), label)
            row_number = max(row_number, number)
            column = 0
            if only is not None and row_number not in only and number not in only:
                continue
            cells = node
        else:       # a cell outside any row: of the row before it
            if only is not None and row_number not in only:
                continue
            cells = (node,)
        digits = str(row_number)
        cut = -len(digits)
        for cell in cells:
            if cell.tag not in cell_tags:
                continue
            ref = cell.get("r")
            if ref is None:
                row, col = row_number, column + 1
            elif ref.endswith(digits) and ref[:cut] in letters:
                row, col = row_number, letters[ref[:cut]]
            else:
                row, col = _reference(ref, label, row_number, column, found)
            if col <= column:
                found.add("cell_order", format_cell(label, row, col), label)
            else:
                column = col
            raw = formula = inline = None
            for child in cell:
                tag = child.tag
                if tag in v_tags:
                    if raw is None:
                        raw = child
                elif tag in f_tags:
                    if formula is None:
                        formula = child
                elif tag in is_tags and inline is None:
                    inline = child
            if formula is not None:
                formulas.add((row, col))
            style = cell.get("s")
            if style is not None and limit >= 0 and not (style.isdigit()
                                                         and int(style) < limit):
                index = _number(style, zero=True)
                if index is None or index >= limit:
                    found.add("style", format_cell(label, row, col), label)
            kind = cell.get("t", "n")
            if kind == "s":
                uses += 1
                text = raw.text if raw is not None else None
                index = int(text) if text is not None and text.isdigit() else \
                    _number(text.strip() if text else None, zero=True)
                if count < 0 or index is None or index >= count:
                    found.add("shared_string", format_cell(label, row, col), label)
                elif headers and (row, col) in headers:
                    _header(found, label, row, col, strings[index], headers)
            elif headers and (row, col) in headers:
                if kind == "inlineStr" and inline is not None:
                    _header(found, label, row, col, item_text(inline), headers)
                elif kind == "str" and raw is not None:
                    _header(found, label, row, col, raw.text or "", headers)
            if formula is not None or raw is not None or kind == "inlineStr":
                if row < top:
                    top = row
                if row > bottom:
                    bottom = row
                if col < left:
                    left = col
                if col > right:
                    right = col
    if dimension is not None or bottom:
        area = _range(dimension) if dimension is not None else None
        if dimension is not None and area is None:
            found.add("dimension", format_range(label), label)
        elif bottom and (area is None or not (
                area[0] <= top and area[1] <= left and area[2] >= bottom
                and area[3] >= right)):
            found.add("dimension", format_range(label), label)
    return formulas, uses


#: Column letters (upper case, as written) and their numbers, filled as met: the common
#: reference is looked up rather than matched.
_LETTERS: dict[str, int] = {}


def _reference(ref: str, label: str, row_number: int, column: int,
               found: _Found) -> tuple[int, int]:
    """A cell's place from its reference (``r``), any but the common one: ``$``, lower
    case, another row's, or none at all (the column after ``column``, noted)."""
    row, col = row_number, column + 1
    match = _CELL.match(ref)
    if match is None or not 1 <= column_index(match.group(1).upper()) <= MAX_COLUMN \
            or not 1 <= int(match.group(2)) <= MAX_ROW:
        found.add("bad_reference", format_cell(label, row, col), label)
        return row, col
    col = column_index(match.group(1).upper())
    if int(match.group(2)) != row_number:
        found.add("cell_row", format_cell(label, int(match.group(2)), col), label)
    elif match.group(1).isupper() and "$" not in ref and len(_LETTERS) < 20000:
        _LETTERS[match.group(1)] = col
    return row, col


def _header(found: _Found, label: str, row: int, col: int, text: str,
            headers: dict[tuple[int, int], str]) -> None:
    if text != headers[(row, col)]:
        found.add("table_header", format_cell(label, row, col), label)


def _child(node: Any, local: str) -> Any:
    for ns in NS_MAIN:
        child = node.find(f"{{{ns}}}{local}")
        if child is not None:
            return child
    return None


def _number(text: str | None, *, zero: bool = False) -> int | None:
    if text is None or not text.strip().isdigit():
        return None
    value = int(text)
    return value if value > 0 or zero else None


def _range(text: str) -> tuple[int, int, int, int] | None:
    match = _RANGE.match(text.strip())
    if match is None:
        return None
    first = (int(match.group(2)), column_index(match.group(1).upper()))
    last = ((int(match.group(4)), column_index(match.group(3).upper())) if match.group(3)
            else first)
    return (min(first[0], last[0]), min(first[1], last[1]), max(first[0], last[0]),
            max(first[1], last[1]))


# -- shared strings and styles -----------------------------------------------------------------


class _StringItems:
    """The shared-string table as stored, without the workbook's redaction (compared,
    never reported): how many items it has, and an item's text when a table's header cell
    is compared with it -- read from the part's tree, item by item, only then."""

    def __init__(self, workbook: Any) -> None:
        self.part = workbook._related("sharedStrings")
        self._package = workbook.package
        self._items: list | None = None

    @property
    def items(self) -> list:
        if self._items is None:
            root = self._package.tree(self.part) if self.part else None
            self._items = [] if root is None else [node for node in root
                                                   if node.tag in _SI]
        return self._items

    def __len__(self) -> int:
        return len(self.items)

    def __getitem__(self, index: int) -> str:
        return item_text(self.items[index])


def _style_count(workbook: Any) -> int | None:
    part = workbook._related("styles")
    if part is None:
        return None
    root = workbook.package.tree(part)
    holder = _child(root, "cellXfs") if root is not None else None
    if holder is None:
        return 0
    return sum(1 for node in holder if isinstance(node.tag, str)
               and node.tag.rpartition("}")[2] == "xf")


def _string_counts(workbook: Any, found: _Found, uses: int | None) -> None:
    part = workbook._related("sharedStrings")
    root = workbook.package.tree(part)
    if root is None:
        return
    items = sum(1 for node in root if isinstance(node.tag, str)
                and node.tag.rpartition("}")[2] == "si")
    unique = root.get("uniqueCount")
    count = root.get("count")
    if unique is not None and unique.strip() != str(items):
        found.add("shared_string_count", part)
    elif uses is not None and count is not None and count.strip() != str(uses):
        found.add("shared_string_count", part)


# -- tables, names, the package ----------------------------------------------------------------


def _tables(workbook: Any, found: _Found) -> list:
    tables = workbook._table_parts()
    names: dict[str, str] = {}
    ids: dict[int, str] = {}
    by_sheet: dict[str, list] = {}
    for table in tables:
        if table.area is None:
            found.add("table_ref", table.part)
            continue
        if len(table.columns) != table.area.columns or any(
                flag.code == "columns_mismatch" for flag in table.flags):
            found.add("table_columns", table.part)
        folded = table.name.casefold()
        if folded in names or (table.id is not None and table.id in ids):
            found.add("table_name", table.part)
        names.setdefault(folded, table.part)
        if table.id is not None:
            ids.setdefault(table.id, table.part)
        for other in by_sheet.get(table.sheet, []):
            if _overlap(other.area, table.area):
                found.add("table_overlap", table.part)
        by_sheet.setdefault(table.sheet, []).append(table)
    return [t for t in tables if t.area is not None]


def _overlap(a: Any, b: Any) -> bool:
    return not (a.last.row < b.first.row or b.last.row < a.first.row
                or a.last.column < b.first.column or b.last.column < a.first.column)


def _names(workbook: Any, found: _Found) -> None:
    seen: set[tuple[str, str | None]] = set()
    for name in workbook._defined_names():
        if name.bad_scope is not None:
            found.add("name_scope", name.name)
        key = (name.name.casefold(), name.scope.casefold() if name.scope else name.bad_scope)
        if key in seen:
            found.add("name_duplicate", name.name)
        seen.add(key)


def _package(package: Any, found: _Found) -> None:
    parts = set(package.part_names)
    for source in sorted(parts):
        if not source.endswith(".rels"):
            continue
        owner = _owner(source)
        for rel in package.relationships(owner).values():
            if not rel.is_external and rel.target_part is not None \
                    and rel.target_part not in parts:
                found.add("missing_part", f"{source}#{rel.id}")
    root = package.tree("[Content_Types].xml")
    if root is None:
        return
    defaults = {node.get("Extension", "").lower() for node in root.iter(_CT + "Default")}
    overrides = {node.get("PartName", "").lstrip("/") for node in root.iter(_CT + "Override")}
    for name in sorted(overrides - parts):
        found.add("stray_override", name)
    for part in sorted(parts - {"[Content_Types].xml"}):
        extension = part.rpartition(".")[2].lower() if "." in part.rpartition("/")[2] else ""
        if part not in overrides and extension not in defaults:
            found.add("no_content_type", part)


def _owner(rels: str) -> str:
    """The part a relationships part belongs to (``""`` for the package's own)."""
    folder, _, name = rels.rpartition("/")
    folder = folder[:-len("_rels")].rstrip("/") if folder.endswith("_rels") else folder
    base = name[:-len(".rels")]
    owner = f"{folder}/{base}" if folder else base
    return owner if rels_path_for(owner) == rels else ""


def _calc_chain(workbook: Any, found: _Found, formulas: dict[str, set[tuple[int, int]]],
                every: list, checked: "dict[str, set[int] | None]") -> None:
    part = workbook._related("calcChain")
    if part is None:
        return
    by_id = {sheet.id: sheet for sheet in every if sheet.id is not None}
    sheet_id: int | None = None
    root = workbook.package.tree(part)
    for node in root if root is not None else ():
        if not isinstance(node.tag, str) or node.tag.rpartition("}")[2] != "c":
            continue
        written = node.get("i")
        if written is not None:
            sheet_id = _number(written, zero=True)
        sheet = by_id.get(sheet_id)
        match = _CELL.match(node.get("r") or "")
        if sheet is None or match is None or sheet.part not in formulas:
            continue    # a sheet not checked (or none): nothing to compare with
        position = (int(match.group(2)), column_index(match.group(1).upper()))
        only = checked.get(sheet.part)
        if only is not None and position[0] not in only:
            continue    # a row the step's check did not read
        if position not in formulas[sheet.part]:
            found.add("calc_chain", format_cell(_label(sheet), position[0], position[1]),
                      sheet.part)
