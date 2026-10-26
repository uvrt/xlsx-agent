"""The structure check: what an application would repair, found by code and address.

:meth:`Workbook.validate <xlsx_agent.Workbook.validate>` reads the package as it is now --
edited or not -- and lists what is out of order in the structure the format's parts must
keep with each other (ECMA-376 Part 1, §18; the package rules of Part 2).  It is the check
an edit is held to (every edit report carries it, :mod:`xlsx_agent.edit`) and the one a
save repeats on the bytes it hands over.  Each issue is a
:class:`~xlsx_agent.read.types.Problem`: a code, the first address it was met at and how
many times; never a cell's text, a formula or a name's formula (ROADMAP.md, "Design rules",
7).  Worksheets are streamed, never parsed whole.

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
from typing import Any, Iterable

from ooxml_edit.opc import rels_path_for
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, column_index

from ..read.strings import SharedStrings, item_text
from ..read.types import Problem, _letters
from ..xml import iterparse_xml

#: Every code :func:`check_structure` gives, in the order the module lists them.
CODES = ("row_order", "cell_order", "cell_row", "bad_reference", "shared_string",
         "shared_string_count", "style", "dimension", "table_ref", "table_overlap",
         "table_columns", "table_name", "table_header", "name_duplicate", "name_scope",
         "missing_part", "no_content_type", "stray_override", "calc_chain")

NS_MAIN = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
           "http://purl.oclc.org/ooxml/spreadsheetml/main")
_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
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


def check_structure(workbook: Any, *, sheets: "Iterable[str | None] | None" = None
                    ) -> list[Problem]:
    """The structure issues of ``workbook`` (an :class:`xlsx_agent.Workbook`), each by
    code with its first address and count; empty when there are none.  ``sheets`` limits
    the worksheet checks (cells, dimension, tables' headers) to those sheet parts; the
    workbook-wide checks always run."""
    found = _Found()
    package = workbook.package
    strings = _strings(workbook)
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
    uses = 0
    for sheet in chosen:
        label = _label(sheet)
        cells, used = _sheet(package.read(sheet.part) or b"", label, found, strings, styles,
                             headers.get(sheet.name, {}))
        formulas[sheet.part] = cells
        uses += used
    if strings is not None:
        _string_counts(workbook, found, uses if sheets is None else None)
    _names(workbook, found)
    _package(package, found)
    _calc_chain(workbook, found, formulas, every)
    return found.problems()


# -- worksheets --------------------------------------------------------------------------------


def _label(sheet: Any) -> str:
    from ..address import quote_sheet

    return quote_sheet(sheet.stored_name) if sheet.stored_name else f"#{sheet.id}"


def _sheet(data: bytes, label: str, found: _Found, strings: "SharedStrings | None",
           styles: int | None, headers: dict[tuple[int, int], str]
           ) -> tuple[set[tuple[int, int]], int]:
    """Stream one worksheet: its rows' and cells' order and references, each cell's
    shared string and style, the dimension against the cells, the tables' header cells.
    Returns the positions of the formula cells and how many cells use a shared string."""
    formulas: set[tuple[int, int]] = set()
    uses = 0
    dimension: str | None = None
    used: list[int] | None = None   # first row, first column, last row, last column
    row_number = 0
    column = 0
    in_data = False
    for event, node in iterparse_xml(data, events=("start", "end")):
        tag = node.tag
        if not isinstance(tag, str):
            continue
        ns, _, local = tag[1:].partition("}")
        if ns not in NS_MAIN:
            continue
        if event == "start":
            if local == "sheetData":
                in_data = True
            elif local == "row" and in_data:
                written = node.get("r")
                number = _number(written) if written is not None else row_number + 1
                if number is None or not 1 <= number <= MAX_ROW:
                    found.add("bad_reference", f"{label}!{written}", label)
                    number = row_number + 1
                elif number <= row_number:
                    found.add("row_order", f"{label}!{number}:{number}", label)
                row_number = max(row_number, number)
                column = 0
            continue
        if local == "sheetData":
            in_data = False
        elif local == "dimension":
            dimension = node.get("ref")
        elif local == "c" and in_data:
            written = node.get("r")
            row, col = row_number, column + 1
            if written is not None:
                match = _CELL.match(written)
                if match is None or not 1 <= column_index(match.group(1).upper()) <= MAX_COLUMN \
                        or not 1 <= int(match.group(2)) <= MAX_ROW:
                    found.add("bad_reference", f"{label}!{_letters(col)}{row}", label)
                else:
                    col = column_index(match.group(1).upper())
                    if int(match.group(2)) != row_number:
                        found.add("cell_row", f"{label}!{written.replace('$', '')}", label)
            address = f"{label}!{_letters(col)}{row}"
            if col <= column:
                found.add("cell_order", address, label)
            column = max(column, col)
            kind = node.get("t", "n")
            raw = _child(node, "v")
            has_formula = _child(node, "f") is not None
            if has_formula:
                formulas.add((row, col))
            style = node.get("s")
            if style is not None and styles is not None:
                index = _number(style, zero=True)
                if index is None or index >= styles:
                    found.add("style", address, label)
            text: str | None = None
            if kind == "s":
                uses += 1
                index = _number(raw.text.strip() if raw is not None and raw.text else None,
                                zero=True)
                if strings is None or index is None or index >= len(strings):
                    found.add("shared_string", address, label)
                elif (row, col) in headers:
                    text = strings[index]
            elif kind == "inlineStr" and (row, col) in headers:
                inline = _child(node, "is")
                text = item_text(inline) if inline is not None else None
            elif kind == "str" and (row, col) in headers and raw is not None:
                text = raw.text or ""
            if text is not None and text != headers[(row, col)]:
                found.add("table_header", address, label)
            if has_formula or raw is not None or kind == "inlineStr":
                if used is None:
                    used = [row, col, row, col]
                else:
                    used[0], used[1] = min(used[0], row), min(used[1], col)
                    used[2], used[3] = max(used[2], row), max(used[3], col)
            node.clear()
        elif local == "row" and in_data:
            node.clear()
            parent = node.getparent()
            while parent is not None and node.getprevious() is not None:
                del parent[0]
    if dimension is not None or used is not None:
        area = _range(dimension) if dimension is not None else None
        if dimension is not None and area is None:
            found.add("dimension", f"{label}!", label)
        elif used is not None and (area is None or not (
                area[0] <= used[0] and area[1] <= used[1]
                and area[2] >= used[2] and area[3] >= used[3])):
            found.add("dimension", f"{label}!", label)
    return formulas, uses


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


def _strings(workbook: Any) -> "SharedStrings | None":
    """The shared strings as stored, without the workbook's redaction: compared, never
    reported."""
    part = workbook._related("sharedStrings")
    return SharedStrings(workbook.package.read(part)) if part else None


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
                every: list) -> None:
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
        if position not in formulas[sheet.part]:
            found.add("calc_chain", f"{_label(sheet)}!{_letters(position[1])}{position[0]}",
                      sheet.part)
