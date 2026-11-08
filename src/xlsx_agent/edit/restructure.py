"""The package work of inserting, deleting and moving rows and columns (round 44): every
part that refers to a sheet's cells, rewritten through one :class:`~.grid.Change`.

:func:`restructure` checks everything first -- each refusal collected, none applied -- and
then rewrites, inside one step: the sheet's rows and cells, its column widths, merged
cells, rules (data validation, conditional formats, their ``x14`` forms), hyperlinks,
filter, panes, page breaks, tables, comments and their drawing, anchored drawings; every
formula of every sheet (shared formula groups kept shared where their members still read
alike), defined names, chart series, pivot cache sources and the tables' own formulas.  The
rules are :mod:`.grid`'s, measured from Excel; what this module does not know how to
rewrite is refused (:class:`~.grid.NotDone`), never written half-updated.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Any, Iterable

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, column_index, column_letters

from ..address import format_cell, format_sheet
from ..formula import ShiftError, shift as shift_formula, tokens
from .grid import (COLUMN, REF, ROW, VX, X, X14, XDR, XM, Area, Change, NotDone, R_ID,
                   Rewriter, Tables, anchored, area_cells, map_area, parse_area, sheet_key,
                   sqref_areas, sqref_text, _parse_part)

#: At most this many places are listed where a reference became ``#REF!``.
MAX_LISTED = 50
_TC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
#: Extensions of a worksheet that hold references this module rewrites, or none at all.
_EXT_HANDLED = {
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",     # x14 conditional formats
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",     # x14 data validations
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}",     # sparklines
}
#: Worksheet children whose references are not rewritten: a sheet holding one is refused.
_UNHANDLED = {"scenarios": "scenarios", "dataConsolidate": "a data consolidation",
              "customSheetViews": "custom views", "oleObjects": "embedded objects",
              "controls": "form objects", "cellWatches": "watched cells",
              "smartTags": "smart tags", "webPublishItems": "web publishing items"}


@dataclass(frozen=True)
class Restructured:
    """What an insert, delete or move did.  ``operation`` (``insert_rows``, ...), the
    ``sheet`` (as an address writes it), ``first``, ``count`` and for a move ``before``;
    ``changed`` (false for a move to where the lines are); ``formulas`` how many formulas
    were rewritten (cells, names, rules, chart series, table formulas); ``broken`` where a
    reference became ``#REF!`` (cells by address, ``name:X``, ``chart:part``; at most
    :data:`MAX_LISTED`, ``broken_count`` all); ``removed`` what went with deleted cells, by
    kind (merges, validations, conditional_formats, hyperlinks, comments, table_columns,
    tables); ``notes`` what the caller should know; ``calculation`` the workbook's status
    (``invalidated``: every cached value may be stale)."""

    operation: str
    sheet: str
    first: int
    count: int
    before: int | None = None
    changed: bool = True
    formulas: int = 0
    broken: tuple[str, ...] = ()
    broken_count: int = 0
    removed: dict[str, int] = field(default_factory=dict)
    notes: tuple[str, ...] = ()
    calculation: str = "invalidated"

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"operation": self.operation, "sheet": self.sheet,
                               "first": self.first, "count": self.count,
                               "changed": self.changed, "formulas": self.formulas,
                               "broken": list(self.broken),
                               "broken_count": self.broken_count,
                               "removed": dict(self.removed), "notes": list(self.notes),
                               "calculation": self.calculation}
        if self.before is not None:
            out["before"] = self.before
        return out


# -- the job -----------------------------------------------------------------------------------


@dataclass
class _Table:
    part: str
    root: Any
    area: Area
    name: str
    header: int
    totals: int

    @property
    def body(self) -> tuple[int, int] | None:
        first, last = self.area.r1 + self.header, self.area.r2 - self.totals
        return (first, last) if first <= last else None


@dataclass
class _Cell:
    node: Any
    row: int
    column: int
    text: str | None = None     # the formula's text, a shared member's computed
    group: str | None = None    # the shared group it was in


class Job:
    """One operation on one sheet: :meth:`check` lists every refusal, :meth:`apply`
    rewrites the package (inside the caller's step)."""

    def __init__(self, workbook: Any, sheet: Any, change: Change, operation: str) -> None:
        self.wb = workbook
        self.package = workbook.package
        self.sheet = sheet
        self.part: str = sheet.part
        self.change = change
        self.operation = operation
        self.root = self.package.tree(self.part)
        self.label = format_sheet(sheet.stored_name, sheet.id)
        self.problems: list[tuple[str, str, str]] = []
        self.tables = self._tables()
        bodies = []
        for table in self.tables:
            body = table.body
            if body is not None:
                bodies.append((body[0], table.area.c1, body[1], table.area.c2))
        self.change = change = Change(change.kind, change.axis, change.first, change.count,
                                      change.before, tuple(bodies))
        self.gone = Tables()
        self.rewriter = Rewriter(change, sheet.stored_name, self.gone)
        self.target = sheet_key(sheet.stored_name)
        self.formulas = 0
        self.broken: list[str] = []
        self.broken_count = 0
        self.removed: dict[str, int] = {}
        self.notes: list[str] = []
        self._keep_one: _Table | None = None
        self._sheets = [s for s in workbook._every_sheet
                        if s.part is not None and s.type == "worksheet"
                        and self.package.has_part(s.part)]

    # -- reading -------------------------------------------------------------------------------

    def _tables(self) -> list[_Table]:
        out = []
        for rel in self.package.relationships(self.part).values():
            if rel.is_external or rel.type.rpartition("/")[2] != "table" \
                    or rel.target_part is None or not self.package.has_part(rel.target_part):
                continue
            root = self.package.tree(rel.target_part)
            area = parse_area(root.get("ref", ""))
            if area is None:
                continue
            header = int(root.get("headerRowCount", "1") or 1)
            totals = int(root.get("totalsRowCount", "0") or 0)
            out.append(_Table(rel.target_part, root, area,
                              root.get("displayName") or root.get("name") or "", header,
                              totals))
        return out

    def refuse(self, code: str, where: str, detail: str) -> None:
        self.problems.append((code, where, detail))

    def _where(self, area: "Area | None" = None, row: int | None = None,
               column: int | None = None) -> str:
        if area is not None:
            return f"{self.label}!{area.ref}"
        if row is not None and column is not None:
            return format_cell(self.label, row, column)
        return self.label

    # -- the checks ----------------------------------------------------------------------------

    def check(self) -> None:
        """Every refusal, collected in :attr:`problems` (nothing changes)."""
        root = self.root
        data = root.find(X + "sheetData")
        if data is None:
            self.refuse("not_a_worksheet", self.label, "the sheet holds no cells")
            return
        for local, what in _UNHANDLED.items():
            if root.find(X + local) is not None:
                self.refuse("not_measured", self.label, f"the sheet holds {what}")
        for ext in root.iterfind(f"{X}extLst/{X}ext"):
            if ext.get("uri") not in _EXT_HANDLED and (
                    next(ext.iter(XM + "f"), None) is not None
                    or next(ext.iter(XM + "sqref"), None) is not None):
                self.refuse("not_measured", self.label,
                            f"an extension ({ext.get('uri')}) holding references")
        self._check_tables()
        if self._keep_one is not None:
            # Every data row of a table deleted: Excel deletes all but the first and
            # empties that one (measured: references into it stay, B2:B7 is B2:B2).
            c = self.change
            self.change = Change("delete", ROW, c.first + 1, c.count - 1, bodies=c.bodies)
            self.rewriter = Rewriter(self.change, self.sheet.stored_name, self.gone)
        self._check_cells(data)
        self._check_merges()
        self._check_pivots()
        self._check_drawings()
        self._check_names()

    def _check_cells(self, data: Any) -> None:
        change = self.change
        edge = change.limit - change.count if change.kind != "delete" else None
        axis_row = change.axis == ROW
        if edge is not None:
            # Only a sheet's last lines can be pushed off: its last rows, each row's last cell.
            rows = list(data.iterchildren(X + "row"))
            if axis_row:
                candidates = [c for row in reversed(rows) if _int(row.get("r")) is not None
                              and int(row.get("r")) > edge for c in row.iterchildren(X + "c")]
            else:
                candidates = []
                for row in rows:
                    last = None
                    for last in row.iterchildren(X + "c", reversed=True):
                        break
                    if last is not None:
                        candidates.append(last)
            for cell in candidates:
                position = _position(cell.get("r"))
                if position is None:
                    continue
                line = position[0] if axis_row else position[1]
                if line > edge and not (change.kind == "move"
                                        and change.first <= line <= change.last) and (
                        cell.find(X + "f") is not None or cell.find(X + "v") is not None
                        or cell.find(X + "is") is not None):
                    self.refuse("off_grid", self._where(row=position[0], column=position[1]),
                                f"{change.count} more would push it off the sheet")
        for f in data.iter(X + "f"):
            kind = f.get("t")
            if kind not in ("array", "dataTable"):
                continue
            ref = f.getparent().get("r")
            area = parse_area(f.get("ref") or ref or "")
            if area is None:
                continue
            a, b = area.along(change.axis)
            if kind == "dataTable":
                if change.kind != "insert" or change.first <= b:
                    self.refuse("not_measured", self._where(area), "a data table (what-if table)")
                continue
            if change.splits(a, b):
                self.refuse("array_split", self._where(area),
                            "an array formula is changed only as a whole: Excel refuses to "
                            "cut through it")

    def _check_tables(self) -> None:
        change = self.change
        cut = [t for t in self.tables if _cuts(change, t.area)]
        if len(cut) > 1:
            names = ", ".join(f"t:{t.name}" for t in cut)
            if change.axis == ROW:
                self.refuse("table_rows", names, "Excel refuses to insert, delete or move "
                            "rows through two tables side by side")
            else:
                self.refuse("not_measured", names, "columns through two tables at once")
        for table in self.tables:
            where = f"t:{table.name}"
            area = table.area
            if change.axis == COLUMN:
                if change.kind == "move":
                    c1, c2 = area.c1, area.c2
                    if not (change.last < c1 or c2 < change.first) or c1 < change.before <= c2:
                        self.refuse("table_move", where,
                                    "Excel refuses to move columns into, out of or within a "
                                    "table")
                continue
            body = table.body
            r1, r2 = area.r1, area.r2
            if change.kind == "insert":
                if table.totals and change.first > r2 - table.totals \
                        and change.first <= r2 and change.first != r2 - table.totals + 1:
                    self.refuse("not_measured", where, "rows inserted inside a totals row")
                continue
            if change.kind == "delete":
                first, last = change.first, change.last
                if last < r1 or r2 < first:
                    continue
                if table.header and first <= r1 <= last:
                    self.refuse("table_header_row", where,
                                "Excel refuses to delete a table's header row")
                    continue
                if body is not None and first <= body[0] and body[1] <= last:
                    if (first, last) == body:
                        self._keep_one = table      # measured: Excel keeps one, empty
                    else:
                        self.refuse("not_measured", where, "every data row of a table "
                                    "deleted together with other rows")
                    continue
                if table.totals:
                    t1 = r2 - table.totals + 1
                    if first <= r2 and last >= t1 and first < t1:
                        self.refuse("not_measured", where,
                                    "data rows deleted together with the totals row")
                continue
            # a move of rows
            first, last, before = change.first, change.last, change.before
            inside = r1 + table.header <= first and last <= r2 - table.totals
            outside = last < r1 or r2 < first
            if inside:
                if not (r1 + table.header <= before <= r2 - table.totals + 1):
                    self.refuse("table_move", where,
                                "Excel refuses to move rows out of a table")
            elif outside:
                if r1 < before <= r2:
                    if table.header and before <= r1:
                        continue
                    if before > r2 - table.totals + 1:
                        self.refuse("not_measured", where, "rows moved into a totals row")
            else:
                self.refuse("table_move", where,
                            "the rows moved cut through a table's edge")

    def _check_merges(self) -> None:
        change = self.change
        if change.kind != "move":
            return
        merges = self.root.find(X + "mergeCells")
        for merge in (merges if merges is not None else []):
            area = parse_area(merge.get("ref", ""))
            if area is None:
                continue
            a, b = area.along(change.axis)
            if a == b:
                continue
            if change.splits(a, b) or (
                    (b < change.first or change.last < a) and a < change.before <= b):
                self.refuse("merge_split", self._where(area),
                            "Excel refuses to move part of a merged area")

    def _check_pivots(self) -> None:
        change = self.change
        for rel in self.package.relationships(self.part).values():
            if rel.is_external or rel.type.rpartition("/")[2] != "pivotTable" \
                    or rel.target_part is None or not self.package.has_part(rel.target_part):
                continue
            location = self.package.tree(rel.target_part).find(X + "location")
            area = parse_area(location.get("ref", "")) if location is not None else None
            if area is None:
                continue
            a, b = area.along(change.axis)
            touched = change.splits(a, b) or (change.kind == "delete" and
                                              change.first <= a and b <= change.last) \
                or (change.kind == "move" and not (b < change.first or change.last < a))
            if touched:
                self.refuse("pivot_table", self._where(area),
                            "the change would cut through a pivot table")

    def _check_drawings(self) -> None:
        change = self.change
        for kind, part in self._related_parts():
            if kind == "drawing":
                root = self.package.tree(part)
                for anchor in root:
                    span = _anchor_span(anchor, change.axis)
                    if span is None:
                        continue
                    a, b = span
                    if change.kind == "delete" and change.first <= a and b <= change.last:
                        self.refuse("not_measured", self.label,
                                    "a drawing object wholly inside the deleted lines")
                    if change.kind == "move" and (change.splits(a, b) or (
                            (b < change.first or change.last < a)
                            and a < change.before <= b)):
                        self.refuse("not_measured", self.label,
                                    "a drawing object the moved lines cut through")
            elif kind == "vmlDrawing":
                text = self.package.read(part) or b""
                if b"ObjectType=" in text and any(
                        f'ObjectType="{t}"'.encode() in text
                        for t in ("Button", "Checkbox", "Drop", "List", "Radio", "Spin",
                                  "Scroll", "GBox", "Label", "Dialog", "Edit")):
                    self.refuse("not_measured", self.label, "form objects in a VML drawing")

    def _check_names(self) -> None:
        names = self.wb._root.find(X + "definedNames")
        for node in (names if names is not None else []):
            text = node.text or ""
            if not self.rewriter.mentions(text):
                continue
            try:
                found = tokens(text)
            except ShiftError:
                self.refuse("not_measured", f"name:{node.get('name')}",
                            "a defined name whose formula cannot be read")
                continue
            for token in found:
                if token.kind != "ref" or token.book or not token.sheet:
                    continue
                key = _prefix_key(token.sheet)
                if key != self.target:
                    continue
                kinds = _anchoring(token.body, self.change.axis)
                if kinds == "mixed":
                    self.refuse("not_measured", f"name:{node.get('name')}",
                                "a defined name with a reference anchored in part")

    # -- applying ------------------------------------------------------------------------------

    def apply(self) -> None:
        """Rewrite the package (inside the caller's step)."""
        self._table_columns_gone()
        for sheet in self._sheets:
            self._formulas_of(sheet)
        self._names()
        self._charts()
        self._pivot_caches()
        self._other_tables()
        self._rules_elsewhere()
        self._sheet()

    # tables whose columns go (structured references) -------------------------------------------

    def _table_columns_gone(self) -> None:
        change = self.change
        if change.axis != COLUMN or change.kind != "delete":
            return
        for table in self.tables:
            columns = table.root.find(X + "tableColumns")
            if columns is None:
                continue
            nodes = columns.findall(X + "tableColumn")
            gone = set()
            for offset, node in enumerate(nodes):
                if change.first <= table.area.c1 + offset <= change.last:
                    gone.add((node.get("name") or "").casefold())
            if not gone:
                continue
            key = table.name.casefold()
            self.gone.deleted[key] = None if len(gone) == len(nodes) else gone

    # formulas -----------------------------------------------------------------------------------

    def _rewrite(self, text: str, home: str | None, where: str | None,
                 table: str | None = None) -> str:
        """``where`` names the place in the report; ``None``: a cell that goes."""
        try:
            new = self.rewriter.formula(text, home, table)
        except ShiftError:
            raise NotDone("not_measured", where or self.label,
                          "a formula that cannot be read") from None
        if new != text and where is not None:
            self.formulas += 1
            if REF in new and new.count(REF) > text.count(REF):
                self.broken_count += 1
                if len(self.broken) < MAX_LISTED:
                    self.broken.append(where)
        return new

    def _formulas_of(self, sheet: Any) -> None:
        """Every formula on ``sheet`` rewritten; shared groups kept where they read alike.
        On the operated sheet the cells are only read here; they move in :meth:`_sheet`."""
        root = self.package.tree(sheet.part)
        data = root.find(X + "sheetData")
        if data is None:
            return
        home = sheet_key(sheet.stored_name)
        on_target = sheet.part == self.part
        label = format_sheet(sheet.stored_name, sheet.id)
        if not on_target and not self._sheet_may_refer(sheet.part):
            return
        tables = self._cell_tables(sheet.part)
        cells, groups = _formula_cells(data)
        changed = False
        for cell in cells:
            if cell.text is None:
                continue
            table = _table_at(tables, cell.row, cell.column)
            place = _moved_cell((cell.row, cell.column), self.change) if on_target \
                else (cell.row, cell.column)
            where = format_cell(label, *place) if place is not None else None
            new = self._rewrite(cell.text, home, where, table)
            if new != cell.text:
                cell.text = new
                changed = True
                if cell.group is None:
                    cell.node.find(X + "f").text = new
        if on_target:
            self._target_cells = (cells, groups)
            return
        if changed:
            for group in groups.values():
                _reshare(group, None)
            self.package.mark_dirty(sheet.part)
            self._touched_rules(sheet, root, home, label)
        else:
            self._touched_rules(sheet, root, home, label)

    def _sheet_may_refer(self, part: str) -> bool:
        data = self.package.read(part) or b""
        if self.gone.deleted and b"[" in data:
            return True
        folded = data.decode("utf-8", "replace").casefold()
        return any(needle in folded for needle in self.rewriter._needles)

    def _cell_tables(self, part: str) -> list[tuple[Area, str]]:
        out = []
        for rel in self.package.relationships(part).values():
            if rel.is_external or rel.type.rpartition("/")[2] != "table" \
                    or rel.target_part is None or not self.package.has_part(rel.target_part):
                continue
            root = self.package.tree(rel.target_part)
            area = parse_area(root.get("ref", ""))
            if area is not None:
                out.append((area, root.get("displayName") or root.get("name") or ""))
        return out

    def _touched_rules(self, sheet: Any, root: Any, home: str, label: str) -> None:
        """Another sheet's rule formulas (``x14`` rules may refer across sheets)."""
        changed = False
        for node in _rule_formulas(root):
            text = node.text or ""
            new = self._rewrite(text, home, label)
            if new != text:
                node.text = new
                changed = True
        if changed:
            self.package.mark_dirty(sheet.part)

    def _names(self) -> None:
        root = self.wb._root
        names = root.find(X + "definedNames")
        if names is None:
            return
        changed = False
        for node in names:
            text = node.text or ""
            new = self._rewrite_name(text, f"name:{node.get('name')}")
            if new != text:
                node.text = new
                changed = True
        if changed:
            self.package.mark_dirty(self.wb.part)

    def _rewrite_name(self, text: str, where: str) -> str:
        """A defined name's formula: its fully relative references stay (measured)."""
        if not self.rewriter.mentions(text) and not (self.gone.deleted and "[" in text):
            return text
        try:
            found = tokens(text)
        except ShiftError:
            raise NotDone("not_measured", where, "a formula that cannot be read") from None
        out = []
        for token in found:
            if token.kind == "ref" and not token.book and token.sheet \
                    and _anchoring(token.body, self.change.axis) == "relative":
                out.append(token.text)
                continue
            out.append(self._rewrite(token.text, None, where)
                       if token.kind in ("ref", "structured") else token.text)
        return "".join(out)

    def _charts(self) -> None:
        for part in self.package.part_names:
            if not (part.startswith("xl/charts/") and part.endswith(".xml")):
                continue
            data = self.package.read(part) or b""
            if not self.rewriter.mentions(data.decode("utf-8", "replace")):
                continue
            root = self.package.tree(part)
            changed = False
            for node in root.iter():
                if not isinstance(node.tag, str) or etree.QName(node).localname != "f":
                    continue
                text = node.text or ""
                new = self._rewrite(text, None, f"chart:{part}")
                if new != text:
                    node.text = new
                    changed = True
            if changed:
                self.package.mark_dirty(part)

    def _pivot_caches(self) -> None:
        for part in self.package.part_names:
            if not part.startswith("xl/pivotCache/pivotCacheDefinition"):
                continue
            root = self.package.tree(part)
            changed = False
            for source in root.iter(X + "worksheetSource"):
                if source.get("sheet") is None or sheet_key(source.get("sheet")) != self.target \
                        or source.get("ref") is None:
                    continue
                area = parse_area(source.get("ref"))
                if area is None:
                    continue
                moved = map_area(area, self.change)
                if moved is None:
                    continue        # measured: a source wholly deleted is left as it was
                if moved.ref != area.ref:
                    source.set("ref", moved.ref)
                    changed = True
                    self.formulas += 1
            if changed:
                self.package.mark_dirty(part)

    def _other_tables(self) -> None:
        """Every table's own formulas (calculated columns, totals) rewritten."""
        for sheet in self._sheets:
            home = sheet_key(sheet.stored_name)
            for area, name in self._cell_tables(sheet.part):
                for rel in self.package.relationships(sheet.part).values():
                    if rel.is_external or rel.type.rpartition("/")[2] != "table":
                        continue
                    root = self.package.tree(rel.target_part)
                    if (root.get("displayName") or root.get("name") or "") != name:
                        continue
                    changed = False
                    for node in root.iter(X + "calculatedColumnFormula",
                                          X + "totalsRowFormula"):
                        text = node.text or ""
                        new = self._rewrite(text, home, f"t:{name}", name)
                        if new != text:
                            node.text = new
                            changed = True
                    if changed:
                        self.package.mark_dirty(rel.target_part)

    def _rules_elsewhere(self) -> None:
        """(Done with each sheet's formulas: :meth:`_touched_rules`.)"""

    # the operated sheet -------------------------------------------------------------------------

    def _sheet(self) -> None:
        root = self.root
        change = self.change
        data = root.find(X + "sheetData")
        cells, groups = getattr(self, "_target_cells", ([], {}))
        neighbour = self._neighbour()
        self._rows_and_cells(data, cells, groups, neighbour)
        if change.axis == COLUMN:
            self._columns(root, neighbour)
        self._merges(root)
        self._rules(root, neighbour if change.kind == "insert" else self._moved_neighbour())
        self._hyperlinks(root)
        self._filter(root)
        self._views(root)
        self._breaks(root)
        self._sqrefs(root)
        self._tables_on_sheet()
        self._fill_calculated(data)
        self._comments()
        self._drawings()
        self._pivot_locations()
        _dimension(root, data)
        self.package.mark_dirty(self.part)

    def _moved_neighbour(self) -> int | None:
        """A move inserts before placing (Excel's "Insert Cut Cells"): the moved lines also
        take the rules (validations, conditional formats) of the line before their new
        place -- numbered as after the move (measured); none from a hidden line."""
        change = self.change
        if change.kind != "move" or change.target <= 1:
            return None
        line = change.before - 1
        if change.axis == ROW:
            row = _row_node(self.root.find(X + "sheetData"), line)
            if row is not None and _true(row.get("hidden")):
                return None
        else:
            col = _col_node(self.root, line)
            if col is not None and _true(col.get("hidden")):
                return None
        return change.target - 1

    def _neighbour(self) -> int | None:
        """The line an insert takes its format from (above, or to the left), ``None``
        when there is none or it is hidden (measured: nothing comes from a hidden line)."""
        change = self.change
        if change.kind != "insert" or change.first <= 1:
            return None
        line = change.first - 1
        if change.axis == ROW:
            row = _row_node(self.root.find(X + "sheetData"), line)
            if row is not None and _true(row.get("hidden")):
                return None
        else:
            col = _col_node(self.root, line)
            if col is not None and _true(col.get("hidden")):
                return None
        return line

    def _rows_and_cells(self, data: Any, cells: list[_Cell], groups: dict,
                        neighbour: int | None) -> None:
        change = self.change
        axis_row = change.axis == ROW
        rows = list(data.iterchildren(X + "row"))
        _number_rows(rows)
        removed_cells: list[Any] = []
        if axis_row:
            placed: list[tuple[int, Any]] = []
            for row in rows:
                number = int(row.get("r"))
                new = change.index(number)
                if new is None:
                    removed_cells.extend(row.iterchildren(X + "c"))
                    data.remove(row)
                    continue
                if new != number:
                    row.set("r", str(new))
                    digits = str(new)
                    for cell in row.iterchildren(X + "c"):
                        ref = cell.get("r")
                        cell.set("r", ref.rstrip("0123456789") + digits)
                placed.append((new, row))
            if self._keep_one is not None:
                self._clear_kept_row(placed)
            if neighbour is not None:
                placed += self._inserted_rows(placed, neighbour)
            if change.kind != "delete":
                placed.sort(key=lambda item: item[0])
                for _, row in placed:
                    data.append(row)
        else:
            for row in rows:
                kept: list[tuple[int, Any]] = []
                for cell in list(row.iterchildren(X + "c")):
                    position = _position(cell.get("r"))
                    if position is None:
                        continue
                    new = change.index(position[1])
                    if new is None:
                        removed_cells.append(cell)
                        row.remove(cell)
                        continue
                    if new != position[1]:
                        cell.set("r", f"{column_letters(new)}{position[0]}")
                    kept.append((new, cell))
                if neighbour is not None:
                    kept += _inserted_cells(row, neighbour, change)
                kept.sort(key=lambda item: item[0])
                for _, cell in kept:
                    row.append(cell)
                if row.get("spans") is not None:
                    if kept:
                        row.set("spans", f"{kept[0][0]}:{kept[-1][0]}")
                    else:
                        del row.attrib["spans"]
        gone = {id(node) for node in removed_cells}
        for cell in cells:
            if id(cell.node) not in gone:
                position = _position(cell.node.get("r"))
                assert position is not None
                cell.row, cell.column = position
        survivors = [cell for cell in cells if id(cell.node) not in gone]
        for cell in survivors:
            f = cell.node.find(X + "f")
            if f is not None and f.get("t") == "array" and f.get("ref"):
                area = parse_area(f.get("ref"))
                moved = map_area(area, change) if area is not None else None
                if moved is not None:
                    f.set("ref", moved.ref)
        for group in groups.values():
            _reshare([c for c in group if id(c.node) not in gone], None)

    def _clear_kept_row(self, placed: list[tuple[int, Any]]) -> None:
        """Every data row of a table deleted: Excel keeps one, empty (measured)."""
        table = self._keep_one
        assert table is not None
        first = self.change.first - 1
        for number, row in placed:
            if number != first:
                continue
            for cell in list(row.iterchildren(X + "c")):
                position = _position(cell.get("r"))
                if position and table.area.c1 <= position[1] <= table.area.c2:
                    for child in list(cell):
                        cell.remove(child)
                    if cell.get("t") is not None:
                        del cell.attrib["t"]
            # its hyperlinks go with the contents (measured)
            links = self.root.find(X + "hyperlinks")
            for link in list(links if links is not None else []):
                area = parse_area(link.get("ref", ""))
                if area is not None and area.r1 == area.r2 == first \
                        and table.area.c1 <= area.c1 and area.c2 <= table.area.c2:
                    links.remove(link)
                    self._count("hyperlinks")
                    rid = link.get(R_ID)
                    if rid and not any(other.get(R_ID) == rid for other in links):
                        self.package.remove_relationship(self.part, rid)
            if links is not None and not len(links):
                self.root.remove(links)

    def _inserted_rows(self, placed: list[tuple[int, Any]], neighbour: int
                       ) -> list[tuple[int, Any]]:
        """New rows taking the format of the row above (its height, style and outline;
        never hidden) and its cells' styles."""
        change = self.change
        above = next((row for number, row in placed if number == neighbour), None)
        if above is None:
            return []
        out = []
        for offset in range(change.count):
            number = change.first + offset
            row = etree.Element(X + "row")
            row.set("r", str(number))
            for name in ("s", "customFormat", "ht", "customHeight", "outlineLevel",
                         "thickTop", "thickBot"):
                if above.get(name) is not None:
                    row.set(name, above.get(name))
            for prefixed in above.attrib:
                if prefixed.startswith("{") and prefixed.endswith("}dyDescent"):
                    row.set(prefixed, above.get(prefixed))
            for cell in above.iterchildren(X + "c"):
                position = _position(cell.get("r"))
                if position is None or cell.get("s") in (None, "0"):
                    continue
                if any(t.header and t.area.r1 == neighbour
                       and t.area.c1 <= position[1] <= t.area.c2 for t in self.tables):
                    continue
                new = etree.SubElement(row, X + "c")
                new.set("r", f"{column_letters(position[1])}{number}")
                new.set("s", cell.get("s"))
            if above.get("spans") is not None:
                row.set("spans", above.get("spans"))
            if len(row) or len(row.attrib) > 1:
                out.append((number, row))
        return out

    def _columns(self, root: Any, neighbour: int | None) -> None:
        """Column widths and styles (``cols``) through a column change."""
        change = self.change
        cols = root.find(X + "cols")
        if cols is None:
            return
        entries = []
        for col in cols.findall(X + "col"):
            lo, hi = int(col.get("min", "0")), int(col.get("max", "0"))
            if lo < 1 or hi < lo:
                continue
            for x, y in change.pieces(lo, hi) if change.kind == "move" else \
                    _col_pieces(change, lo, hi):
                entries.append((x, y, col))
        if neighbour is not None:
            source = next((col for x, y, col in entries if x <= neighbour <= y
                           and not _true(col.get("hidden"))), None)
            if source is not None:
                entries.append((change.first, change.last, source))
        for col in cols.findall(X + "col"):
            cols.remove(col)
        entries.sort(key=lambda e: e[0])
        last_end = 0
        for x, y, source in entries:
            if x <= last_end:
                x = last_end + 1
            if y < x:
                continue
            node = copy.deepcopy(source)
            node.set("min", str(x))
            node.set("max", str(y))
            cols.append(node)
            last_end = y
        if not len(cols):
            root.remove(cols)

    def _merges(self, root: Any) -> None:
        merges = root.find(X + "mergeCells")
        if merges is None:
            return
        for merge in list(merges):
            area = parse_area(merge.get("ref", ""))
            if area is None:
                continue
            moved = area_cells(area, self.change)
            if len(moved) != 1 or (moved[0].r1, moved[0].c1) == (moved[0].r2, moved[0].c2):
                merges.remove(merge)
                self._count("merges")
                continue
            merge.set("ref", moved[0].ref)
        if not len(merges):
            root.remove(merges)
        else:
            merges.set("count", str(len(merges)))

    def _rules(self, root: Any, neighbour: int | None) -> None:
        """Data validations and conditional formats, main and ``x14``.  ``neighbour``
        the line whose rules new lines take.  Measured for a move: validations always;
        conditional formats only when the lines move up (or left) and the format's range
        ends just before their new place, never when it lies across that place."""
        change = self.change

        def format_neighbour_of(areas: list[Area]) -> int | None:
            if change.kind == "insert":
                return neighbour
            if neighbour is None or change.before > change.first:
                return None
            for area in areas:
                a, b = area.along(change.axis)
                if a < change.before <= b:
                    return None
            return neighbour
        home = self.target
        priorities = [int(r.get("priority")) for r in root.iter(X + "cfRule")
                      if (r.get("priority") or "").isdigit()]
        next_priority = [max(priorities, default=0) + 1]
        # main conditional formats
        for block in list(root.iterfind(X + "conditionalFormatting")):
            areas = sqref_areas(block.get("sqref", ""))
            if areas is None:
                raise NotDone("not_measured", self.label, "a rule's range that cannot be read")
            rules = block.findall(X + "cfRule")
            formulas = [[f.text or "" for f in rule.findall(X + "formula")] for rule in rules]
            groups = self._regroup(areas, formulas, format_neighbour_of(areas),
                                   linked=len(areas) == 1)
            if not groups:
                root.remove(block)
                self._count("conditional_formats")
                continue
            self._place_groups(block, groups, rules, "sqref", X + "formula", next_priority)
        for block in list(root.iter(X14 + "conditionalFormatting")):
            sq = block.find(XM + "sqref")
            areas = sqref_areas(sq.text or "") if sq is not None else None
            if areas is None:
                raise NotDone("not_measured", self.label, "a rule's range that cannot be read")
            rules = block.findall(X14 + "cfRule")
            formulas = [[f.text or "" for f in rule.iter(XM + "f")] for rule in rules]
            groups = self._regroup(areas, formulas, format_neighbour_of(areas),
                                   absolute_ok=True)
            if not groups:
                block.getparent().remove(block)
                self._count("conditional_formats")
                continue
            self._place_groups(block, groups, rules, None, XM + "f", next_priority)
        # data validations
        validations = root.find(X + "dataValidations")
        for dv in list(validations if validations is not None else []):
            areas = sqref_areas(dv.get("sqref", ""))
            if areas is None:
                raise NotDone("not_measured", self.label, "a rule's range that cannot be read")
            nodes = [dv.find(X + "formula1"), dv.find(X + "formula2")]
            formulas = [[n.text or "" if n is not None else "" for n in nodes]]
            groups = self._regroup(areas, formulas, neighbour)
            if not groups:
                validations.remove(dv)
                self._count("validations")
                continue
            for index, (group_areas, texts) in enumerate(groups):
                node = dv if index == 0 else copy.deepcopy(dv)
                if index:
                    dv.addnext(node)
                node.set("sqref", sqref_text(group_areas))
                for name, text in zip(("formula1", "formula2"), texts[0]):
                    child = node.find(X + name)
                    if child is not None:
                        child.text = text
        if validations is not None:
            if not len(validations):
                root.remove(validations)
            else:
                validations.set("count", str(len(validations)))
        for block in root.iter(X14 + "dataValidations"):
            for dv in list(block.findall(X14 + "dataValidation")):
                sq = dv.find(XM + "sqref")
                areas = sqref_areas(sq.text or "") if sq is not None else None
                if areas is None:
                    raise NotDone("not_measured", self.label,
                                  "a rule's range that cannot be read")
                nodes = list(dv.iter(XM + "f"))
                groups = self._regroup(areas, [[n.text or "" for n in nodes]], neighbour,
                                       absolute_ok=True)
                if not groups:
                    block.remove(dv)
                    self._count("validations")
                    continue
                for index, (group_areas, texts) in enumerate(groups):
                    node = dv if index == 0 else copy.deepcopy(dv)
                    if index:
                        dv.addnext(node)
                    node.find(XM + "sqref").text = sqref_text(group_areas)
                    for child, text in zip(node.iter(XM + "f"), texts[0]):
                        child.text = text
            block.set("count", str(len(block.findall(X14 + "dataValidation"))))
        # sparklines: data ranges are references; locations follow their cells
        for spark in root.iter("{http://schemas.microsoft.com/office/spreadsheetml/2009/9/"
                               "main}sparkline"):
            f = spark.find(XM + "f")
            if f is not None and f.text:
                f.text = self._rewrite(f.text, home, self.label)
            sq = spark.find(XM + "sqref")
            area = parse_area(sq.text or "") if sq is not None else None
            if area is not None:
                moved = area_cells(area, self.change)
                if not moved:
                    spark.getparent().remove(spark)
                else:
                    sq.text = moved[0].ref

    def _regroup(self, areas: list[Area], formulas: list[list[str]],
                 neighbour: int | None, absolute_ok: bool = False, linked: bool = True
                 ) -> list[tuple[list[Area], list[list[str]]]]:
        """A rule's areas and formulas through the change (see
        :func:`~.grid.anchored`), and grown by an insert after a line it covers."""
        flat = [text for texts in formulas for text in texts]
        try:
            groups = anchored(areas, flat, self.change, self.rewriter, self.target,
                              linked=linked)
        except ShiftError:
            raise NotDone("not_measured", self.label,
                          "a rule's formula that cannot be read or moved") from None
        out = []
        for group_areas, texts in groups:
            if neighbour is not None:
                group_areas = [self._grown(area, neighbour) for area in group_areas]
            self.formulas += sum(1 for a, b in zip(texts, flat) if a != b)
            nested, at = [], 0
            for original in formulas:
                nested.append(texts[at:at + len(original)])
                at += len(original)
            out.append((group_areas, nested))
        return out

    def _grown(self, area: Area, neighbour: int) -> Area:
        """An area that ends on the line an insert follows grows into the new lines."""
        a, b = area.along(self.change.axis)
        if b == neighbour:
            return area.with_along(self.change.axis, a, b + self.change.count)
        return area

    def _place_groups(self, block: Any, groups: list, rules: list, attribute: str | None,
                      formula_tag: str, next_priority: list[int]) -> None:
        for index, (group_areas, texts) in enumerate(groups):
            node = block if index == 0 else copy.deepcopy(block)
            if index:
                block.addnext(node)
                for rule in node.iter(X + "cfRule", X14 + "cfRule"):
                    rule.set("priority", str(next_priority[0]))
                    next_priority[0] += 1
            if attribute is not None:
                node.set(attribute, sqref_text(group_areas))
            else:
                node.find(XM + "sqref").text = sqref_text(group_areas)
            node_rules = [r for r in node if etree.QName(r).localname == "cfRule"]
            for rule, rule_texts in zip(node_rules, texts):
                for child, text in zip(rule.iter(formula_tag), rule_texts):
                    child.text = text

    def _hyperlinks(self, root: Any) -> None:
        links = root.find(X + "hyperlinks")
        if links is None:
            return
        for link in list(links):
            area = parse_area(link.get("ref", ""))
            if area is None:
                continue
            moved = area_cells(area, self.change)
            if not moved:
                links.remove(link)
                self._count("hyperlinks")
                rid = link.get(R_ID)
                if rid and not any(l.get(R_ID) == rid for l in links):
                    self.package.remove_relationship(self.part, rid)
                continue
            link.set("ref", _bbox(moved).ref)
        if not len(links):
            root.remove(links)

    def _filter(self, root: Any) -> None:
        """The sheet's filter: a range that keeps its columns' places (moves shrink and
        grow it as a delete and an insert would: measured)."""
        node = root.find(X + "autoFilter")
        if node is None:
            return
        area = parse_area(node.get("ref", ""))
        if area is None:
            return
        change = self.change
        if change.kind == "move":
            a, b = area.along(change.axis)
            if change.first <= a and b <= change.last:
                moved = map_area(area, change)
            else:
                span = change._delete_insert(a, b, area.across(change.axis))
                moved = None if span is None else area.with_along(change.axis, *span)
        else:
            moved = map_area(area, change)
        if moved is None:
            root.remove(node)
            self._drop_filter_name()
            return
        _shift_filter_columns(node, area, moved, change)
        node.set("ref", moved.ref)
        if change.kind == "move":
            self._set_filter_name(moved)
        sort = node.find(X + "sortState")
        if sort is not None:
            _move_sqref_attr(sort, "ref", change)

    def _set_filter_name(self, area: Area) -> None:
        """The filter's hidden name names the filter's range (a move: measured)."""
        from ..address import quote_sheet

        names = self.wb._root.find(X + "definedNames")
        index = str(self._sheet_index())
        for node in (names if names is not None else []):
            if node.get("name") == "_xlnm._FilterDatabase" and node.get("localSheetId") == index:
                first = f"${column_letters(area.c1)}${area.r1}"
                last = f"${column_letters(area.c2)}${area.r2}"
                node.text = f"{quote_sheet(self.sheet.stored_name)}!{first}:{last}"
                self.package.mark_dirty(self.wb.part)

    def _drop_filter_name(self) -> None:
        names = self.wb._root.find(X + "definedNames")
        index = self._sheet_index()
        for node in list(names if names is not None else []):
            if node.get("name") == "_xlnm._FilterDatabase" and node.get("localSheetId") \
                    == str(index):
                names.remove(node)
                self.package.mark_dirty(self.wb.part)

    def _sheet_index(self) -> int:
        sheets = [node for node in self.wb._root.iter(X + "sheet")]
        for index, node in enumerate(sheets):
            if node.get("name") == self.sheet.stored_name:
                return index
        return -1

    def _views(self, root: Any) -> None:
        change = self.change
        for view in root.iterfind(f"{X}sheetViews/{X}sheetView"):
            pane = view.find(X + "pane")
            if pane is not None and pane.get("state") in ("frozen", "frozenSplit"):
                name = "ySplit" if change.axis == ROW else "xSplit"
                split = _int(pane.get(name))
                if split:
                    frozen = _frozen(change, split)
                    if frozen != split:
                        if frozen:
                            pane.set(name, str(frozen))
                        else:
                            del pane.attrib[name]
                top = pane.get("topLeftCell")
                position = _position(top) if top else None
                if position is not None:
                    line = position[0] if change.axis == ROW else position[1]
                    new = _top_line(change, line)
                    if new != line:
                        r, c = (new, position[1]) if change.axis == ROW else (position[0], new)
                        pane.set("topLeftCell", f"{column_letters(c)}{r}")
            for selection in view.iterfind(X + "selection"):
                cell = selection.get("activeCell")
                position = _position(cell) if cell else None
                if position is not None:
                    line = position[0] if change.axis == ROW else position[1]
                    new = change.index(line)
                    new = new if new is not None else max(1, min(change.first,
                                                                 change.limit))
                    r, c = (new, position[1]) if change.axis == ROW else (position[0], new)
                    selection.set("activeCell", f"{column_letters(c)}{r}")
                if selection.get("sqref"):
                    _move_sqref_attr(selection, "sqref", change,
                                     fallback=selection.get("activeCell"))
            top = view.get("topLeftCell")
            position = _position(top) if top else None
            if position is not None:
                line = position[0] if change.axis == ROW else position[1]
                new = _top_line(change, line)
                r, c = (new, position[1]) if change.axis == ROW else (position[0], new)
                view.set("topLeftCell", f"{column_letters(c)}{r}")

    def _breaks(self, root: Any) -> None:
        change = self.change
        name = "rowBreaks" if change.axis == ROW else "colBreaks"
        breaks = root.find(X + name)
        if breaks is None:
            return
        for brk in list(breaks):
            line = _int(brk.get("id"))
            if line is None:
                continue
            new = change.index(line)
            if new is None:
                breaks.remove(brk)
            else:
                brk.set("id", str(new))
        nodes = sorted(breaks, key=lambda n: int(n.get("id", "0")))
        for node in nodes:
            breaks.remove(node)
            breaks.append(node)
        breaks.set("count", str(len(breaks)))
        breaks.set("manualBreakCount", str(sum(1 for n in breaks if _true(n.get("man")))))
        if not len(breaks):
            root.remove(breaks)

    def _sqrefs(self, root: Any) -> None:
        for tag in ("ignoredErrors/" + X + "ignoredError", "protectedRanges/" + X +
                    "protectedRange"):
            for node in list(root.iterfind(X + tag)):
                if not _move_sqref_attr(node, "sqref", self.change):
                    node.getparent().remove(node)

    def _tables_on_sheet(self) -> None:
        change = self.change
        for table in self.tables:
            root = table.root
            area = table.area
            if change.axis == COLUMN and change.kind == "delete":
                key = table.name.casefold()
                if key in self.gone.deleted and self.gone.deleted[key] is None:
                    self._remove_table(table)
                    continue
            if change.axis == ROW and change.kind == "delete" and table.totals:
                totals = area.r2 - table.totals + 1
                if change.first <= totals and area.r2 <= change.last:
                    _drop_totals(root)
            if change.kind == "move" and change.axis == ROW:
                pieces = area_cells(area, change)
                moved = _bbox(pieces)
            else:
                moved = map_area(area, change)
            if moved is None:
                self._remove_table(table)
                continue
            if change.axis == COLUMN:
                self._table_columns(table, moved)
            root.set("ref", moved.ref)
            filt = root.find(X + "autoFilter")
            if filt is not None:
                totals = int(root.get("totalsRowCount", "0") or 0)
                filt_area = Area(moved.r1, moved.c1, moved.r2 - totals, moved.c2)
                _shift_filter_columns(filt, area, moved, change)
                filt.set("ref", filt_area.ref)
            sort = root.find(X + "sortState")
            if sort is not None:
                _move_sqref_attr(sort, "ref", change)
            self.package.mark_dirty(table.part)

    def _fill_calculated(self, data: Any) -> None:
        """Rows inserted inside a table get its calculated columns' formula (measured);
        rows moved in keep their own cells."""
        change = self.change
        if change.kind != "insert" or change.axis != ROW:
            return
        for table in self.tables:
            r1, r2 = table.area.r1, table.area.r2
            if not (r1 < change.first <= r2 or (table.body and change.first == table.body[0])):
                continue
            if table.header and change.first <= r1:
                continue
            columns = table.root.find(X + "tableColumns")
            for offset, column in enumerate(columns if columns is not None else []):
                formula = column.find(X + "calculatedColumnFormula")
                if formula is None or not formula.text or column.find(
                        X + "calculatedColumnFormula") is None:
                    continue
                if any(t.kind == "ref" for t in tokens(formula.text)):
                    raise NotDone("not_measured", f"t:{table.name}",
                                  "a calculated column whose formula has cell references")
                for line in range(change.first, change.first + change.count):
                    row = _row_node(data, line)
                    if row is None:
                        row = etree.Element(X + "row")
                        row.set("r", str(line))
                        _insert_row(data, row, line)
                    column_number = table.area.c1 + offset
                    ref = f"{column_letters(column_number)}{line}"
                    cell = next((c for c in row.iterchildren(X + "c") if c.get("r") == ref),
                                None)
                    if cell is None:
                        cell = etree.Element(X + "c")
                        cell.set("r", ref)
                        _insert_cell(row, cell, column_number)
                    f = etree.SubElement(cell, X + "f")
                    f.text = formula.text
                    self.formulas += 1

    def _table_columns(self, table: _Table, moved: Area) -> None:
        change = self.change
        columns = table.root.find(X + "tableColumns")
        if columns is None:
            return
        nodes = columns.findall(X + "tableColumn")
        if change.kind == "delete":
            for offset, node in enumerate(nodes):
                if change.first <= table.area.c1 + offset <= change.last:
                    columns.remove(node)
                    self._count("table_columns")
        elif change.kind == "insert" and table.area.c1 < change.first <= table.area.c2:
            taken = {(n.get("name") or "").casefold() for n in nodes}
            ids = [int(n.get("id")) for n in nodes if (n.get("id") or "").isdigit()]
            at = change.first - table.area.c1
            new_names = []
            for offset in range(change.count):
                word = _column_word(self.wb)
                number = 1
                while f"{word}{number}".casefold() in taken:
                    number += 1
                name = f"{word}{number}"
                taken.add(name.casefold())
                node = etree.Element(X + "tableColumn")
                node.set("id", str(max(ids, default=0) + 1 + offset))
                node.set("name", name)
                nodes[at + offset - 1].addnext(node) if at + offset - 1 < len(nodes) \
                    else columns.append(node)
                nodes = columns.findall(X + "tableColumn")
                new_names.append(name)
            if table.header:
                self._write_headers(moved.r1, change.first, new_names)
        columns.set("count", str(len(columns.findall(X + "tableColumn"))))

    def _write_headers(self, row: int, first: int, names: list[str]) -> None:
        """The new table columns' header cells hold their names (inline text: no shared
        string is added in a structural step)."""
        data = self.root.find(X + "sheetData")
        node = _row_node(data, row)
        if node is None:
            node = etree.Element(X + "row")
            node.set("r", str(row))
            _insert_row(data, node, row)
        for offset, name in enumerate(names):
            column = first + offset
            ref = f"{column_letters(column)}{row}"
            cell = next((c for c in node.iterchildren(X + "c") if c.get("r") == ref), None)
            if cell is None:
                cell = etree.Element(X + "c")
                cell.set("r", ref)
                _insert_cell(node, cell, column)
            for child in list(cell):
                cell.remove(child)
            cell.set("t", "inlineStr")
            text = etree.SubElement(etree.SubElement(cell, X + "is"), X + "t")
            text.text = name

    def _remove_table(self, table: _Table) -> None:
        for rel in list(self.package.relationships(self.part).values()):
            if rel.target_part == table.part:
                parts = self.root.find(X + "tableParts")
                for node in list(parts if parts is not None else []):
                    if node.get(R_ID) == rel.id:
                        parts.remove(node)
                if parts is not None:
                    if len(parts):
                        parts.set("count", str(len(parts)))
                    else:
                        self.root.remove(parts)
                self.package.remove_relationship(self.part, rel.id)
        if self.package.has_part(table.part):
            self.package.remove_part(table.part)
        self._count("tables")

    def _comments(self) -> None:
        change = self.change
        for kind, part in self._related_parts():
            if kind == "comments":
                root = self.package.tree(part)
                comment_list = root.find(X + "commentList")
                for comment in list(comment_list if comment_list is not None else []):
                    position = _position(comment.get("ref", ""))
                    if position is None:
                        continue
                    new = _moved_cell(position, change)
                    if new is None:
                        comment_list.remove(comment)
                        self._count("comments")
                    else:
                        comment.set("ref", f"{column_letters(new[1])}{new[0]}")
                self.package.mark_dirty(part)
            elif kind == "threadedComments":
                root = self.package.tree(part)
                for comment in list(root):
                    position = _position(comment.get("ref", ""))
                    if position is None:
                        continue
                    new = _moved_cell(position, change)
                    if new is None:
                        root.remove(comment)
                    else:
                        comment.set("ref", f"{column_letters(new[1])}{new[0]}")
                self.package.mark_dirty(part)
            elif kind == "vmlDrawing":
                self._vml(part)

    def _vml(self, part: str) -> None:
        change = self.change
        root = self.package.tree(part)
        if root is None:
            return
        changed = False
        for data in root.iter(VX + "ClientData"):
            if data.get("ObjectType") != "Note":
                continue
            row, column = data.find(VX + "Row"), data.find(VX + "Column")
            if row is None or column is None:
                continue
            position = (int((row.text or "0").strip()) + 1, int((column.text or "0").strip())
                        + 1)
            new = _moved_cell(position, change)
            shape = data.getparent()
            if new is None:
                shape.getparent().remove(shape)
                changed = True
                continue
            row.text, column.text = str(new[0] - 1), str(new[1] - 1)
            anchor = data.find(VX + "Anchor")
            if anchor is not None and anchor.text:
                numbers = [n.strip() for n in anchor.text.split(",")]
                if len(numbers) == 8 and all(n.lstrip("-").isdigit() for n in numbers):
                    values = [int(n) for n in numbers]
                    places = (2, 6) if change.axis == ROW else (0, 4)
                    delta = (new[0] - position[0]) if change.axis == ROW \
                        else (new[1] - position[1])
                    for at in places:
                        values[at] = max(0, values[at] + delta)
                    anchor.text = ", ".join(str(v) for v in values)
            changed = True
        if changed:
            self.package.mark_dirty(part)

    def _drawings(self) -> None:
        change = self.change
        for kind, part in self._related_parts():
            if kind != "drawing":
                continue
            root = self.package.tree(part)
            changed = False
            for anchor in root:
                tag = etree.QName(anchor).localname if isinstance(anchor.tag, str) else ""
                if tag not in ("twoCellAnchor", "oneCellAnchor"):
                    continue
                if anchor.get("editAs") == "absolute":
                    continue
                marks = [anchor.find(XDR + "from")]
                if tag == "twoCellAnchor" and anchor.get("editAs") != "oneCell":
                    marks.append(anchor.find(XDR + "to"))
                name = "row" if change.axis == ROW else "col"
                positions = [int(m.find(XDR + name).text) + 1 for m in marks
                             if m is not None and m.find(XDR + name) is not None]
                if not positions:
                    continue
                if len(positions) == 2:
                    span = change.span(positions[0], positions[1]) if change.kind != "move" \
                        else (change.index(positions[0]), change.index(positions[1]))
                    if span is None:
                        continue
                    new = list(span)
                else:
                    moved = change.index(positions[0])
                    new = [moved if moved is not None else change.first]
                if tag == "twoCellAnchor" and anchor.get("editAs") == "oneCell":
                    to = anchor.find(XDR + "to")
                    if to is not None and to.find(XDR + name) is not None:
                        delta = new[0] - positions[0]
                        node = to.find(XDR + name)
                        node.text = str(int(node.text) + delta)
                for mark, value in zip(marks, new):
                    node = mark.find(XDR + name)
                    if int(node.text) != value - 1:
                        node.text = str(value - 1)
                        changed = True
            if changed:
                self.package.mark_dirty(part)

    def _pivot_locations(self) -> None:
        for rel in self.package.relationships(self.part).values():
            if rel.is_external or rel.type.rpartition("/")[2] != "pivotTable" \
                    or rel.target_part is None:
                continue
            root = self.package.tree(rel.target_part)
            location = root.find(X + "location")
            area = parse_area(location.get("ref", "")) if location is not None else None
            if area is None:
                continue
            moved = map_area(area, self.change)
            if moved is not None and moved.ref != area.ref:
                location.set("ref", moved.ref)
                self.package.mark_dirty(rel.target_part)

    def _related_parts(self) -> list[tuple[str, str]]:
        out = []
        for rel in self.package.relationships(self.part).values():
            if rel.is_external or rel.target_part is None \
                    or not self.package.has_part(rel.target_part):
                continue
            out.append((rel.type.rpartition("/")[2], rel.target_part))
        return out

    def _count(self, kind: str) -> None:
        self.removed[kind] = self.removed.get(kind, 0) + 1


# -- helpers -----------------------------------------------------------------------------------


def _column_word(workbook: Any) -> str:
    """What Excel names a new table column, in the workbook's language: ``Kolom`` in a
    Dutch Excel (measured, Excel for Mac 16.106), ``Column`` otherwise."""
    language = str(getattr(workbook.locale, "language", "en-US") or "en-US")
    return "Kolom" if language.lower().startswith("nl") else "Column"


def _cuts(change: Change, area: Area) -> bool:
    """Whether a change goes through a table's lines: inserted inside them (after its
    first), deleted from them, or moved from or into them."""
    a, b = area.along(change.axis)
    if change.kind == "insert":
        return a < change.first <= b
    overlap = not (change.last < a or b < change.first)
    if change.kind == "delete":
        return overlap
    return overlap or a < change.before <= b


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def _true(value: str | None) -> bool:
    return value in ("1", "true")


def _position(ref: str | None) -> tuple[int, int] | None:
    if not ref:
        return None
    letters = ref.rstrip("0123456789")
    digits = ref[len(letters):]
    if not letters or not digits or not letters.isalpha():
        return None
    return int(digits), column_index(letters.upper())


def _prefix_key(prefix: str) -> Any:
    from .grid import _prefix_key as key

    return key(prefix)


def _anchoring(body: str, axis: str) -> str:
    """Whether a reference's parts along ``axis`` are all ``"absolute"``, all
    ``"relative"`` or ``"mixed"`` (``"none"`` when it has none along the axis)."""
    seen = set()
    for piece in body.split(":"):
        part = _parse_part(piece)
        if part is None:
            continue
        if axis == ROW and part.row is not None:
            seen.add("absolute" if part.row_anchor else "relative")
        if axis == COLUMN and part.column is not None:
            seen.add("absolute" if part.column_anchor else "relative")
    if not seen:
        return "none"
    return seen.pop() if len(seen) == 1 else "mixed"


def _number_rows(rows: list[Any]) -> None:
    """Rows and cells without ``r`` given theirs (the format allows leaving it out)."""
    if rows and not rows[0].getparent().xpath(
            "boolean(x:row[not(@r)] | x:row/x:c[not(@r)])", namespaces={"x": X[1:-1]}):
        return
    number = 0
    for row in rows:
        given = row.get("r")
        number = int(given) if given else number + 1
        if not given:
            row.set("r", str(number))
        column = 0
        for cell in row.iterchildren(X + "c"):
            ref = cell.get("r")
            position = _position(ref) if ref else None
            column = position[1] if position else column + 1
            if not ref:
                cell.set("r", f"{column_letters(column)}{number}")


def _formula_cells(data: Any) -> tuple[list[_Cell], dict[str, list[_Cell]]]:
    """Every formula cell, with its text (a shared member's computed from its master),
    and the shared groups."""
    rows = list(data.iterchildren(X + "row"))
    _number_rows(rows)
    cells: list[_Cell] = []
    masters: dict[str, tuple[int, int, str]] = {}
    groups: dict[str, list[_Cell]] = {}
    pending: list[_Cell] = []
    for row in rows:
        for node in row.iterchildren(X + "c"):
            f = node.find(X + "f")
            if f is None:
                continue
            position = _position(node.get("r"))
            if position is None:
                continue
            cell = _Cell(node, position[0], position[1])
            if f.get("t") == "shared" and f.get("si") is not None:
                si = f.get("si")
                cell.group = si
                groups.setdefault(si, []).append(cell)
                if f.text:
                    masters[si] = (position[0], position[1], f.text)
                    cell.text = f.text
                else:
                    pending.append(cell)
            else:
                cell.text = f.text if f.text is not None else None
            cells.append(cell)
    for cell in pending:
        master = masters.get(cell.group or "")
        if master is None:
            cell.text = None
            continue
        try:
            cell.text = shift_formula(master[2], cell.row - master[0], cell.column - master[1])
        except ShiftError:
            raise NotDone("not_measured", f"{column_letters(cell.column)}{cell.row}",
                          "a shared formula whose member cannot be computed") from None
    return cells, groups


def _reshare(group: list[_Cell], _unused: Any) -> None:
    """A shared group after the change: kept as one when every member still reads as its
    first member's formula moved to it, else each member given its own formula."""
    members = sorted((c for c in group if c.text is not None
                      and c.node.find(X + "f") is not None), key=lambda c: (c.row, c.column))
    if not members:
        return
    master = members[0]
    alike = True
    for cell in members[1:]:
        try:
            if shift_formula(master.text, cell.row - master.row,
                             cell.column - master.column) != cell.text:
                alike = False
                break
        except ShiftError:
            alike = False
            break
    if alike and len(members) > 1:
        rows = [c.row for c in members]
        columns = [c.column for c in members]
        ref = Area(min(rows), min(columns), max(rows), max(columns)).ref
        for cell in members:
            f = cell.node.find(X + "f")
            f.set("t", "shared")
            f.set("si", members[0].group or "0")
            if cell is master:
                f.text = cell.text
                f.set("ref", ref)
            else:
                f.text = None
                if "ref" in f.attrib:
                    del f.attrib["ref"]
        return
    for cell in members:
        f = cell.node.find(X + "f")
        for name in ("t", "si", "ref"):
            if name in f.attrib:
                del f.attrib[name]
        f.text = cell.text


def _table_at(tables: list[tuple[Area, str]], row: int, column: int) -> str | None:
    for area, name in tables:
        if area.r1 <= row <= area.r2 and area.c1 <= column <= area.c2:
            return name
    return None


def _rule_formulas(root: Any) -> Iterable[Any]:
    for node in root.iter(X + "formula", X + "formula1", X + "formula2", XM + "f"):
        yield node


def _row_node(data: Any, number: int) -> Any:
    if data is None:
        return None
    for row in data.iterchildren(X + "row"):
        if row.get("r") == str(number):
            return row
    return None


def _col_node(root: Any, column: int) -> Any:
    cols = root.find(X + "cols")
    for col in (cols if cols is not None else []):
        if int(col.get("min", "0")) <= column <= int(col.get("max", "0")):
            return col
    return None


def _insert_row(data: Any, node: Any, number: int) -> None:
    for row in data.iterchildren(X + "row"):
        if int(row.get("r", "0")) > number:
            row.addprevious(node)
            return
    data.append(node)


def _insert_cell(row: Any, node: Any, column: int) -> None:
    for cell in row.iterchildren(X + "c"):
        position = _position(cell.get("r"))
        if position is not None and position[1] > column:
            cell.addprevious(node)
            return
    row.append(node)


def _inserted_cells(row: Any, neighbour: int, change: Change) -> list[tuple[int, Any]]:
    """New cells of an inserted column taking the style of the cell to their left."""
    out = []
    for cell in row.iterchildren(X + "c"):
        position = _position(cell.get("r"))
        if position is None or position[1] != neighbour or cell.get("s") in (None, "0"):
            continue
        for offset in range(change.count):
            node = etree.Element(X + "c")
            node.set("r", f"{column_letters(change.first + offset)}{position[0]}")
            node.set("s", cell.get("s"))
            out.append((change.first + offset, node))
    return out


def _col_pieces(change: Change, lo: int, hi: int) -> list[tuple[int, int]]:
    """A ``col`` range's columns after an insert or delete, as runs (an insert inside
    splits it: the new columns take their format from the column to their left)."""
    if change.kind == "insert":
        if lo < change.first <= hi:
            return [(lo, change.first - 1), (change.first + change.count,
                                              min(hi + change.count, change.limit))]
        span = change.span(lo, hi)
        return [] if span is None else [span]
    span = change.span(lo, hi)
    return [] if span is None else [span]


def _bbox(areas: list[Area]) -> Area:
    return Area(min(a.r1 for a in areas), min(a.c1 for a in areas),
                max(a.r2 for a in areas), max(a.c2 for a in areas))


def _moved_cell(position: tuple[int, int], change: Change) -> tuple[int, int] | None:
    line = position[0] if change.axis == ROW else position[1]
    new = change.index(line)
    if new is None:
        return None
    return (new, position[1]) if change.axis == ROW else (position[0], new)


def _move_sqref_attr(node: Any, name: str, change: Change, fallback: str | None = None
                     ) -> bool:
    areas = sqref_areas(node.get(name, ""))
    if areas is None:
        return True
    moved = [piece for area in areas for piece in area_cells(area, change)]
    if not moved:
        if fallback:
            node.set(name, fallback)
            return True
        return False
    node.set(name, sqref_text(moved))
    return True


def _shift_filter_columns(node: Any, before: Area, after: Area, change: Change) -> None:
    """A filter's column settings (``colId``: from the filter's first column) through a
    column change."""
    if change.axis != COLUMN:
        return
    for column in list(node.findall(X + "filterColumn")):
        index = _int(column.get("colId"))
        if index is None:
            continue
        new = change.index(before.c1 + index)
        if new is None or not after.c1 <= new <= after.c2:
            node.remove(column)
        else:
            column.set("colId", str(new - after.c1))


def _drop_totals(root: Any) -> None:
    """A table's totals row deleted: the table keeps none (measured)."""
    for name in ("totalsRowCount",):
        if name in root.attrib:
            del root.attrib[name]
    root.set("totalsRowShown", "0")
    for column in root.iter(X + "tableColumn"):
        for name in ("totalsRowLabel", "totalsRowFunction"):
            if name in column.attrib:
                del column.attrib[name]
        for child in column.findall(X + "totalsRowFormula"):
            column.remove(child)


def _frozen(change: Change, split: int) -> int:
    if change.kind == "move":
        if change.target <= split < change.first:
            return split + change.count      # measured: J moved before A
        if change.first <= split < change.target:
            return max(0, split - change.count)
        return split
    return _frozen_insert(change, split)


def _frozen_insert(change: Change, split: int) -> int:
    """How many lines stay frozen: lines inserted at or before the split's last line add
    to it (measured); a delete leaves the split as it was (measured: deleting the frozen
    column keeps one frozen); a move leaves it."""
    if change.kind == "insert":
        return split + change.count if change.first <= split else split
    return split


def _top_line(change: Change, line: int) -> int:
    if change.kind == "move":
        if change.target < line <= change.first:
            return line + change.count
        if change.first < line <= change.target:
            return line - change.count
        return line
    return _top_line_insert(change, line)


def _top_line_insert(change: Change, line: int) -> int:
    """The first line a pane shows: moved by lines inserted before it, not at it
    (measured); a delete or a move leaves it (measured for a delete)."""
    if change.kind == "insert":
        return line + change.count if change.first < line else line
    return line


def _anchor_span(anchor: Any, axis: str) -> tuple[int, int] | None:
    if not isinstance(anchor.tag, str):
        return None
    tag = etree.QName(anchor).localname
    if tag not in ("twoCellAnchor", "oneCellAnchor") or anchor.get("editAs") == "absolute":
        return None
    name = "row" if axis == ROW else "col"
    first = anchor.find(f"{XDR}from/{XDR}{name}")
    if first is None:
        return None
    last = anchor.find(f"{XDR}to/{XDR}{name}") if tag == "twoCellAnchor" else None
    a = int(first.text) + 1
    b = int(last.text) + 1 if last is not None else a
    return a, b


def _dimension(root: Any, data: Any) -> None:
    node = root.find(X + "dimension")
    if node is None:
        return
    rows, columns = [], []
    for row in data.iterchildren(X + "row"):
        first = next(row.iterchildren(X + "c"), None)
        if first is None:
            continue
        last = next(row.iterchildren(X + "c", reversed=True))
        for cell in (first, last):
            position = _position(cell.get("r"))
            if position is not None:
                rows.append(position[0])
                columns.append(position[1])
    if not rows:
        node.set("ref", "A1")
        return
    node.set("ref", Area(min(rows), min(columns), max(rows), max(columns)).ref)


# -- the operations ------------------------------------------------------------------------------

#: The operations, by name: (kind, axis).
OPERATIONS = {"insert_rows": ("insert", ROW), "insert_columns": ("insert", COLUMN),
              "delete_rows": ("delete", ROW), "delete_columns": ("delete", COLUMN),
              "move_rows": ("move", ROW), "move_columns": ("move", COLUMN)}


def line_span(value: Any, axis: str, what: str) -> tuple[int, int]:
    """A line or a run of lines as given: a row as ``5`` or ``"5"``, rows as ``"5:7"``; a
    column as ``"C"``, ``3`` or ``"C:E"`` (``$`` allowed).  :class:`EditError`
    ``bad_position`` naming what is wrong and what is taken."""
    from . import EditError

    limit = MAX_ROW if axis == ROW else MAX_COLUMN
    forms = ("a row number (5) or rows (\"5:7\")" if axis == ROW
             else "a column's letters (\"C\"), its number (3) or columns (\"C:E\")")
    if isinstance(value, bool):
        raise EditError("bad_position", f"{what}: {forms}")
    if isinstance(value, int):
        if not 1 <= value <= limit:
            raise EditError("bad_position", f"{what}: {value} is not on the sheet "
                            f"(1 to {limit}); {forms}")
        return value, value
    if not isinstance(value, str) or not value.strip():
        raise EditError("bad_position", f"{what}: {forms}")
    pieces = value.replace("$", "").strip().split(":")
    if len(pieces) > 2:
        raise EditError("bad_position", f"{what}: {value!r}; {forms}")
    found = []
    for piece in pieces:
        piece = piece.strip()
        if piece.isdigit():
            number = int(piece)
            if axis == COLUMN and piece:
                number = int(piece)
        elif axis == COLUMN and piece.isalpha() and len(piece) <= 3:
            number = column_index(piece.upper())
        else:
            raise EditError("bad_position", f"{what}: {value!r}; {forms}")
        if not 1 <= number <= limit:
            raise EditError("bad_position", f"{what}: {value!r} is not on the sheet; {forms}")
        found.append(number)
    first, last = min(found), max(found)
    return first, last


def restructure(workbook: Any, sheet: Any, operation: str, first: Any, count: Any = None,
                before: Any = None, *, revision: int | None = None) -> Restructured:
    """One operation (:data:`OPERATIONS`) on ``sheet``, as one step: see
    :mod:`.grid`.  :class:`~xlsx_agent.edit.EditError` before anything changes."""
    from . import EditError

    kind, axis = OPERATIONS[operation]
    noun = "rows" if axis == ROW else "columns"
    if revision is not None and revision != workbook.version:
        raise EditError("revision", f"the call was made for version {revision}; the "
                        f"workbook is at version {workbook.version}: read what changed, then "
                        "call again")
    try:
        found = workbook._cell_sheet(sheet)
    except (KeyError, ValueError) as error:
        raise EditError("not_a_worksheet", str(error).strip("'\"")) from None
    label = format_sheet(found.stored_name, found.id)
    if found.type != "worksheet" or found.part is None:
        raise EditError("not_a_worksheet", f"{label} holds no cells to move")
    if workbook.strict:
        raise EditError("strict")
    if found.protected:
        raise EditError("protected", addresses=[label])
    if workbook._allowed() is not None:
        raise EditError("not_writable", f"the host's write policy is set: {noun} are "
                        "inserted, deleted or moved only in a workbook without one",
                        addresses=[label])
    a, b = line_span(first, axis, "first" if kind != "move" else "source")
    if count is None:
        count = b - a + 1
    elif isinstance(count, bool) or not isinstance(count, int) or count < 1:
        raise EditError("bad_position", f"count: a whole number of {noun}, 1 or more")
    elif b != a and count != b - a + 1:
        raise EditError("bad_position", f"count: {count} {noun}, but {first!r} names "
                        f"{b - a + 1}: give one or the other")
    limit = MAX_ROW if axis == ROW else MAX_COLUMN
    if a + count - 1 > limit:
        raise EditError("bad_position", f"{count} {noun} from {a} go past the sheet's last "
                        f"({limit})")
    target = 0
    if kind == "move":
        if before is None:
            raise EditError("bad_position", f"before: the {noun[:-1]} the {noun} are placed "
                            "before")
        target, target_end = line_span(before, axis, "before")
        if target_end != target:
            raise EditError("bad_position", "before: one line, the one the moved "
                            f"{noun} are placed before")
    change = Change(kind, axis, a, count, target)
    shown = _shown(axis, a, a + count - 1)
    if kind == "move" and a <= target <= a + count:
        return Restructured(operation, label, a, count, target, changed=False,
                            notes=(f"{shown} are already before {_shown(axis, target, target)}"
                                   ": nothing moved",),
                            calculation=workbook.calculation_status)
    editor = workbook.editor
    with editor.batch() as edit:
        job = Job(workbook, found, change, operation)
        job.check()
        if job.problems:
            raise EditError(job.problems[0][0],
                            "; ".join(f"{where}: {detail}"
                                      for _, where, detail in job.problems[:10]),
                            problems=[(code, where) for code, where, _ in job.problems])
        try:
            job.apply()
        except NotDone as problem:
            raise EditError(problem.code, f"{problem.where}: {problem.detail}",
                            problems=[(problem.code, problem.where)]) from None
        _bookkeeping(workbook, editor, edit, found, job)
    workbook._forget()
    notes = list(job.notes)
    if job._keep_one is not None:
        notes.append(f"every data row of t:{job._keep_one.name} was deleted: Excel keeps "
                     "one, empty, and so does this")
    return Restructured(operation, label, a, count, target if kind == "move" else None,
                        formulas=job.formulas, broken=tuple(job.broken),
                        broken_count=job.broken_count, removed=dict(job.removed),
                        notes=tuple(notes), calculation=workbook.calculation_status)


def _shown(axis: str, a: int, b: int) -> str:
    if axis == ROW:
        return f"row {a}" if a == b else f"rows {a}:{b}"
    first, last = column_letters(a), column_letters(b)
    return f"column {first}" if a == b else f"columns {first}:{last}"


def _bookkeeping(workbook: Any, editor: Any, edit: Any, sheet: Any, job: Job) -> None:
    """After the rewrite: the cell writer read afresh, the shared strings' counts, the
    calculation chain dropped and a full calculation asked for, the cells written so far
    (what a save reads back) moved with their cells."""
    if editor._cells is not None:
        editor._cells.finish()
    editor._cells = None
    writer = editor.cells
    writer.sheet(sheet.stored_name)
    writer.drop_calc_chain()
    writer.request_full_calculation()
    writer.finish()
    change = job.change
    written = editor._written.get(sheet.part)
    if written:
        moved = set()
        for row, column in written:
            new = _moved_cell((row, column), change)
            if new is not None:
                moved.add(new)
        editor._written[sheet.part] = moved
    from .structure import _label

    label = _label(sheet)
    if label in edit.cells:
        edit.cells[label] = {new for new in (_moved_cell(p, change) for p in edit.cells[label])
                             if new is not None}
    edit.sheets.add(sheet.part)


__all__ = ["Job", "MAX_LISTED", "OPERATIONS", "Restructured", "line_span", "restructure"]
