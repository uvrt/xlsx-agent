"""Moving a table's columns among themselves (round 46, R75): what Excel does when a table
column is dragged to another place in its table -- Cut of the table column's cells and
"Insert Cut Cells" before another of its columns -- as one step of
:func:`~xlsx_agent.edit.rearrange.rearrange_columns`.

**What moves, measured** (Excel for Mac 16.106, 2026-11-09; ``tests/table_columns_book.py``
and ``tests/test_table_columns.py``, oracle cases there):

* *Only the table's rows.*  The cells of the table's rows (header, data, totals) in the
  columns moved, and in the columns they pass, move as whole columns would; the sheet's
  columns do not: their widths and styles, and every cell above and below the table, stay.
* *The table* keeps its name, range, style and totals row; its ``tableColumn`` entries are
  reordered with their cells, each keeping its id, name, calculated formula and totals
  function or label.  The filter's criteria follow their columns (``colId`` renumbered),
  except a moved column's own, which Excel drops; so is a sort key on a moved column (and
  the sort state with its last key), the others changing as references do.
* *References* (in cells on every sheet, defined names, chart series, table formulas, rule
  formulas, pivot cache sources): one whose rows all lie in the table's rows changes as for
  a move of whole columns (:class:`~.grid.Rewriter`, ``band``); a whole column, or a cell
  or range reaching past the table's rows (``D2:D12``, ``C2:E9``), stays as it is -- but a
  range in one column, running from the table's rows out of one end of them, loses its rows
  in the table when it is in the column the cut cells were put before, or the one they were
  taken from (Excel numbers that one after inserting them: one further right for a move to
  the left): ``D5:D12`` is ``D11:D12``, ``D1:D4`` is ``D1:D2`` (listed in ``reshaped``).
  Structured references name columns, so they stay.  Measured for one column moved at a
  time, which is how :func:`~.rearrange.rearrange_columns` moves a table's columns.
* *Rules, hyperlinks, comments* follow their cells: the part of a data validation or a
  conditional format in the table's rows moves (each part's formula rewritten for its
  first cell, the others reading relative to it, as :func:`~.grid.anchored`), the rest
  stays.  A moved column also takes the
  data validations and conditional formats of the table column it now follows when they
  cover that column's data rows exactly, and a data format (``dataDxfId``): its cells' own
  number format, else that column's (as Excel gives a table column it inserts).

**Refused, as Excel refuses**: the block moved is not whole table columns (Excel refuses
a header or data cells alone, or cells from outside, with a parameter error), or the place
is outside the table.  **Refused, not measured yet**: an array formula, a merged area, a
pivot table, a sparkline, a protected range or an ignored error in the rows and columns
the move reaches; a hyperlink across the table's edge; a moved column with filter
criteria while the filter hides rows; parts of a sheet :mod:`.restructure` does not
rewrite.
"""

from __future__ import annotations

import copy
from typing import Any

from lxml import etree
from ooxml_edit.sheets import column_letters

from ..address import format_sheet
from ..formula import ShiftError, shift as shift_formula
from .grid import (COLUMN, X, X14, XM, VX, Area, Change, NotDone, Rewriter, _grouped, _merged,
                   anchored, area_cells, parse_area, sqref_areas, sqref_text)
from .restructure import (_EXT_HANDLED, _UNHANDLED, Job, Restructured, _Table, _position,
                          _reshare)


class TableColumnsJob(Job):
    """A move of whole table columns (``change``: columns, ``first`` .. ``last`` placed
    before ``before``) within the table of ``table_part``: only the table's rows move."""

    def __init__(self, workbook: Any, sheet: Any, change: Change, table_part: str) -> None:
        super().__init__(workbook, sheet, change, "move_table_columns")
        table = next((t for t in self.tables if t.part == table_part), None)
        if table is None:
            raise NotDone("not_measured", self.label, "the table is not on the sheet")
        self.table: _Table = table
        self.band = (table.area.r1, table.area.r2)
        c = self.change
        # where the cut cells are taken out and put in, as Excel numbers them: it inserts
        # first, so a block moved left is taken out from one column further right
        out = c.first if c.before > c.first else c.last + 1
        trim = frozenset(range(c.before, c.before + c.count)) | \
            frozenset(range(out, out + c.count))
        self.rewriter = Rewriter(c, sheet.stored_name, self.gone, band=self.band, trim=trim)
        #: The columns the move reaches: the block and the columns it passes.
        self.reach = (min(c.first, c.before), max(c.last, c.before - 1))

    # -- checks ------------------------------------------------------------------------------

    def _rect(self) -> Area:
        return Area(self.band[0], self.reach[0], self.band[1], self.reach[1])

    def _meets(self, area: Area) -> bool:
        rect = self._rect()
        return not (area.r2 < rect.r1 or rect.r2 < area.r1 or area.c2 < rect.c1
                    or rect.c2 < area.c1)

    def check(self) -> None:
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
        change, table, area = self.change, self.table, self.table.area
        where = f"t:{table.name}"
        if not (area.c1 <= change.first and change.last <= area.c2
                and area.c1 <= change.before <= area.c2 + 1):
            self.refuse("table_move", where, f"t:{table.name} holds "
                        f"{column_letters(area.c1)}:{column_letters(area.c2)}: its columns "
                        "are moved among themselves")
            return
        for f in data.iter(X + "f"):
            if f.get("t") not in ("array", "dataTable"):
                continue
            found = parse_area(f.get("ref") or f.getparent().get("r") or "")
            if found is not None and self._meets(found):
                self.refuse("array_split" if f.get("t") == "array" else "not_measured",
                            self._where(found), "an array formula or a data table in the "
                            "table's rows the move reaches")
        merges = root.find(X + "mergeCells")
        for merge in (merges if merges is not None else []):
            found = parse_area(merge.get("ref", ""))
            if found is not None and self._meets(found):
                self.refuse("merge_split", self._where(found), "a merged area in the "
                            "table's rows the move reaches")
        for rel in self.package.relationships(self.part).values():
            if rel.is_external or rel.type.rpartition("/")[2] != "pivotTable" \
                    or rel.target_part is None or not self.package.has_part(rel.target_part):
                continue
            location = self.package.tree(rel.target_part).find(X + "location")
            found = parse_area(location.get("ref", "")) if location is not None else None
            if found is not None and self._meets(found):
                self.refuse("pivot_table", self._where(found),
                            "a pivot table in the table's rows the move reaches")
        for tag in ("ignoredErrors/" + X + "ignoredError",
                    "protectedRanges/" + X + "protectedRange"):
            for node in root.iterfind(X + tag):
                if any(self._meets(a) for a in sqref_areas(node.get("sqref", "")) or []):
                    self.refuse("not_measured", self.label, "a protected range or an ignored "
                                "error in the table's rows the move reaches")
        for spark in root.iter("{http://schemas.microsoft.com/office/spreadsheetml/2009/9/"
                               "main}sparkline"):
            sq = spark.find(XM + "sqref")
            found = parse_area(sq.text or "") if sq is not None else None
            if found is not None and self._meets(found):
                self.refuse("not_measured", self.label, "a sparkline in the table's rows "
                            "the move reaches")
        links = root.find(X + "hyperlinks")
        for link in (links if links is not None else []):
            found = parse_area(link.get("ref", ""))
            if found is not None and self._meets(found) and not (
                    self.band[0] <= found.r1 and found.r2 <= self.band[1]):
                self.refuse("not_measured", self._where(found),
                            "a hyperlink across the table's edge")
        self._check_filter(data)
        self._check_names()

    def _check_filter(self, data: Any) -> None:
        """A moved column's criteria go (measured); while the filter hides rows, what Excel
        then shows is not measured."""
        moved = self._filter_columns(moved=True)
        if not moved:
            return
        body = self.table.body
        if body is None:
            return
        for row in data.iterchildren(X + "row"):
            number = int(row.get("r") or 0)
            if body[0] <= number <= body[1] and row.get("hidden") in ("1", "true"):
                self.refuse("not_measured", f"t:{self.table.name}", "a column with filter "
                            "criteria moved while the filter hides rows: clear the filter "
                            "first")
                return

    def _filter_columns(self, moved: bool) -> list[Any]:
        node = self.table.root.find(X + "autoFilter")
        out = []
        for column in (node.findall(X + "filterColumn") if node is not None else []):
            raw = column.get("colId", "")
            if not raw.isdigit():
                continue
            at = self.table.area.c1 + int(raw)
            if (self.change.first <= at <= self.change.last) == moved:
                out.append(column)
        return out

    # -- applying ----------------------------------------------------------------------------

    def apply(self) -> None:
        for sheet in self._sheets:
            self._formulas_of(sheet)
        self._names()
        self._charts()
        self._pivot_caches()
        self._other_tables()
        self._sheet()

    def _pivot_caches(self) -> None:
        for part in self.package.part_names:
            if not part.startswith("xl/pivotCache/pivotCacheDefinition"):
                continue
            root = self.package.tree(part)
            changed = False
            for source in root.iter(X + "worksheetSource"):
                if source.get("sheet") is None or source.get("ref") is None or \
                        self.rewriter.target != _key(source.get("sheet")):
                    continue
                new = self.rewriter.body(source.get("ref"))
                if new != source.get("ref") and "#" not in new:
                    source.set("ref", new)
                    changed = True
                    self.formulas += 1
            if changed:
                self.package.mark_dirty(part)

    def _sheet(self) -> None:
        root = self.root
        data = root.find(X + "sheetData")
        cells, groups = getattr(self, "_target_cells", ([], {}))
        self._cells(data)
        for cell in cells:
            position = _position(cell.node.get("r"))
            if position is not None:
                cell.row, cell.column = position
        for group in groups.values():
            _reshare(group, None)
        self._band_rules(root)
        self._hyperlinks(root)
        self._comments()
        self._table()
        self.package.mark_dirty(self.part)

    def _moved(self, row: int, column: int) -> tuple[int, int]:
        if self.band[0] <= row <= self.band[1]:
            new = self.change.index(column)
            assert new is not None
            return row, new
        return row, column

    def _cells(self, data: Any) -> None:
        lo, hi = self.reach
        for row in data.iterchildren(X + "row"):
            number = int(row.get("r") or 0)
            if not self.band[0] <= number <= self.band[1]:
                continue
            placed = []
            moved = False
            for cell in row.iterchildren(X + "c"):
                position = _position(cell.get("r"))
                if position is None:
                    continue
                column = position[1]
                if lo <= column <= hi:
                    new = self.change.index(column)
                    assert new is not None
                    if new != column:
                        cell.set("r", f"{column_letters(new)}{number}")
                        moved = True
                        column = new
                placed.append((column, cell))
            if not moved:
                continue
            placed.sort(key=lambda item: item[0])
            tail = [child for child in row if child.tag != X + "c"]
            for _, cell in placed:
                row.append(cell)
            for child in tail:
                row.append(child)

    # rules -------------------------------------------------------------------------------------

    def _band_groups(self, areas: list[Area], flat: list[str]
                     ) -> list[tuple[list[Area], list[str]]]:
        """A rule's areas and formulas (written for the areas' top-left cell): the parts in
        the table's rows the move reaches go with their cells, the rest stay; each group's
        formulas written for its own top-left cell."""
        top = min(a.r1 for a in areas), min(a.c1 for a in areas)
        pieces: list[tuple[Area, list[str]]] = []
        try:
            for area in areas:
                for part, inside in self._split(area):
                    texts = [shift_formula(f, part.r1 - top[0], part.c1 - top[1]) if f else f
                             for f in flat]
                    if inside:
                        for moved, written in anchored([part], texts, self.change,
                                                       self.rewriter, self.target):
                            first = min(a.r1 for a in moved), min(a.c1 for a in moved)
                            for piece in moved:
                                pieces.append((piece, [
                                    shift_formula(t, piece.r1 - first[0], piece.c1 - first[1])
                                    if t else t for t in written]))
                    else:
                        pieces.append((part, [self.rewriter.formula(t, self.target) if t
                                              else t for t in texts]))
        except ShiftError:
            raise NotDone("not_measured", self.label,
                          "a rule's formula that cannot be read or moved") from None
        groups = _grouped(pieces)
        self.formulas += sum(1 for _, texts in groups for a, b in zip(texts, flat) if a != b)
        return groups

    def _split(self, area: Area) -> list[tuple[Area, bool]]:
        """``area`` as the parts above, in and below the table's rows (in: ``True`` when
        it also reaches the moved columns), then left of, in and right of those: a rule's
        formula may name moved cells from any of them, row by row."""
        r1, r2 = self.band
        lo, hi = self.reach
        out = []
        for a, b, rows_in in ((area.r1, min(area.r2, r1 - 1), False),
                              (max(area.r1, r1), min(area.r2, r2), True),
                              (max(area.r1, r2 + 1), area.r2, False)):
            if a > b:
                continue
            if not rows_in:
                out.append((Area(a, area.c1, b, area.c2), False))
                continue
            for c, d, columns_in in ((area.c1, min(area.c2, lo - 1), False),
                                     (max(area.c1, lo), min(area.c2, hi), True),
                                     (max(area.c1, hi + 1), area.c2, False)):
                if c <= d:
                    out.append((Area(a, c, b, d), columns_in))
        return out

    def _neighbour_areas(self) -> tuple[int, Area] | None:
        """The table column the moved block now follows, and the block's data rows: a rule
        covering exactly that column's data rows covers them too (measured)."""
        body = self.table.body
        change = self.change
        neighbour = change.target - 1
        if body is None or neighbour < self.table.area.c1:
            return None
        return neighbour, Area(body[0], change.target, body[1], change.target + change.count - 1)

    def _grow(self, groups: list[tuple[list[Area], Any]]) -> list[tuple[list[Area], Any]]:
        found = self._neighbour_areas()
        if found is None:
            return groups
        neighbour, block = found
        body = self.table.body
        assert body is not None
        for areas, _ in groups:
            rows: set[int] = set()
            for area in areas:
                if area.c1 <= neighbour <= area.c2:
                    rows.update(range(max(area.r1, self.band[0]),
                                      min(area.r2, self.band[1]) + 1))
            if rows == set(range(body[0], body[1] + 1)):
                areas.append(block)
                areas[:] = _merged(areas)
                break
        return groups

    def _band_rules(self, root: Any) -> None:
        priorities = [int(r.get("priority")) for r in root.iter(X + "cfRule")
                      if (r.get("priority") or "").isdigit()]
        next_priority = [max(priorities, default=0) + 1]
        for block in list(root.iterfind(X + "conditionalFormatting")):
            areas = self._areas(block.get("sqref", ""))
            rules = block.findall(X + "cfRule")
            formulas = [[f.text or "" for f in rule.findall(X + "formula")] for rule in rules]
            groups = self._grow(self._regrouped(areas, formulas))
            self._place_groups(block, groups, rules, "sqref", X + "formula", next_priority)
        for block in list(root.iter(X14 + "conditionalFormatting")):
            sq = block.find(XM + "sqref")
            areas = self._areas(sq.text if sq is not None else "")
            rules = block.findall(X14 + "cfRule")
            formulas = [[f.text or "" for f in rule.iter(XM + "f")] for rule in rules]
            groups = self._grow(self._regrouped(areas, formulas))
            self._place_groups(block, groups, rules, None, XM + "f", next_priority)
        validations = root.find(X + "dataValidations")
        for dv in list(validations if validations is not None else []):
            areas = self._areas(dv.get("sqref", ""))
            nodes = [dv.find(X + "formula1"), dv.find(X + "formula2")]
            texts = [[n.text or "" if n is not None else "" for n in nodes]]
            for index, (group_areas, written) in enumerate(
                    self._grow(self._regrouped(areas, texts))):
                node = dv if index == 0 else copy.deepcopy(dv)
                if index:
                    dv.addnext(node)
                node.set("sqref", sqref_text(group_areas))
                for name, text in zip(("formula1", "formula2"), written[0]):
                    child = node.find(X + name)
                    if child is not None:
                        child.text = text
        if validations is not None:
            validations.set("count", str(len(validations)))
        for block in root.iter(X14 + "dataValidations"):
            for dv in list(block.findall(X14 + "dataValidation")):
                sq = dv.find(XM + "sqref")
                areas = self._areas(sq.text if sq is not None else "")
                nodes = list(dv.iter(XM + "f"))
                for index, (group_areas, written) in enumerate(self._grow(
                        self._regrouped(areas, [[n.text or "" for n in nodes]]))):
                    node = dv if index == 0 else copy.deepcopy(dv)
                    if index:
                        dv.addnext(node)
                    node.find(XM + "sqref").text = sqref_text(group_areas)
                    for child, text in zip(node.iter(XM + "f"), written[0]):
                        child.text = text
            block.set("count", str(len(block.findall(X14 + "dataValidation"))))
        for spark in root.iter("{http://schemas.microsoft.com/office/spreadsheetml/2009/9/"
                               "main}sparkline"):
            f = spark.find(XM + "f")
            if f is not None and f.text:
                f.text = self._rewrite(f.text, self.target, self.label)

    def _areas(self, text: str | None) -> list[Area]:
        areas = sqref_areas(text or "")
        if not areas:
            raise NotDone("not_measured", self.label, "a rule's range that cannot be read")
        return areas

    def _regrouped(self, areas: list[Area], formulas: list[list[str]]
                   ) -> list[tuple[list[Area], list[list[str]]]]:
        flat = [text for texts in formulas for text in texts]
        out = []
        for group_areas, texts in self._band_groups(areas, flat):
            nested, at = [], 0
            for original in formulas:
                nested.append(texts[at:at + len(original)])
                at += len(original)
            out.append((group_areas, nested))
        return out

    # hyperlinks, comments ------------------------------------------------------------------------

    def _hyperlinks(self, root: Any) -> None:
        links = root.find(X + "hyperlinks")
        for link in (links if links is not None else []):
            area = parse_area(link.get("ref", ""))
            if area is None or not self._meets(area):
                continue
            moved = area_cells(area, self.change)
            link.set("ref", Area(min(a.r1 for a in moved), min(a.c1 for a in moved),
                                 max(a.r2 for a in moved), max(a.c2 for a in moved)).ref)

    def _comments(self) -> None:
        for kind, part in self._related_parts():
            if kind in ("comments", "threadedComments"):
                root = self.package.tree(part)
                nodes = root.iter(X + "comment") if kind == "comments" else list(root)
                changed = False
                for comment in nodes:
                    position = _position(comment.get("ref", ""))
                    if position is None:
                        continue
                    new = self._moved(*position)
                    if new != position:
                        comment.set("ref", f"{column_letters(new[1])}{new[0]}")
                        changed = True
                if changed:
                    self.package.mark_dirty(part)
            elif kind == "vmlDrawing":
                self._vml_notes(part)

    def _vml_notes(self, part: str) -> None:
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
            position = (int((row.text or "0").strip()) + 1,
                        int((column.text or "0").strip()) + 1)
            new = self._moved(*position)
            if new == position:
                continue
            column.text = str(new[1] - 1)
            anchor = data.find(VX + "Anchor")
            if anchor is not None and anchor.text:
                numbers = [n.strip() for n in anchor.text.split(",")]
                if len(numbers) == 8 and all(n.lstrip("-").isdigit() for n in numbers):
                    values = [int(n) for n in numbers]
                    for at in (0, 4):
                        values[at] = max(0, values[at] + new[1] - position[1])
                    anchor.text = ", ".join(str(v) for v in values)
            changed = True
        if changed:
            self.package.mark_dirty(part)

    # the table -----------------------------------------------------------------------------------

    def _table(self) -> None:
        table, change = self.table, self.change
        root, c1 = table.root, table.area.c1
        columns = root.find(X + "tableColumns")
        nodes = columns.findall(X + "tableColumn") if columns is not None else []
        if columns is not None and len(nodes) == table.area.c2 - c1 + 1:
            order = []
            for place in range(c1, table.area.c2 + 1):
                origin = change.origin(place)
                assert origin is not None
                order.append(nodes[origin - c1])
            for node in nodes:
                columns.remove(node)
            for node in order:
                columns.append(node)
            self._data_format(order)
        dropped = self._filter_columns(moved=True)
        for node in dropped:
            node.getparent().remove(node)
        if dropped:
            self.notes.append(f"t:{table.name}: the filter criteria of the moved column were "
                              "dropped, as Excel drops them")
        for node in self._filter_columns(moved=False):
            at = change.index(c1 + int(node.get("colId")))
            assert at is not None
            node.set("colId", str(at - c1))
        filt = root.find(X + "autoFilter")
        if filt is not None:
            kept = sorted(filt.findall(X + "filterColumn"), key=lambda n: int(n.get("colId")))
            for node in kept:
                filt.remove(node)
            for index, node in enumerate(kept):
                filt.insert(index, node)
        for sort in list(root.iter(X + "sortState")):
            self._sort(sort)
        self.package.mark_dirty(table.part)

    def _sort(self, sort: Any) -> None:
        """A sort key on a moved column goes, and the sort state with its last key; the
        others, and the sorted range, change as references do (measured)."""
        change = self.change
        for condition in sort.findall(X + "sortCondition"):
            area = parse_area(condition.get("ref", ""))
            if area is None:
                continue
            if not (area.c2 < change.first or change.last < area.c1):
                sort.remove(condition)
                continue
            condition.set("ref", self.rewriter.body(area.ref))
        if sort.find(X + "sortCondition") is None:
            sort.getparent().remove(sort)
            self.notes.append(f"t:{self.table.name}: its sort state went with the moved "
                              "column it sorted by, as Excel drops it")
            return
        if sort.get("ref"):
            sort.set("ref", self.rewriter.body(sort.get("ref")))

    def _data_format(self, order: list[Any]) -> None:
        """A moved table column's data format (``dataDxfId``, what new rows take): the one
        number format its own data cells share, else the column's it now follows -- its
        ``dataDxfId`` or its cells' one format (measured: a General column moved after
        ``0.00`` ones takes ``0.00``, a ``0.00`` one keeps it after a General one or at the
        table's first place; one with a format of its own after another format is not
        measured)."""
        change = self.change
        c1 = self.table.area.c1
        neighbour = change.target - 1
        for column in range(change.target, change.target + change.count):
            node = order[column - c1]
            if node.get("dataDxfId") is not None:
                continue
            dxf = self._format_of(column)
            if dxf is None and neighbour >= c1:
                dxf = order[neighbour - c1].get("dataDxfId") or self._format_of(neighbour)
            if dxf is not None:
                node.set("dataDxfId", dxf)

    def _format_of(self, column: int) -> str | None:
        """A dxf holding the number format every data cell of ``column`` has (one the
        workbook defines, 164 on), added to the styles; ``None`` for none or several."""
        body = self.table.body
        if body is None:
            return None
        data = self.root.find(X + "sheetData")
        styles: set[str] = set()
        rows = {int(row.get("r") or 0): row for row in data.iterchildren(X + "row")}
        for number in range(body[0], body[1] + 1):
            row = rows.get(number)
            ref = f"{column_letters(column)}{number}"
            cell = None if row is None else next(
                (c for c in row.iterchildren(X + "c") if c.get("r") == ref), None)
            styles.add(cell.get("s", "0") if cell is not None else "0")
        if len(styles) != 1:
            return None
        from .format import StyleTable

        table = StyleTable(self.wb)
        xf = table.xf(int(styles.pop()))
        number = int(xf.get("numFmtId", "0") or 0)
        if number < 164:
            return None
        formats = table.root.find(X + "numFmts")
        code = next((n.get("formatCode") for n in (formats if formats is not None else [])
                     if n.get("numFmtId") == str(number)), None)
        if code is None:
            return None
        dxfs = table._container("dxfs")
        for index, known in enumerate(dxfs.findall(X + "dxf")):
            only = list(known)
            if len(only) == 1 and only[0].tag == X + "numFmt" \
                    and only[0].get("formatCode") == code:
                return str(index)
        node = etree.SubElement(dxfs, X + "dxf")
        fmt = etree.SubElement(node, X + "numFmt")
        fmt.set("numFmtId", str(number))
        fmt.set("formatCode", code)
        dxfs.set("count", str(len(dxfs.findall(X + "dxf"))))
        table.styles.package.mark_dirty(table.styles.part)
        return str(len(dxfs.findall(X + "dxf")) - 1)


def _key(name: str) -> str:
    from .grid import sheet_key

    return sheet_key(name)


def move_table_columns(workbook: Any, sheet: Any, table_part: str, first: int, count: int,
                       before: int) -> Restructured:
    """Columns ``first`` .. ``first + count - 1`` of the table in ``table_part`` placed
    before column ``before`` (within the table: ``before`` up to one past its last column),
    as one step (or part of the open batch): see the module.
    :class:`~xlsx_agent.edit.EditError` before anything changes."""
    from . import EditError
    from .restructure import _shown

    label = format_sheet(sheet.stored_name, sheet.id)
    if workbook.strict:
        raise EditError("strict")
    if sheet.protected:      # moves on a protected sheet are not measured (round 46, lane W)
        raise EditError("protected", "a table's columns are not moved on a protected sheet: "
                        "unprotect it first (wb.unprotect_sheet)", addresses=[label])
    if workbook._allowed() is not None:
        raise EditError("not_writable", "the host's write policy is set: columns are moved "
                        "only in a workbook without one", addresses=[label])
    change = Change("move", COLUMN, first, count, before)
    if first <= before <= first + count:
        shown = _shown(COLUMN, first, first + count - 1)
        return Restructured("move_table_columns", label, first, count, before, changed=False,
                            notes=(f"{shown} are already before "
                                   f"{_shown(COLUMN, before, before)}: nothing moved",),
                            calculation=workbook.calculation_status)
    editor = workbook.editor
    with editor.batch() as edit:
        try:
            job = TableColumnsJob(workbook, sheet, change, table_part)
        except NotDone as problem:
            raise EditError(problem.code, f"{problem.where}: {problem.detail}",
                            problems=[(problem.code, problem.where)]) from None
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
        _bookkeeping(workbook, editor, edit, sheet, job)
    workbook._forget()
    return Restructured("move_table_columns", label, first, count, before,
                        formulas=job.formulas, broken=tuple(job.broken),
                        broken_count=job.broken_count, removed=dict(job.removed),
                        notes=tuple(job.notes), calculation=workbook.calculation_status,
                        reshaped=tuple(job.reshaped), reshaped_count=job.reshaped_count)


def _bookkeeping(workbook: Any, editor: Any, edit: Any, sheet: Any,
                 job: TableColumnsJob) -> None:
    """As :func:`.restructure._bookkeeping`, the cells written so far moved with the
    table's rows only."""
    if editor._cells is not None:
        editor._cells.finish()
    editor._cells = None
    writer = editor.cells
    writer.sheet(sheet.stored_name)
    writer.drop_calc_chain()
    writer.request_full_calculation()
    writer.finish()
    written = editor._written.get(sheet.part)
    if written:
        editor._written[sheet.part] = {job._moved(*p) for p in written}
    from .structure import _label

    name = _label(sheet)
    if name in edit.cells:
        edit.cells[name] = {job._moved(*p) for p in edit.cells[name]}
    edit.sheets.add(sheet.part)
    edit.whole.add(sheet.part)


__all__ = ["TableColumnsJob", "move_table_columns"]
