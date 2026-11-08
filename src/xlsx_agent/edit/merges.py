"""Merged cells (production R56): ``wb.merge("Stock!A1:C2")``, ``wb.unmerge(...)``.

::

    wb.merge("Report!A1:F1")                     # refused while B1:F1 hold anything
    wb.merge("Report!A1:F1", drop=True)          # ... or their content dropped, as Excel does
    wb.unmerge("Report!A1:F1")                   # the area's cells keep their formats
    wb.unmerge("Report!A1", fill="anchor")       # every cell given the first cell's format

Each is one undoable step (or part of the open batch), refused before anything changes.

**What Excel does, measured** (Excel for Mac 16.106, 2026-11-08, its object model's
``merge`` and ``unmerge`` on a made-up sheet, saved and read):

* *Merge* gives every cell of the area the first cell's format (a cell in italics there
  took the first cell's bold and fill), keeps the first cell's value and drops the others'
  (``B10`` holding 5 and ``A11`` holding text were emptied).  When the first cell is empty
  and another holds a value, Excel moves the first value in reading order (row by row)
  into it (``E10:F10`` with only ``F10`` filled: ``E10`` held it after).  Here a merge
  whose other cells hold anything is refused (``merge_content``, every such cell listed)
  unless ``drop=True``, which does what Excel does; a formula that would have to move
  into the first cell is refused even then (moving it would move its references: write
  the first cell instead).
* *Unmerge* removes the area; every cell keeps the format it has -- which, for an area
  Excel merged, is the first cell's.  ``fill="anchor"`` gives every cell the first cell's
  format anyway, for an area another program merged with only its first cell formatted.
* Merging a range that overlaps a merged area made Excel merge the two into one larger
  area; here it is refused (``merge_overlap``, with the areas: merge their union, or
  unmerge first), and so is a range overlapping a table (Excel does not merge a table's
  cells) or an array formula.  A range already merged: ``changed`` false.

The host's write policy must cover every cell of the area; a protected sheet refuses it.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from lxml import etree
from ooxml_edit.sheets import Area, Cell, Formula

from ..address import AddressError, RangeAddress, format_cell, format_range, parse_address
from . import EditError
from .sheets import _X, find_sheet, placed

#: The most cells listed in a refusal or a result.
MAX_LISTED = 50


@dataclass(frozen=True)
class MergeChange:
    """What :func:`merge` or :func:`unmerge` did: the area(s) (``areas``), whether
    anything changed (``note`` says why not), the cells whose content a merge dropped
    (``dropped``), the cell whose value moved into the first one (``moved``), and how
    many cells were given the first cell's format (``styled``)."""

    action: str
    areas: tuple[str, ...]
    changed: bool = True
    dropped: tuple[str, ...] = ()
    moved: str | None = None
    styled: int = 0
    note: str | None = None

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"action": self.action, "areas": list(self.areas[:limit]),
                               "changed": self.changed}
        if self.dropped:
            out["dropped"] = list(self.dropped[:limit])
            if len(self.dropped) > limit:
                out["dropped_cut"] = len(self.dropped) - limit
        if self.moved:
            out["moved"] = self.moved
        if self.styled:
            out["styled"] = self.styled
        if self.note:
            out["note"] = self.note
        return out


def _target(workbook: Any, address: Any) -> tuple[Any, Area, str]:
    """The sheet and area an address names (``Stock!A1:C2``)."""
    try:
        parsed = parse_address(address) if isinstance(address, str) else address
    except (AddressError, ValueError) as error:
        raise EditError("merge_range", f"{error}: give a range with its sheet, as "
                        "Stock!A1:C2") from None
    if not isinstance(parsed, RangeAddress) or parsed.sheet is None:
        raise EditError("merge_range", "give a range with its sheet, as Stock!A1:C2")
    sheet_ref = parsed.sheet
    sheet = find_sheet(workbook, f"#{sheet_ref.id}" if sheet_ref.id is not None
                       else sheet_ref.name)
    if sheet.type != "worksheet" or sheet.part is None:
        raise EditError("not_a_worksheet", addresses=[str(address)])
    label = sheet.stored_name
    return sheet, parsed.area, label


def _guards(workbook: Any, sheet: Any, area: Area, label: str) -> Any:
    if workbook.strict:
        raise EditError("strict")
    if sheet.protected:
        raise EditError("protected", addresses=[format_range(label, area)])
    allowed = workbook._allowed()
    if allowed is not None:
        refused = [format_cell(label, c.row, c.column) for c in area.cells()
                   if not allowed.allows(sheet.part, c)]
        if refused:
            raise EditError("not_writable", f"{len(refused)} cell(s) of the area",
                            addresses=refused)
    writer = workbook.editor.cells.sheet(sheet.stored_name)
    if writer is None:
        raise EditError("not_a_worksheet", addresses=[format_range(label, area)])
    return writer


def _overlaps(a: Area, b: Area) -> bool:
    return not (a.last.row < b.first.row or b.last.row < a.first.row
                or a.last.column < b.first.column or b.last.column < a.first.column)


def merge(workbook: Any, address: Any, *, drop: bool = False) -> MergeChange:
    """Merge the range ``address`` (``Stock!A1:C2``) as Excel merges it (see the module)."""
    sheet, area, label = _target(workbook, address)
    shown = format_range(label, area)
    if area.first == area.last:
        return MergeChange("merge", (shown,), changed=False,
                           note="one cell: there is nothing to merge")
    writer = _guards(workbook, sheet, area, label)
    merges = list(workbook._layout(sheet, edit=True).merges)
    if any(m.first == area.first and m.last == area.last for m in merges):
        return MergeChange("merge", (shown,), changed=False, note="the range is merged")
    clash = [format_range(label, m) for m in merges if _overlaps(m, area)]
    if clash:
        union = _union([area] + [m for m in merges if _overlaps(m, area)])
        raise EditError("merge_overlap", f"merge their union, {format_range(label, union)}, "
                        "or unmerge first", addresses=clash)
    tables = [t for t in writer.tables() if t.area is not None and _overlaps(t.area, area)]
    if tables:
        raise EditError("merge_table", "Excel does not merge a table's cells",
                        addresses=[format_range(label, t.area) for t in tables])
    arrays = [format_range(label, a) for a in workbook._layout(sheet, edit=True).arrays
              if _overlaps(a, area)]
    if arrays:
        raise EditError("merge_array", addresses=arrays)
    anchor = area.first
    held = [c for c in area.cells() if c != anchor and _holds(writer.content(c))]
    dropped = [format_cell(label, c.row, c.column) for c in held]
    moved: Cell | None = None
    if held and not drop:
        raise EditError("merge_content", f"{len(held)} cell(s); a merge keeps only the first "
                        "cell's: pass drop=True to drop theirs, as Excel does, or move them "
                        "first", addresses=dropped)
    if held and not _holds(writer.content(anchor)):
        moved = held[0]
        if writer.content(moved).formula is not None:
            raise EditError("merge_content", "the first cell is empty and the first value "
                            "is a formula, which Excel would move into it: write the first "
                            "cell instead", addresses=[dropped[0]])
    editor = workbook.editor
    style = writer.style(anchor)
    styled = 0
    with editor.batch():
        if moved is not None:
            writer.move(moved, anchor)
            editor.touched(sheet, [moved, anchor])
            dropped.pop(0)
            if writer.style(anchor) != style:
                _restyle(editor, sheet, writer, anchor, style)
        for cell in area.cells():
            if cell == anchor:
                continue
            if editor.write(sheet, cell, None, style=style):
                styled += 1
        tree = workbook.package.tree(sheet.part)
        holder = tree.find(_X + "mergeCells")
        if holder is None:
            holder = etree.Element(_X + "mergeCells")
            placed(tree, holder)
        node = etree.SubElement(holder, _X + "mergeCell")
        node.set("ref", area.ref)
        holder.set("count", str(sum(1 for n in holder if n.tag == _X + "mergeCell")))
        workbook.package.mark_dirty(sheet.part)
        workbook._forget()
    return MergeChange("merge", (shown,), dropped=tuple(dropped),
                       moved=format_cell(label, moved.row, moved.column) if moved else None,
                       styled=styled)


def unmerge(workbook: Any, address: Any, *, fill: str | None = None) -> MergeChange:
    """Unmerge every merged area the range (or cell) ``address`` overlaps (see the
    module); ``fill="anchor"`` gives each area's cells its first cell's format.  None
    there: ``changed`` false."""
    if fill not in (None, "anchor"):
        raise EditError("bad_value", 'fill is "anchor" or None')
    sheet, area, label = _target(workbook, address)
    found = [m for m in workbook._layout(sheet, edit=True).merges if _overlaps(m, area)]
    if not found:
        return MergeChange("unmerge", (format_range(label, area),), changed=False,
                           note="no merged area there")
    covered = _union(found)
    writer = _guards(workbook, sheet, covered, label)
    editor = workbook.editor
    styled = 0
    with editor.batch():
        tree = workbook.package.tree(sheet.part)
        holder = tree.find(_X + "mergeCells")
        refs = {m.ref for m in found}
        if holder is not None:
            for node in list(holder):
                ref = node.get("ref") or ""
                try:
                    parsed = Area.parse(ref, lenient=True)
                except ValueError:
                    continue
                if parsed.ref in refs:
                    holder.remove(node)
            remaining = sum(1 for n in holder if n.tag == _X + "mergeCell")
            if remaining:
                holder.set("count", str(remaining))
            else:
                tree.remove(holder)
        workbook.package.mark_dirty(sheet.part)
        if fill == "anchor":
            for merged in found:
                style = writer.style(merged.first)
                for cell in merged.cells():
                    if cell != merged.first and writer.style(cell) != style:
                        if _restyle(editor, sheet, writer, cell, style):
                            styled += 1
        workbook._forget()
    return MergeChange("unmerge", tuple(format_range(label, m) for m in found),
                       styled=styled)


def _holds(content: Any) -> bool:
    return content.kind != "blank" or content.formula is not None


def _restyle(editor: Any, sheet: Any, writer: Any, cell: Cell, style: int) -> bool:
    """Give ``cell`` the format ``style``, its content kept; whether it changed (a shared
    or array formula's cell, rich text and an error are left as they are)."""
    content = writer.content(cell)
    if content.formula is not None:
        if content.formula_type != "normal":
            return False
        value: Any = Formula(content.formula)
    elif content.kind == "blank":
        value = None
    elif content.kind == "text":
        if writer.runs(cell) is not None:
            return False
        value = content.value
    elif content.kind in ("number", "bool"):
        value = content.value
    else:
        return False
    return editor.write(sheet, cell, value, style=style)


def _union(areas: list[Area]) -> Area:
    return Area(Cell(min(a.first.row for a in areas), min(a.first.column for a in areas)),
                Cell(max(a.last.row for a in areas), max(a.last.column for a in areas)))


__all__ = ["MergeChange", "merge", "unmerge"]
