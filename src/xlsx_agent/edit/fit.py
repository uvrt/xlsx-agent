"""Column widths fitted to their cells, as Excel's AutoFit fits them (measured; the model and
its error are in :mod:`xlsx_agent.edit.layout`).

Each cell is read as it is stored (its value, or a formula's cached value), shown as Excel
shows it (text as it is; a number through its format, :func:`xlsx_agent.shown.shown_text`),
and measured in its own font.  What is not measured is listed by reason and leaves the
width to the cells that are.
"""

from __future__ import annotations

from typing import Any

from ..address import format_cell
from .format import Problems, Target
from .layout import (MEASURED_FONTS, Face, digit_width, face_of, fitted_width, measurable,
                     text_pixels)

#: The most cells one fit reads.
MAX_FIT_CELLS = 2_000_000
#: The most cells listed for each reason a cell is skipped.
MAX_SKIPPED = 20

REASONS = {
    "merged": "inside a merged area",
    "wrapped": "wrapped, indented, rotated or shrunk text, whose fitted width is not measured",
    "line_break": "text with a line break",
    "not_measured": "a value whose shown text is not measured (a date, a boolean, an "
                    "error, or a number format not measured)",
    "no_cache": "a formula without a cached value",
}


def fit_columns(workbook: Any, found: list[Target], digit: int,
                problems: Problems) -> tuple[dict[tuple[str, int], float], dict[str, list[str]]]:
    """Each column's fitted width (by sheet part and column number) over the cells the
    targets cover, and the cells skipped by reason; an unmeasured font is a problem."""
    from .format import _label

    editor = workbook.editor
    styles = editor.cells.styles
    fonts: dict[int, Face] = {}
    shapes: dict[int, tuple[bool, str | None]] = {}
    widths: dict[tuple[str, int], float] = {}
    skipped: dict[str, list[str]] = {}
    unmeasured: dict[str, list[str]] = {}
    read = 0
    for item in found:
        writer = editor.cells.sheet(item.sheet.stored_name)
        label = _label(item.sheet)
        merged = workbook.merges(item.sheet)
        area = item.area
        for cell in writer.cells_in(area):
            read += 1
            if read > MAX_FIT_CELLS:
                problems.add("too_many_cells", item.text, f"more than {MAX_FIT_CELLS:,} cells to "
                             "fit; fit fewer columns at a time")
                return {}, {}
            content = writer.content(cell)
            if content.kind == "blank":
                continue
            address = format_cell(label, cell.row, cell.column)

            def skip(reason: str) -> None:
                listed = skipped.setdefault(reason, [])
                if len(listed) < MAX_SKIPPED:
                    listed.append(address)

            if any(m.contains(cell) and m.first != m.last for m in merged):
                skip("merged")
                continue
            style = content.style
            if style not in shapes:
                shapes[style] = _shape(workbook, styles, style)
            plain, code = shapes[style]
            if not plain:
                skip("wrapped")
                continue
            if content.formula is not None and content.kind == "blank":
                skip("no_cache")
                continue
            if style not in fonts:
                fonts[style] = face_of(styles.font(style))
            face = fonts[style]
            if not measurable(face.name):
                unmeasured.setdefault(face.name, []).append(address)
                continue
            if content.kind == "text":
                text = str(content.value)
                if "\n" in text or "\r" in text:
                    skip("line_break")
                    continue
                pixels = text_pixels(face, text, digit)
            elif content.kind == "number":
                from ..shown import shown_text

                shown = shown_text(content.value, code, workbook.locale)
                if shown is None:
                    skip("not_measured")
                    continue
                # a number: each character as wide as the cell font's widest digit (measured)
                pixels = len(shown) * digit_width(face)
                pixels += max(0, pixels // (5 * digit) - 2)
            else:
                skip("not_measured")
                continue
            key = (item.sheet.part, cell.column)
            widths[key] = max(widths.get(key, 0.0), fitted_width(pixels, digit))
    for name, cells in unmeasured.items():
        problems.add("unmeasured_font", cells[0], f"{len(cells)} cell(s) in {name}, whose "
                     "widths are not measured; fonts that are: " + ", ".join(MEASURED_FONTS)
                     + ". Give these columns a width instead")
    return widths, skipped


def _shape(workbook: Any, styles: Any, style: int) -> tuple[bool, str | None]:
    """Whether a cell format shows its text on one plain line (no wrap, indent, rotation or
    shrink), and its number format's code."""
    from lxml import etree

    root = styles.package.tree(styles.part)
    xfs = root.find("{%s}cellXfs" % root.nsmap.get(None, ""))
    plain = True
    if xfs is not None and 0 <= style < len(xfs):
        for child in xfs[style]:
            if isinstance(child.tag, str) and etree.QName(child).localname == "alignment":
                for name in ("wrapText", "shrinkToFit"):
                    if (child.get(name) or "").strip().lower() in ("1", "true"):
                        plain = False
                if (child.get("indent") or "0") != "0" or \
                        (child.get("textRotation") or "0") != "0":
                    plain = False
    return plain, workbook.styles.format(style).code


__all__ = ["REASONS", "fit_columns"]
