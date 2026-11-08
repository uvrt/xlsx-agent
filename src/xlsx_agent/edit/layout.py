"""A sheet's layout from code (0.2, E7; production round 44, R53): column widths (given or
fitted to the content), row heights, freeze panes, gridlines and sheet protection.

::

    wb.set_column_widths("Stock!A:C", 14)          # characters, as Excel's Column Width
    wb.set_column_widths("Stock!A:H", "fit")       # fitted as Excel's AutoFit fits them
    wb.set_row_heights("Stock!1:1", 30)            # points; None: the row fits its text
    wb.freeze_panes("Stock", rows=1, columns=1)    # or at="B2"; nothing frozen: rows=0
    wb.set_gridlines("Stock", False)
    wb.protect_sheet("Stock", password="s3cret")   # or no password
    wb.unprotect_sheet("Stock", password="s3cret")

Each is one step (or part of the open batch), refused before anything changes, and changes
no value: the calculation status stays as it was.  Something already so is ``changed: 0``
with a note, never an error.

**Widths** are characters of the workbook's default font's widest digit, as Excel's Column
Width dialog takes them: ``width`` is stored as Excel stores it -- measured (Excel for Mac
16.106, 2026-11-08, Calibri 11, widest digit 6 px): 12 is stored ``12.83203125``, 20
``20.83203125``, 8.43 ``9.33203125``, 1 ``1.83203125``: the characters times the digit's
pixels, rounded, plus 5 px, over the digit's pixels, truncated to 1/256 (ECMA-376 Part 1,
18.3.1.13).  ``None`` takes a column's own width away.  The widest digit is the default
font's (the ``Normal`` style's) digits' advance widths at its size, each rounded to whole
pixels at 72 per inch (measured: 6 px for Calibri 11, Arial 10, Times New Roman 12 and
Aptos Narrow 11; 8 for Calibri 16).

**Fitted widths** (``"fit"``) are Excel's AutoFit, measured in the same session (88
columns of text and numbers in the default fonts Calibri 11, Arial 10, Times New Roman 12
and Calibri 16, and in eleven other fonts and sizes): a text's width in pixels is the sum
of its characters' advance widths at the font's size in points, each rounded to whole
pixels (the face's bold widths for bold); past five widest digits Excel adds one pixel per
five widest digits, less two (``max(0, T // (5 * digit) - 2)``); a number is as wide as its
shown text has characters, each the widest digit; then 7 px of padding, over the widest
digit, truncated to 1/256, written with ``bestFit``.  A column's width is its widest cell's.
On the probe this is exact for Calibri, Arial, Times New Roman, Courier New, Cambria and
Aptos at 10 to 16 points; at 8 points Excel's width was 2 px narrower and at 20 points 2 px
wider (0.33 of a character), and a word with accented letters up to 4 px off in Arial.
Fonts are measured with ooxml-common's advance widths
(:func:`ooxml_common.text.fontmap.metrics_for`): a face is measured when its table is its
own or a metric-compatible clone's (Calibri, Arial, Helvetica, Times New Roman, Courier
New, Cambria, Aptos, Aptos Display and the others :func:`measurable` accepts).  A cell in
another font refuses the fit (``unmeasured_font``, naming the font and the fonts that are
measured): Aptos Narrow, which has no table of its own, measured 15% narrower than Aptos's
widths on average (a 40-letter word 163 px against 200). Cells a fit does not measure are
listed apart (``skipped``) and leave the width to the others: merged cells, wrapped,
rotated or indented text, text with a line break, booleans and errors (shown in the
application's language), dates and other numbers whose shown text is not measured
(:func:`xlsx_agent.shown.shown_text`), and formulas without a cached value.  A column with
nothing measured keeps its width.

**Heights** are points (0 to 409), written with ``customHeight`` (measured: Excel's own 30
is ``ht="30" customHeight="1"``); ``None`` lets the row fit its text again.

**Freeze panes** as Excel writes them (measured, frozen at C4): ``pane xSplit="2"
ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"`` and a selection in
each pane (``C1`` top right, ``A4`` bottom left, ``C4`` bottom right); rows alone or
columns alone take the one pane they make (ECMA-376 Part 1, 18.3.1.66).  ``rows=0,
columns=0`` unfreezes.  **Gridlines**: ``showGridLines="0"`` hides them (measured).

**Protection** as Excel's Protect Sheet writes it (measured): ``sheetProtection sheet="1"
objects="1" scenarios="1"``; with a password, its SHA-512 hash with a random 16-byte salt
and 100,000 rounds (``algorithmName``, ``hashValue``, ``saltValue``, ``spinCount``; ECMA-376
Part 1, 18.3.1.85): the hash Excel wrote for a password is the one :func:`password_hash`
computes from its salt, and Excel unprotected a sheet protected here with its password
(measured).  Unprotecting a sheet with a password needs it (``wrong_password``); one
protected with the legacy 16-bit ``password`` is unprotected only with ``force=True`` (the
host's own code: that hash is not checked here).  Which cells stay open to input is each
cell's ``locked`` (``wb.format(..., protection={"locked": False})``).
"""

from __future__ import annotations

import base64
import hashlib
import math
import os
from dataclasses import dataclass
from typing import Any

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, SML_NS, Area, Cell, column_letters, number_text

from .format import Formatted, Problems, Target, guard, targets

_X = "{%s}" % SML_NS

#: CT_Worksheet's children in schema order (ECMA-376 Part 1, 18.3.1.99).
SHEET_ORDER = ("sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
               "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
               "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
               "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
               "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
               "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
               "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
               "webPublishItems", "tableParts", "extLst")

#: Excel's padding of a column width set by hand, and of a fitted one, in pixels (measured).
PADDING = 5
FIT_PADDING = 7
#: A sheet's base column width, in characters, when it gives none (ECMA-376 Part 1,
#: 18.3.1.81 ``baseColWidth``).
BASE_COLUMN_WIDTH = 8
#: The widest a column is (Excel's column width limit, in characters).
MAX_WIDTH = 255
#: What a password is hashed with, as Excel's Protect Sheet hashed it (measured).
SPIN_COUNT = 100_000


# -- the default font --------------------------------------------------------------------------


@dataclass(frozen=True)
class Face:
    """A cell's font as measured: name, size in points, bold."""

    name: str
    size: float
    bold: bool = False


def measurable(name: str) -> bool:
    """Whether ``name``'s advance widths are known: ooxml-common has a table of its own,
    or of a metric-compatible clone (Carlito for Calibri)."""
    from ooxml_common.text.fontmap import metrics_for, substitution_for

    if metrics_for(name) is None:
        return False
    found = substitution_for(name)
    if found is None:
        return True
    return found.metric_compatible or found.metrics.casefold() == found.office.casefold()


MEASURED_FONTS = ("Calibri", "Arial", "Helvetica", "Times New Roman", "Courier New",
                  "Cambria", "Aptos", "Aptos Display")


def _advances(face: Face, text: str) -> list[float]:
    """Each character's advance width in points (= pixels at 72 per inch)."""
    from ooxml_common.text.fontmap import metrics_for

    metrics = metrics_for(face.name)
    assert metrics is not None
    bold = face.bold and bool(metrics.bold_widths)
    table = metrics.bold_widths if bold else metrics.widths
    default = metrics.bold_default_width if bold else metrics.default_width
    scale = face.size / metrics.units_per_em
    return [table.get(ch, default) * scale for ch in text]


def digit_width(face: Face) -> int:
    """The widest digit's width in whole pixels (at least 1)."""
    return max(1, max(_round(width) for width in _advances(face, "0123456789")))


def _round(value: float) -> int:
    return int(math.floor(value + 0.5))


def text_pixels(face: Face, text: str, digit: int) -> int:
    """A text's width in pixels as Excel's AutoFit measures it (see the module)."""
    total = sum(_round(width) for width in _advances(face, text))
    return total + max(0, total // (5 * digit) - 2)


def fitted_width(pixels: int, digit: int) -> float:
    """A fitted column's width, in the file's unit, for content ``pixels`` wide."""
    return math.floor((pixels + FIT_PADDING) / digit * 256) / 256


def stored_width(characters: float, digit: int) -> float:
    """A width given in characters as Excel stores it (see the module)."""
    pixels = _round(characters * digit) + PADDING
    return math.floor(pixels / digit * 256) / 256


def default_face(workbook: Any) -> Face:
    """The workbook's default font: the ``Normal`` style's (``cellStyleXfs`` 0)."""
    styles = workbook.editor.cells.styles
    root = styles.package.tree(styles.part) if styles is not None else None
    font = None
    if root is not None:
        base = root.find(_X + "cellStyleXfs")
        number = 0
        if base is not None and len(base):
            raw = base[0].get("fontId", "0")
            number = int(raw) if raw.isdigit() else 0
        fonts = root.find(_X + "fonts")
        if fonts is not None and number < len(fonts):
            font = fonts[number]
    return face_of(font)


def face_of(font: Any) -> Face:
    """A ``font`` element as a :class:`Face` (Calibri 11 when it names none, Excel's
    default font of the ECMA-376 era)."""
    name, size, bold = "Calibri", 11.0, False
    if font is not None:
        for child in font:
            tag = etree.QName(child).localname if isinstance(child.tag, str) else ""
            if tag == "name" and child.get("val"):
                name = child.get("val")
            elif tag == "sz":
                try:
                    size = float(child.get("val", "11"))
                except ValueError:
                    pass
            elif tag == "b":
                bold = child.get("val", "1").strip().lower() not in ("0", "false")
    return Face(name, size, bold)


# -- columns -----------------------------------------------------------------------------------


class Columns:
    """A sheet's ``cols``, read once: each column's attributes, set by column and written
    back as runs of equal columns (ECMA-376 Part 1, 18.3.1.13)."""

    def __init__(self, workbook: Any, writer: Any) -> None:
        self.workbook = workbook
        self.writer = writer
        self.root = writer.workbook.package.tree(writer.part)
        self.runs: list[tuple[int, int, dict[str, str]]] = []
        node = self.root.find(_X + "cols")
        if node is not None:
            for col in node.iterchildren(_X + "col"):
                low, high = col.get("min", ""), col.get("max", "")
                if low.isdigit() and high.isdigit():
                    attrs = {k: v for k, v in col.attrib.items() if k not in ("min", "max")}
                    self.runs.append((int(low), int(high), attrs))
        self.changed = False

    def get(self, number: int) -> dict[str, str]:
        for low, high, attrs in self.runs:
            if low <= number <= high:
                return dict(attrs)
        return {}

    def default_width(self) -> float:
        """The sheet's default column width in the file's unit: ``defaultColWidth``, or
        ``baseColWidth`` (default 8) as Excel stores it."""
        node = self.root.find(_X + "sheetFormatPr")
        if node is not None and node.get("defaultColWidth"):
            try:
                return float(node.get("defaultColWidth"))
            except ValueError:
                pass
        base = BASE_COLUMN_WIDTH
        if node is not None and (node.get("baseColWidth") or "").isdigit():
            base = int(node.get("baseColWidth"))
        return stored_width(base, digit_width(default_face(self.workbook)))

    def set(self, first: int, last: int, *, width_default: bool = False,
            **changes: "str | None") -> bool:
        """Set (or, ``None``, remove) attributes of columns ``first`` to ``last``; with
        ``width_default`` a column given attributes gets the sheet's default width when it
        has none (Excel writes one with a column's style, measured)."""
        out: list[tuple[int, int, dict[str, str]]] = []
        touched = False
        for low, high, attrs in self.runs:
            if high < first or low > last:
                out.append((low, high, attrs))
                continue
            if low < first:
                out.append((low, first - 1, dict(attrs)))
            middle = dict(attrs)
            self._apply(middle, changes, width_default)
            touched |= middle != attrs
            out.append((max(low, first), min(high, last), middle))
            if high > last:
                out.append((last + 1, high, dict(attrs)))
        covered = sorted((max(low, first), min(high, last)) for low, high, _ in self.runs
                         if not (high < first or low > last))
        gaps: list[tuple[int, int]] = []
        position = first
        for low, high in covered:
            if low > position:
                gaps.append((position, low - 1))
            position = max(position, high + 1)
        if position <= last:
            gaps.append((position, last))
        for low, high in gaps:
            fresh: dict[str, str] = {}
            self._apply(fresh, changes, width_default)
            if fresh:
                touched = True
                out.append((low, high, fresh))
        out = [(low, high, attrs) for low, high, attrs in sorted(out) if attrs]
        self.runs = out
        self.changed |= touched
        return touched

    def _apply(self, attrs: dict[str, str], changes: dict[str, "str | None"],
               width_default: bool) -> None:
        for name, value in changes.items():
            if value is None:
                attrs.pop(name, None)
            else:
                attrs[name] = value
        if width_default and attrs and "width" not in attrs:
            attrs["width"] = number_text(self.default_width())

    def write(self) -> None:
        if not self.changed:
            return
        merged: list[tuple[int, int, dict[str, str]]] = []
        for low, high, attrs in self.runs:
            if merged and merged[-1][1] + 1 == low and merged[-1][2] == attrs:
                merged[-1] = (merged[-1][0], high, attrs)
            else:
                merged.append((low, high, attrs))
        node = self.root.find(_X + "cols")
        if not merged:
            if node is not None:
                self.root.remove(node)
        else:
            fresh = etree.Element(_X + "cols")
            order = ("width", "style", "hidden", "bestFit", "customWidth", "phonetic",
                     "outlineLevel", "collapsed")
            for low, high, attrs in merged:
                col = etree.SubElement(fresh, _X + "col")
                col.set("min", str(low))
                col.set("max", str(high))
                for name in order:
                    if name in attrs:
                        col.set(name, attrs[name])
                for name, value in attrs.items():
                    if name not in order:
                        col.set(name, value)
            if node is not None:
                node.addprevious(fresh)
                self.root.remove(node)
            else:
                place(self.root, fresh)
        self.writer.workbook.package.mark_dirty(self.writer.part)
        self.changed = False


def place(root: Any, child: Any) -> None:
    """Insert ``child`` among a worksheet's children in schema order."""
    name = etree.QName(child).localname
    later = SHEET_ORDER[SHEET_ORDER.index(name) + 1:]
    for existing in root:
        if isinstance(existing.tag, str) and etree.QName(existing).localname in later:
            existing.addprevious(child)
            return
    root.append(child)


# -- the calls -----------------------------------------------------------------------------------


def _start(workbook: Any) -> None:
    from . import EditError

    if workbook.strict:
        raise EditError("strict")


def set_column_widths(workbook: Any, target: Any, width: Any) -> Formatted:
    """Give the columns ``target`` names (``Stock!A:C``, or a range: its columns) a width in
    characters, ``"fit"`` (Excel's AutoFit, on the range's cells or the whole columns'),
    or ``None`` (their own width taken away)."""
    _start(workbook)
    problems = Problems()
    found = targets(workbook, target, problems, spans=("cells", "columns"))
    fit = width == "fit" or width == "auto"
    if not fit and width is not None and (isinstance(width, bool)
                                          or not isinstance(width, (int, float))
                                          or not math.isfinite(width)
                                          or not 0 <= width <= MAX_WIDTH):
        problems.add("bad_format", "width", f'characters, 0 to {MAX_WIDTH}, "fit", or None '
                     f"(the sheet's default), not {width!r}")
    guard(workbook, found, problems)
    editor = workbook.editor
    if editor.cells.styles is None and width is not None:
        problems.add("bad_format", "width", "the workbook has no stylesheet: its default "
                     "font, which widths are measured in, is not known")
    digit = digit_width(default_face(workbook)) if width is not None else 1
    fits: dict[tuple[str, int], float] = {}
    skipped: dict[str, list[str]] = {}
    notes: list[str] = []
    if fit and not problems.found:
        from .fit import fit_columns

        fits, skipped = fit_columns(workbook, found, digit, problems)
    problems.raise_any()
    widths: dict[str, float] = {}
    changed = 0
    with editor.batch():
        for item in found:
            writer = editor.cells.sheet(item.sheet.stored_name)
            columns = Columns(workbook, writer)
            for number in range(item.area.first.column, item.area.last.column + 1):
                letters = column_letters(number)
                if fit:
                    if (item.sheet.part, number) not in fits:
                        continue
                    value = fits[(item.sheet.part, number)]
                    done = columns.set(number, number, width=number_text(value),
                                       bestFit="1", customWidth="1")
                elif width is None:
                    value = None
                    done = columns.set(number, number, width=None, bestFit=None,
                                       customWidth=None, width_default=True)
                else:
                    value = stored_width(width, digit)
                    done = columns.set(number, number, width=number_text(value),
                                       customWidth="1", bestFit=None)
                changed += done
                if value is not None:
                    widths[_shown_column(item, letters)] = value
            columns.write()
            editor.touched_format(item.sheet)
        workbook._forget()
    if fit and not widths:
        notes.append("no cell in these columns could be measured; their widths are as "
                     "they were")
    if not changed:
        notes.append("the columns already have these widths")
    return Formatted(changed=changed, widths=widths, skipped=skipped, notes=tuple(notes),
                     columns=tuple(item.text for item in found))


def _shown_column(item: Target, letters: str) -> str:
    from ..address import format_cell

    sheet = format_cell(_label(item.sheet), 1, 1).rpartition("!")[0]
    return f"{sheet}!{letters}" if sheet else letters


def _label(sheet: Any) -> str:
    from .structure import _label as label

    return label(sheet)


def set_row_heights(workbook: Any, target: Any, height: Any) -> Formatted:
    """Give the rows ``target`` names (``Stock!1:3``, or a range: its rows) a height in
    points, or with ``None`` let them fit their text again."""
    _start(workbook)
    problems = Problems()
    found = targets(workbook, target, problems, spans=("cells", "rows"))
    if height is not None and (isinstance(height, bool) or not isinstance(height, (int, float))
                               or not math.isfinite(height) or not 0 <= height <= 409):
        problems.add("bad_format", "height", f"points, 0 to 409, or None (the row fits its "
                     f"text), not {height!r}")
    guard(workbook, found, problems)
    total = sum(item.area.rows for item in found)
    if total > 100_000:
        problems.add("too_many_cells", found[0].text, f"{total:,} rows; at most 100,000 in "
                     "one call")
    problems.raise_any()
    editor = workbook.editor
    changed = 0
    heights: dict[str, float] = {}
    with editor.batch():
        for item in found:
            writer = editor.cells.sheet(item.sheet.stored_name)
            for number in range(item.area.first.row, item.area.last.row + 1):
                changed += writer.set_row_height(number, height)
                if height is not None:
                    heights[f"{_label(item.sheet)}!{number}"] = float(height)
            editor.touched_format(item.sheet)
        workbook._forget()
    notes = () if changed else ("the rows already have this height",)
    return Formatted(changed=changed, heights=heights, notes=notes,
                     rows=tuple(item.text for item in found))


def _sheet_target(workbook: Any, sheet: Any, problems: Problems) -> Target | None:
    text = sheet if isinstance(sheet, str) else getattr(sheet, "name", None)
    if not isinstance(text, str) or not text.strip():
        problems.add("bad_target", "sheet", "a sheet's name, as Stock")
        return None
    address = text if text.rstrip().endswith("!") else _quoted(text.strip()) + "!"
    found = targets(workbook, address, problems, spans=("sheet",))
    return found[0] if found else None


def _quoted(name: str) -> str:
    if name.startswith("'") and name.endswith("'"):
        return name
    return "'" + name.replace("'", "''") + "'"


def _view(root: Any) -> Any:
    """The sheet's first ``sheetView``, made when it has none."""
    views = root.find(_X + "sheetViews")
    if views is None:
        views = etree.Element(_X + "sheetViews")
        place(root, views)
    view = views.find(_X + "sheetView")
    if view is None:
        view = etree.SubElement(views, _X + "sheetView", workbookViewId="0")
    return view


def freeze_panes(workbook: Any, sheet: Any, *, rows: int | None = None,
                 columns: int | None = None, at: str | None = None) -> Formatted:
    """Freeze the top ``rows`` and the left ``columns`` of ``sheet`` (or everything above
    and left of cell ``at``); ``rows=0, columns=0`` unfreezes."""

    _start(workbook)
    problems = Problems()
    item = _sheet_target(workbook, sheet, problems)
    if at is not None:
        if rows is not None or columns is not None:
            problems.add("bad_format", "at", "give at, or rows and columns, not both")
        else:
            try:
                cell = Cell.parse(at.strip().replace("$", "").upper()) \
                    if isinstance(at, str) else None
                if cell is None or not (1 <= cell.row and 1 <= cell.column <= MAX_COLUMN):
                    raise ValueError(at)
                rows, columns = cell.row - 1, cell.column - 1
            except (ValueError, AttributeError):
                problems.add("bad_format", "at", f"a cell, as B2, not {at!r}")
    rows = rows or 0
    columns = columns or 0
    for name, value, limit in (("rows", rows, 1_048_575), ("columns", columns, MAX_COLUMN - 1)):
        if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= limit:
            problems.add("bad_format", name, f"a whole number of {name}, 0 to {limit}")
    if item is not None:
        guard(workbook, [item], problems, whole_sheet=True)
    problems.raise_any()
    assert item is not None
    editor = workbook.editor
    with editor.batch():
        writer = editor.cells.sheet(item.sheet.stored_name)
        root = writer.workbook.package.tree(writer.part)
        view = _view(root)
        before = etree.tostring(view)
        for node in view.findall(_X + "pane"):
            view.remove(node)
        kept = [node for node in view.findall(_X + "selection") if not node.get("pane")]
        for node in view.findall(_X + "selection"):
            view.remove(node)
        if rows or columns:
            top_left = Cell(rows + 1, columns + 1)
            pane = etree.Element(_X + "pane")
            if columns:
                pane.set("xSplit", str(columns))
            if rows:
                pane.set("ySplit", str(rows))
            pane.set("topLeftCell", top_left.ref)
            active = "bottomRight" if rows and columns else ("bottomLeft" if rows else "topRight")
            pane.set("activePane", active)
            pane.set("state", "frozen")
            view.insert(0, pane)
            selections = []
            if columns:
                selections.append(("topRight", Cell(1, columns + 1)))
            if rows:
                selections.append(("bottomLeft", Cell(rows + 1, 1)))
            if rows and columns:
                selections.append(("bottomRight", top_left))
            anchor = pane
            for name, cell in selections:
                node = etree.Element(_X + "selection")
                node.set("pane", name)
                node.set("activeCell", cell.ref)
                node.set("sqref", cell.ref)
                anchor.addnext(node)
                anchor = node
        else:
            for node in kept[:1]:
                view.insert(0, node)
        changed = etree.tostring(view) != before
        if changed:
            writer.workbook.package.mark_dirty(writer.part)
            editor.touched_format(item.sheet)
    notes = () if changed else ("the panes are already so",)
    return Formatted(changed=int(changed), notes=notes)


def set_gridlines(workbook: Any, sheet: Any, shown: Any) -> Formatted:
    """Show or hide ``sheet``'s gridlines (``showGridLines``)."""
    _start(workbook)
    problems = Problems()
    item = _sheet_target(workbook, sheet, problems)
    if not isinstance(shown, bool):
        problems.add("bad_format", "shown", f"true or false, not {shown!r}")
    if item is not None:
        guard(workbook, [item], problems, whole_sheet=True)
    problems.raise_any()
    assert item is not None
    editor = workbook.editor
    with editor.batch():
        writer = editor.cells.sheet(item.sheet.stored_name)
        root = writer.workbook.package.tree(writer.part)
        views = root.find(_X + "sheetViews")
        view = views.find(_X + "sheetView") if views is not None else None
        now = view is None or (view.get("showGridLines") or "1").strip().lower() not in ("0", "false")
        changed = now != shown
        if changed:
            view = _view(root)
            if shown:
                view.attrib.pop("showGridLines", None)
            else:
                attrs = dict(view.attrib)
                view.attrib.clear()
                view.set("showGridLines", "0")    # Excel writes it first (measured)
                for name, value in attrs.items():
                    view.set(name, value)
            writer.workbook.package.mark_dirty(writer.part)
            editor.touched_format(item.sheet)
    notes = () if changed else (f"the gridlines are already {'shown' if shown else 'hidden'}",)
    return Formatted(changed=int(changed), notes=notes)


def password_hash(password: str, salt: bytes, spin_count: int = SPIN_COUNT,
                  algorithm: str = "SHA-512") -> str:
    """A sheet password's hash (ECMA-376 Part 1, 18.3.1.85, with 18.2.29's procedure):
    the salt and the password's UTF-16LE bytes hashed, then the hash and each round's
    number (four bytes, little-endian) hashed ``spin_count`` times; base64."""
    name = algorithm.replace("-", "").lower()
    digest = hashlib.new(name, salt + password.encode("utf-16-le")).digest()
    for number in range(spin_count):
        digest = hashlib.new(name, digest + number.to_bytes(4, "little")).digest()
    return base64.b64encode(digest).decode("ascii")


def _protection(root: Any) -> Any | None:
    node = root.find(_X + "sheetProtection")
    return node


def _protected(node: Any | None) -> bool:
    return node is not None and (node.get("sheet") or "").strip().lower() in ("1", "true")


def _matches(node: Any, password: str) -> bool | None:
    """Whether ``password`` is the protection's; ``None`` when its hash is not one read
    here (the legacy ``password`` attribute, an unknown algorithm)."""
    algorithm = node.get("algorithmName")
    if not algorithm or algorithm.replace("-", "").lower() not in hashlib.algorithms_available:
        return None
    try:
        salt = base64.b64decode(node.get("saltValue") or "")
        spins = int(node.get("spinCount") or "0")
    except ValueError:
        return None
    return password_hash(password, salt, spins, algorithm) == node.get("hashValue")


def protect_sheet(workbook: Any, sheet: Any, *, password: str | None = None) -> Formatted:
    """Protect ``sheet`` as Excel's Protect Sheet does, with ``password`` when given."""
    _start(workbook)
    problems = Problems()
    item = _sheet_target(workbook, sheet, problems)
    if password is not None and (not isinstance(password, str) or not password
                                 or len(password) > 255):
        problems.add("bad_format", "password", "text of 1 to 255 characters, or None")
    if item is not None and not item.sheet.protected:
        guard(workbook, [item], problems, whole_sheet=True)
    problems.raise_any()
    assert item is not None
    editor = workbook.editor
    writer = editor.cells.sheet(item.sheet.stored_name)
    root = writer.workbook.package.tree(writer.part)
    current = _protection(root)
    if current is not None and _protected(current) and (
            (password is None and not current.get("hashValue") and not current.get("password"))
            or (password is not None and _matches(current, password))):
        return Formatted(notes=("the sheet is already protected so",))
    with editor.batch():
        fresh = etree.Element(_X + "sheetProtection")
        if password is not None:
            salt = os.urandom(16)
            fresh.set("algorithmName", "SHA-512")
            fresh.set("hashValue", password_hash(password, salt))
            fresh.set("saltValue", base64.b64encode(salt).decode("ascii"))
            fresh.set("spinCount", str(SPIN_COUNT))
        fresh.set("sheet", "1")
        fresh.set("objects", "1")
        fresh.set("scenarios", "1")
        if current is not None:
            current.addprevious(fresh)
            root.remove(current)
        else:
            place(root, fresh)
        writer.workbook.package.mark_dirty(writer.part)
        editor.touched_format(item.sheet)
        workbook._forget()
    return Formatted(changed=1)


def unprotect_sheet(workbook: Any, sheet: Any, *, password: str | None = None,
                    force: bool = False) -> Formatted:
    """Take ``sheet``'s protection away; a password-protected sheet needs its password
    (or, for the legacy hash this does not check, ``force``)."""
    from . import EditError

    _start(workbook)
    problems = Problems()
    item = _sheet_target(workbook, sheet, problems)
    problems.raise_any()
    assert item is not None
    editor = workbook.editor
    writer = editor.cells.sheet(item.sheet.stored_name)
    root = writer.workbook.package.tree(writer.part)
    current = _protection(root)
    if current is None or not _protected(current):
        return Formatted(notes=("the sheet is not protected",))
    if not force and (current.get("hashValue") or current.get("password")):
        if password is None:
            raise EditError("wrong_password", "the sheet has a password: give it",
                            addresses=[item.text])
        matched = _matches(current, password)
        if matched is None:
            raise EditError("wrong_password", "the sheet's password is the legacy 16-bit "
                            "hash, not checked here: the host's own code may pass force=True",
                            addresses=[item.text])
        if not matched:
            raise EditError("wrong_password", addresses=[item.text])
    problems = Problems()
    guard_item = Target(item.sheet, item.area, item.span, item.text)
    _unprotected_guard(workbook, guard_item, problems)
    problems.raise_any()
    with editor.batch():
        root.remove(current)
        writer.workbook.package.mark_dirty(writer.part)
        editor.touched_format(item.sheet)
        workbook._forget()
    return Formatted(changed=1)


def _unprotected_guard(workbook: Any, item: Target, problems: Problems) -> None:
    """The write policy, asked of a sheet about to be unprotected (its protection is what
    unprotecting changes, so that is not a refusal)."""
    from .policy import PolicyError

    try:
        allowed = workbook._allowed()
    except PolicyError as error:
        problems.add("bad_policy", "policy", error.message)
        return
    if allowed is None:
        return
    from .format import _covered

    if not _covered(allowed, item.sheet.part, Area(Cell(1, 1), Cell(1_048_576, MAX_COLUMN))):
        problems.add("not_writable", item.text, "the host's write policy does not cover the "
                     "whole sheet")
