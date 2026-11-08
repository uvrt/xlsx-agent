"""Cell formatting from code (0.2, E7; production round 44, R53): ``wb.format(...)``.

::

    wb.format("Stock!A1:H1", font={"bold": True, "color": "#1F3864", "size": 12},
              fill="#DDEBF7", border={"bottom": "thin"},
              alignment={"horizontal": "center", "vertical": "center", "wrap": True})
    wb.format("Stock!D2:D40", number_format="0.0")          # a code, or a built-in id
    wb.format("Stock!J2:J40", like="Stock!D2")              # that cell's whole format
    wb.format("Stock!A1:H1", style="Heading 1")             # a named cell style
    wb.format("Stock!E2:E40", protection={"locked": False}) # open to input once protected

**Only what is asked changes.**  Each cell's format (its ``xf`` in ``cellXfs``, ECMA-376
Part 1, 18.8.45) is the cell's own with the asked parts changed: setting bold keeps the
font's size, colour and name and the cell's fill, borders, number format, alignment and
protection.  A cell with no format of its own starts from the one its row or column gives
it (18.3.1.73, 18.3.1.13), as one typed there would.

**Styles are reused.**  A font, fill, border, number format or cell format equal to one the
workbook has is that one, never a copy: formatting a thousand cells alike adds at most one
of each, and formatting them again adds nothing.  Nothing that exists is changed or
reordered: what is added is appended, with the lists' ``count``.

**What Excel writes** (measured, Excel for Mac 16.106, 2026-11-08, Excel formatting cells
itself through AppleScript and saving): a font's parts in the order ``b``, ``i``,
``strike``, ``u``, ``sz``, ``color``, ``name``, ``family``, ``scheme`` (single underline
``<u/>``); a solid fill ``patternFill patternType="solid"`` with ``fgColor rgb`` and
``bgColor indexed="64"``; a border side with no colour chosen ``<color indexed="64"/>``; a
number format whose code is a built-in one's by its id (``0.00`` is 2, ``#,##0.00`` 4), any
other added from 164 on; an indent given to a cell aligned ``general`` makes it ``left``;
unlocked ``<protection locked="0"/>``; and each ``apply*`` flag of a cell format is ``1``
exactly where that part differs from the cell style it is based on (``xfId``), left out
where it is the same -- what is written here too, for every part a call changes.

**A named cell style** (``style``) is applied as Excel applies one (measured, the same
session: five styles applied by name to a cell with a number format, an italic font, thin
borders, centred and unlocked): the cell's format becomes the style's ``cellStyleXfs``
format (``xfId`` the style) for every part the style includes, and keeps its own for the
parts the style leaves out (``apply*="0"`` on the style's format; left out, a part is
included), with the ``apply*`` flags as above.  Styles are named as the file names them
(``cellStyles``), case aside; a built-in style also by its English name (ECMA-376 Part 1,
18.8.7: ``builtinId`` 16 is ``Heading 1``): a Dutch Excel saves ``Heading 1`` as ``Kop 1``
(measured).  A style the workbook does not have is refused with the names it has.

**like** (``like="Stock!D2"``) gives the cells that cell's whole format, from any sheet of
the workbook (one style table); further options apply on top of it.

**Whole columns and rows** (``Stock!D:D``, ``Stock!3:3``) are formatted as Excel formats
them (measured): a column's ``col style`` (with a width, the sheet's default when it has
none of its own) and every cell it holds; a row's ``s`` with ``customFormat="1"`` and
every cell it holds.  Excel makes no cell where nothing is.

**Refused** before anything changes, every problem listed (:class:`~xlsx_agent.edit.
EditError`): an option or value that is not one (``bad_format``, with what is valid), a
colour that is not ``#RRGGBB`` (``bad_format``), a style the workbook does not have
(``no_cell_style``, with the names it has), ``like`` and ``style`` together, a protected
sheet (``protected``; unless its protection allows formatting cells, ``formatCells="0"``,
and for whole columns or rows ``formatColumns``/``formatRows`` too, as Excel allows,
measured), cells the host's write policy does not cover (``not_writable``), and
more than :data:`MAX_CELLS` cells in one call (``too_many_cells``: format whole columns or
rows instead).  Formatting changes no value: the calculation status stays as it was.
"""

from __future__ import annotations

import copy
import math
import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, SML_NS, Area, Cell, number_text

from .report import capped

_X = "{%s}" % SML_NS

#: The most cells one call formats (a whole column or row counts the cells it holds).
MAX_CELLS = 1_000_000

#: CT_Font's parts in the order Excel writes them (ECMA-376 Part 1, 18.8.22; measured).
FONT_ORDER = ("b", "i", "strike", "condense", "extend", "outline", "shadow", "u",
              "vertAlign", "sz", "color", "name", "family", "charset", "scheme")
#: CT_Border's sides, in schema order (18.8.4).
BORDER_SIDES = ("left", "right", "top", "bottom", "diagonal")
#: ST_BorderStyle (18.18.3).
BORDER_STYLES = ("none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair",
                 "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
                 "mediumDashDotDot", "slantDashDot")
#: ST_HorizontalAlignment, ST_VerticalAlignment (18.18.40, 18.18.88).
HORIZONTAL = ("general", "left", "center", "right", "fill", "justify", "centerContinuous",
              "distributed")
VERTICAL = ("top", "center", "bottom", "justify", "distributed")
#: CT_CellAlignment's attributes in schema order (18.8.1).
ALIGNMENT_ORDER = ("horizontal", "vertical", "textRotation", "wrapText", "indent",
                   "relativeIndent", "justifyLastLine", "shrinkToFit", "readingOrder")
#: ST_UnderlineValues (18.18.85); ``True`` is ``single``.
UNDERLINES = ("single", "double", "singleAccounting", "doubleAccounting", "none")

#: The options each part takes.
FONT_KEYS = ("bold", "italic", "underline", "strike", "color", "size", "name")
ALIGNMENT_KEYS = ("horizontal", "vertical", "wrap", "indent", "rotation", "shrink")
PROTECTION_KEYS = ("locked", "hidden")
BORDER_KEYS = ("all", "outline", "left", "right", "top", "bottom")

#: The built-in number formats a code is written as (ECMA-376 Part 1, 18.8.30, "All
#: Languages"; Excel wrote ``0.00`` and ``#,##0.00`` as 2 and 4, measured).  Dates and times
#: (14-22, 45-47) are left out: Excel shows some of them in the system's own date order, so
#: their codes are written as the workbook's own formats, which show as written.
BUILTIN_CODES = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9,
                 "0.00%": 10, "0.00E+00": 11, "# ?/?": 12, "# ??/??": 13,
                 "#,##0 ;(#,##0)": 37, "#,##0 ;[Red](#,##0)": 38,
                 "#,##0.00;(#,##0.00)": 39, "#,##0.00;[Red](#,##0.00)": 40,
                 "##0.0E+0": 48, "@": 49}

#: ECMA-376 Part 1, 18.8.7 (``cellStyle builtinId``): the English names of the built-in
#: styles a workbook may hold under its application's language (a Dutch Excel writes
#: ``Kop 1`` for ``Heading 1``, measured).
BUILTIN_STYLE_NAMES = {0: "Normal", 3: "Comma", 4: "Currency", 5: "Percent", 6: "Comma [0]",
                       7: "Currency [0]", 8: "Hyperlink", 9: "Followed Hyperlink",
                       10: "Note", 11: "Warning Text", 15: "Title", 16: "Heading 1",
                       17: "Heading 2", 18: "Heading 3", 19: "Heading 4", 20: "Input",
                       21: "Output", 22: "Calculation", 23: "Check Cell", 24: "Linked Cell",
                       25: "Total", 26: "Good", 27: "Bad", 28: "Neutral",
                       53: "Explanatory Text"}

#: The parts of a cell format, with their ``apply*`` flag.
ASPECTS = ("NumberFormat", "Font", "Fill", "Border", "Alignment", "Protection")
_ID_OF = {"NumberFormat": "numFmtId", "Font": "fontId", "Fill": "fillId",
          "Border": "borderId"}

_HEX = re.compile(r"#?([0-9A-Fa-f]{6})")
_COLOUR = '"#RRGGBB", as "#1F3864"'


@dataclass(frozen=True)
class Formatted:
    """What a formatting call did.  ``cells``: the cells formatted, by sheet, as ranges
    (capped as an edit report's); ``changed`` the cells whose format changed;
    ``formats_added`` the cell formats appended to the style table (0 when every format
    needed existed); ``columns``/``rows`` the whole columns and rows formatted; ``widths``
    and ``heights`` what a width or height call set, by column letters or row number;
    ``skipped`` cells a fit did not measure, by reason; ``notes`` what was read as meant."""

    cells: dict[str, list[str]] = field(default_factory=dict)
    changed: int = 0
    formats_added: int = 0
    columns: tuple[str, ...] = ()
    rows: tuple[str, ...] = ()
    widths: dict[str, float] = field(default_factory=dict)
    heights: dict[str, float] = field(default_factory=dict)
    skipped: dict[str, list[str]] = field(default_factory=dict)
    notes: tuple[str, ...] = ()

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"changed": self.changed}
        for name in ("cells", "columns", "rows", "widths", "heights", "skipped", "notes"):
            value = getattr(self, name)
            if value:
                out[name] = dict(value) if isinstance(value, dict) else list(value)
        if self.formats_added:
            out["formats_added"] = self.formats_added
        return out


# -- the request ---------------------------------------------------------------------------------


@dataclass
class Spec:
    """A formatting call's options, checked."""

    font: dict[str, Any] = field(default_factory=dict)
    fill: "str | None | bool" = False          # False: not asked; None: no fill
    border: dict[str, "tuple[str, str | None]"] = field(default_factory=dict)
    alignment: dict[str, Any] = field(default_factory=dict)
    protection: dict[str, bool] = field(default_factory=dict)
    number_format: "int | str | None" = None
    style: int | None = None                   # a cellStyleXfs index
    like: int | None = None                    # a cellXfs index

    @property
    def empty(self) -> bool:
        return not (self.font or self.fill is not False or self.border or self.alignment
                    or self.protection or self.number_format is not None
                    or self.style is not None or self.like is not None)


class Problems:
    """The refusals of one call, listed together (:class:`~xlsx_agent.edit.EditError`)."""

    def __init__(self) -> None:
        self.found: list[tuple[str, str, str]] = []

    def add(self, code: str, where: str, detail: str) -> None:
        self.found.append((code, where, detail))

    def raise_any(self) -> None:
        from . import EditError

        if not self.found:
            return
        code, where, detail = self.found[0]
        same = [w for c, w, _ in self.found if c == code]
        others = [f"{w}: {d}" for c, w, d in self.found[1:]]
        if others:
            detail = detail + "; also " + "; ".join(others[:10])
        raise EditError(code, detail, addresses=same,
                        problems=[(c, w) for c, w, _ in self.found])


def colour(value: Any, where: str, problems: Problems) -> str | None:
    """``#RRGGBB`` (or ``RRGGBB``) as the ARGB a file writes; ``None`` when refused."""
    if isinstance(value, str):
        match = _HEX.fullmatch(value.strip())
        if match:
            return "FF" + match.group(1).upper()
    problems.add("bad_format", where, f"a colour is {_COLOUR}, not {value!r}")
    return None


def _flag(value: Any, where: str, problems: Problems) -> bool | None:
    if isinstance(value, bool):
        return value
    problems.add("bad_format", where, f"true or false, not {value!r}")
    return None


def read_spec(workbook: Any, problems: Problems, *, font: Any = None, fill: Any = False,
              border: Any = None, alignment: Any = None, protection: Any = None,
              number_format: Any = None, like: Any = None, style: Any = None,
              styles: "StyleTable | None" = None) -> Spec:
    """The options of :func:`format_cells`, checked: every problem added to ``problems``."""
    spec = Spec()
    if font is not None:
        spec.font = _font_spec(font, problems)
    if fill is not False:
        if fill is None or fill == "none" or fill == {}:
            spec.fill = None
        elif isinstance(fill, Mapping) and set(fill) <= {"color", "colour"} and fill:
            spec.fill = colour(next(iter(fill.values())), "fill", problems)
        else:
            spec.fill = colour(fill, "fill", problems)
    if border is not None:
        spec.border = _border_spec(border, problems)
    if alignment is not None:
        spec.alignment = _alignment_spec(alignment, problems)
    if protection is not None:
        if not isinstance(protection, Mapping) or not protection:
            problems.add("bad_format", "protection",
                         f"an object of {', '.join(PROTECTION_KEYS)}, as {{\"locked\": false}}")
        else:
            for key, value in protection.items():
                if key not in PROTECTION_KEYS:
                    problems.add("bad_format", f"protection.{key}",
                                 f"one of {', '.join(PROTECTION_KEYS)}")
                    continue
                flag = _flag(value, f"protection.{key}", problems)
                if flag is not None:
                    spec.protection[key] = flag
    if number_format is not None:
        spec.number_format = _number_format_spec(number_format, problems)
    if like is not None and style is not None:
        problems.add("bad_format", "like", "like and style both give the whole format: give one")
    elif like is not None:
        spec.like = _like(workbook, like, problems)
    elif style is not None and styles is not None:
        spec.style = styles.named_style(style, problems)
    return spec


def _font_spec(font: Any, problems: Problems) -> dict[str, Any]:
    if not isinstance(font, Mapping) or not font:
        problems.add("bad_format", "font", f"an object of {', '.join(FONT_KEYS)}, as "
                     '{"bold": true, "size": 12}')
        return {}
    out: dict[str, Any] = {}
    for key, value in font.items():
        where = f"font.{key}"
        if key == "colour":
            key = "color"
        if key in ("bold", "italic", "strike"):
            flag = _flag(value, where, problems)
            if flag is not None:
                out[key] = flag
        elif key == "underline":
            if value is True:
                out[key] = "single"
            elif value is False or value is None:
                out[key] = "none"
            elif isinstance(value, str) and value in UNDERLINES:
                out[key] = value
            else:
                problems.add("bad_format", where, f"true, false or one of {', '.join(UNDERLINES)}")
        elif key == "color":
            found = colour(value, where, problems)
            if found:
                out[key] = found
        elif key == "size":
            if isinstance(value, bool) or not isinstance(value, (int, float)) \
                    or not math.isfinite(value) or not 1 <= value <= 409:
                problems.add("bad_format", where, f"a size in points, 1 to 409, not {value!r}")
            else:
                out[key] = value
        elif key == "name":
            if not isinstance(value, str) or not value.strip() or len(value) > 31:
                problems.add("bad_format", where, "a font's name, as Arial (at most 31 "
                             "characters)")
            else:
                out[key] = value.strip()
        else:
            problems.add("bad_format", where, f"a font takes {', '.join(FONT_KEYS)}")
    return out


def _side(value: Any, where: str, problems: Problems) -> "tuple[str, str | None] | None":
    """A border side: its style and colour (``None``: Excel's automatic colour)."""
    if value is None or value is False:
        return ("none", None)
    if value is True:
        return ("thin", None)
    if isinstance(value, str):
        value = {"style": value}
    if not isinstance(value, Mapping) or not set(value) <= {"style", "color", "colour"}:
        problems.add("bad_format", where, 'a style, as "thin", or {"style": "thin", '
                     f'"color": "#000000"}}; styles: {", ".join(BORDER_STYLES)}')
        return None
    style = value.get("style", "thin")
    if style not in BORDER_STYLES:
        problems.add("bad_format", where, f"a border style is one of {', '.join(BORDER_STYLES)}")
        return None
    given = value.get("color", value.get("colour"))
    found = None if given is None else colour(given, where, problems)
    if given is not None and found is None:
        return None
    return (style, found)


def _border_spec(border: Any, problems: Problems) -> dict[str, "tuple[str, str | None]"]:
    if isinstance(border, (str, bool)) or border is None:
        border = {"all": border}
    if not isinstance(border, Mapping) or not border:
        problems.add("bad_format", "border", f"an object of {', '.join(BORDER_KEYS)}, as "
                     '{"bottom": "thin"}')
        return {}
    out: dict[str, tuple[str, str | None]] = {}
    for key, value in border.items():
        if key not in BORDER_KEYS:
            problems.add("bad_format", f"border.{key}", f"one of {', '.join(BORDER_KEYS)}")
            continue
        side = _side(value, f"border.{key}", problems)
        if side is not None:
            out[key] = side
    return out


def _alignment_spec(alignment: Any, problems: Problems) -> dict[str, Any]:
    if not isinstance(alignment, Mapping) or not alignment:
        problems.add("bad_format", "alignment", f"an object of {', '.join(ALIGNMENT_KEYS)}, "
                     'as {"horizontal": "center", "wrap": true}')
        return {}
    out: dict[str, Any] = {}
    for key, value in alignment.items():
        where = f"alignment.{key}"
        if key == "horizontal":
            if value == "centre":
                value = "center"
            if value not in HORIZONTAL:
                problems.add("bad_format", where, f"one of {', '.join(HORIZONTAL)}")
            else:
                out[key] = value
        elif key == "vertical":
            if value in ("middle", "centre"):
                value = "center"
            if value not in VERTICAL:
                problems.add("bad_format", where, f"one of {', '.join(VERTICAL)}")
            else:
                out[key] = value
        elif key in ("wrap", "shrink"):
            flag = _flag(value, where, problems)
            if flag is not None:
                out[key] = flag
        elif key == "indent":
            if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= 250:
                problems.add("bad_format", where, "a whole number of indent steps, 0 to 250")
            else:
                out[key] = value
        elif key == "rotation":
            if isinstance(value, bool) or not isinstance(value, int) \
                    or not (-90 <= value <= 180 or value == 255):
                problems.add("bad_format", where, "degrees, -90 to 90 (or 0 to 180 as the "
                             "file writes them), or 255 for text stacked downwards")
            else:
                # ECMA-376 Part 1, 18.8.1 textRotation: 0-90 up, 91-180 down by (n - 90).
                out[key] = 90 - value if value < 0 else value
        else:
            problems.add("bad_format", where, f"alignment takes {', '.join(ALIGNMENT_KEYS)}")
    return out


def _number_format_spec(value: Any, problems: Problems) -> "int | str | None":
    if isinstance(value, bool):
        problems.add("bad_format", "number_format", "a format code, as 0.0, or a built-in id")
        return None
    if isinstance(value, int):
        from ..read.styles import BUILTIN_FORMATS

        if value in BUILTIN_FORMATS:
            return value
        problems.add("bad_format", "number_format",
                     f"{value} is not a built-in format id; give its code, as 0.0")
        return None
    if not isinstance(value, str) or not value.strip() or len(value) > 255:
        problems.add("bad_format", "number_format",
                     "a format code (at most 255 characters), as 0.0 or #,##0.00")
        return None
    return value


def _like(workbook: Any, like: Any, problems: Problems) -> int | None:
    from ..address import AddressError, NameNotFound, RangeAddress, parse_address

    try:
        parsed = parse_address(like) if isinstance(like, str) else None
    except (AddressError, ValueError):
        parsed = None
    if not isinstance(parsed, RangeAddress) or parsed.kind != "cell" or parsed.sheet is None:
        problems.add("bad_format", "like", f"one cell with its sheet, as Stock!D2, not {like!r}")
        return None
    try:
        sheet = workbook.sheet_of(parsed)
    except (NameNotFound, KeyError, ValueError):
        problems.add("bad_format", "like", f"no sheet in {like!r}; sheets: "
                     + ", ".join(s.name for s in workbook.sheets[:20]))
        return None
    writer = workbook.editor.cells.sheet(sheet.stored_name) if sheet.part else None
    if writer is None:
        problems.add("bad_format", "like", f"{like} is not on a worksheet")
        return None
    return writer.style(parsed.area.first, inherited=True)


# -- the style table ------------------------------------------------------------------------------


def _canonical(node: Any) -> tuple:
    """An element's identity, its children's order aside where the schema leaves it free
    (a font's parts): tag, attributes, text, children."""
    children = tuple(sorted(_canonical(child) for child in node)) \
        if node.tag == _X + "font" else tuple(_canonical(child) for child in node)
    return (etree.QName(node).localname, tuple(sorted(node.attrib.items())),
            (node.text or "").strip(), children)


class StyleTable:
    """The workbook's ``styles.xml``, for formats made from code: fonts, fills, borders
    and number formats found or appended, cell formats through ooxml-edit's
    :class:`~ooxml_edit.sheets.Styles` (its index reused, :meth:`Styles.add_format`)."""

    LISTS = {"fonts": "font", "fills": "fill", "borders": "border"}
    #: styleSheet's children in schema order (18.8.39).
    ORDER = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
             "cellStyles", "dxfs", "tableStyles", "colors", "extLst")

    def __init__(self, workbook: Any) -> None:
        self.workbook = workbook
        self.styles = workbook.editor.cells.styles
        self._keys: dict[str, dict[tuple, int]] = {}
        self.added = 0

    @property
    def root(self) -> Any:
        return self.styles.package.tree(self.styles.part)

    def items(self, container: str) -> list[Any]:
        node = self.root.find(_X + container)
        tag = self.LISTS.get(container, "xf")
        return [] if node is None else node.findall(_X + tag)

    def xf(self, index: int) -> Any:
        xfs = self.items("cellXfs")
        return xfs[index] if 0 <= index < len(xfs) else xfs[0]

    def style_xf(self, index: int) -> Any | None:
        found = self.items("cellStyleXfs")
        return found[index] if 0 <= index < len(found) else None

    def find_or_add(self, container: str, element: Any) -> int:
        """The index of an item equal to ``element`` in ``container`` (``fonts``, ``fills``,
        ``borders``), or its own once appended."""
        keys = self._keys.get(container)
        items = self.items(container)
        if keys is None:
            keys = {}
            for position, item in enumerate(items):
                keys.setdefault(_canonical(item), position)
            self._keys[container] = keys
        key = _canonical(element)
        if key in keys:
            return keys[key]
        parent = self._container(container)
        if items:
            items[-1].addnext(element)
        else:
            parent.append(element)
        if parent.get("count") is not None or len(parent) > 1:
            parent.set("count", str(len(items) + 1))
        keys[key] = len(items)
        self.styles.package.mark_dirty(self.styles.part)
        return len(items)

    def _container(self, name: str) -> Any:
        root = self.root
        node = root.find(_X + name)
        if node is not None:
            return node
        node = etree.Element(_X + name)
        later = self.ORDER[self.ORDER.index(name) + 1:]
        for child in root:
            if etree.QName(child).localname in later:
                child.addprevious(node)
                break
        else:
            root.append(node)
        return node

    def number_format_id(self, given: "int | str") -> int:
        """A built-in id as given; a code as the built-in id it is (:data:`BUILTIN_CODES`),
        the workbook's own format of that code, or one added from 164 on."""
        if isinstance(given, int):
            return given
        if given in BUILTIN_CODES:
            return BUILTIN_CODES[given]
        container = self.root.find(_X + "numFmts")
        used: list[int] = []
        if container is not None:
            for node in container.iterchildren(_X + "numFmt"):
                raw = node.get("numFmtId", "")
                if raw.isdigit():
                    used.append(int(raw))
                    if node.get("formatCode") == given:
                        return int(raw)
        container = self._container("numFmts")
        number = max([163, *used]) + 1
        node = etree.SubElement(container, _X + "numFmt")
        node.set("numFmtId", str(number))
        node.set("formatCode", given)
        container.set("count", str(len(container.findall(_X + "numFmt"))))
        self.styles.package.mark_dirty(self.styles.part)
        return number

    def named_style(self, name: Any, problems: Problems) -> int | None:
        """The ``cellStyleXfs`` index of the cell style named ``name`` (as the file names
        it, case aside, or a built-in style's English name)."""
        container = self.root.find(_X + "cellStyles")
        found: list[tuple[str, int, int | None]] = []
        if container is not None:
            for node in container.iterchildren(_X + "cellStyle"):
                raw, builtin = node.get("xfId", ""), node.get("builtinId", "")
                if raw.isdigit():
                    found.append((node.get("name", ""), int(raw),
                                  int(builtin) if builtin.isdigit() else None))
        if isinstance(name, str) and name.strip():
            wanted = name.strip().casefold()
            for stored, xf_id, _ in found:
                if stored.casefold() == wanted:
                    return xf_id
            for _, xf_id, builtin in found:
                if builtin is not None and BUILTIN_STYLE_NAMES.get(builtin, "").casefold() == wanted:
                    return xf_id
        names = [stored for stored, _, _ in found][:40]
        problems.add("no_cell_style", "style",
                     f"the workbook has no cell style {name!r}; it has: "
                     + (", ".join(names) if names else "none"))
        return None

    # -- one cell format made from another ---------------------------------------------------

    def derive(self, base: int, spec: Spec, edges: "frozenset[str]" = frozenset()) -> int:
        """The cell format that is ``base`` with ``spec`` applied (``edges``: which sides
        of the range the cell is on, for an ``outline`` border)."""
        if spec.like is not None:
            base = spec.like
        start = self.xf(base)
        xf = copy.deepcopy(start)
        changed: set[str] = set()
        if spec.style is not None:
            changed |= self._apply_style(xf, spec.style)
        if spec.number_format is not None:
            xf.set("numFmtId", str(self.number_format_id(spec.number_format)))
            changed.add("NumberFormat")
        if spec.font:
            xf.set("fontId", str(self.find_or_add("fonts", self._font(xf, spec.font))))
            changed.add("Font")
        if spec.fill is not False:
            xf.set("fillId", str(self._fill(spec.fill if isinstance(spec.fill, str) else None)))
            changed.add("Fill")
        if spec.border:
            wanted = self._border(xf, spec.border, edges)
            if wanted is not None:
                xf.set("borderId", str(self.find_or_add("borders", wanted)))
                changed.add("Border")
        if spec.alignment:
            self._alignment(xf, spec.alignment)
            changed.add("Alignment")
        if spec.protection:
            self._protection(xf, spec.protection)
            changed.add("Protection")
        self._flags(xf, changed)
        if not changed and spec.like is None:
            return base
        before = self.styles.count
        index = self.styles.add_format(xf)
        self.added += self.styles.count - before
        return index

    def _apply_style(self, xf: Any, style: int) -> set[str]:
        source = self.style_xf(style)
        if source is None:
            return set()
        for aspect in ASPECTS:
            if source.get("apply" + aspect, "1").strip().lower() in ("0", "false"):
                continue                                   # the style leaves it out
            if aspect in _ID_OF:
                xf.set(_ID_OF[aspect], source.get(_ID_OF[aspect], "0"))
            else:
                tag = _X + aspect.lower()
                for node in xf.findall(tag):
                    xf.remove(node)
                given = source.find(tag)
                if aspect == "Protection":
                    locked, hidden = _part(source, "Protection")
                    self._protection(xf, {"locked": locked, "hidden": hidden})
                elif given is not None and len(given.attrib):
                    _insert_xf_child(xf, copy.deepcopy(given))
        xf.set("xfId", str(style))
        return set(ASPECTS)

    def _flags(self, xf: Any, changed: Iterable[str]) -> None:
        """Each changed part's ``apply*`` flag: ``1`` where it differs from the cell style
        the format is based on, left out where it is the same (measured, Excel)."""
        raw = xf.get("xfId", "0")
        style = self.style_xf(int(raw)) if raw.isdigit() else None
        for aspect in changed:
            if style is not None and _part(xf, aspect) == _part(style, aspect):
                xf.attrib.pop("apply" + aspect, None)
            else:
                xf.set("apply" + aspect, "1")
        # schema order of the attributes, as Excel writes them
        order = ("numFmtId", "fontId", "fillId", "borderId", "xfId", "quotePrefix",
                 "pivotButton", "applyNumberFormat", "applyFont", "applyFill",
                 "applyBorder", "applyAlignment", "applyProtection")
        items = dict(xf.attrib)
        xf.attrib.clear()
        for name in order:
            if name in items:
                xf.set(name, items.pop(name))
        for name, value in items.items():
            xf.set(name, value)

    def _font(self, xf: Any, wanted: Mapping[str, Any]) -> Any:
        fonts = self.items("fonts")
        raw = xf.get("fontId", "0")
        number = int(raw) if raw.isdigit() else 0
        font = copy.deepcopy(fonts[number]) if number < len(fonts) else etree.Element(_X + "font")
        parts = {etree.QName(child).localname: child for child in font
                 if isinstance(child.tag, str)}

        def put(tag: str, value: str | None = None) -> None:
            node = etree.Element(_X + tag)
            if value is not None:
                node.set("val" if tag != "color" else "rgb", value)
            parts[tag] = node

        for key, value in wanted.items():
            tag = {"bold": "b", "italic": "i", "strike": "strike", "underline": "u",
                   "size": "sz", "color": "color", "name": "name"}[key]
            if key in ("bold", "italic", "strike"):
                if value:
                    put(tag)
                else:
                    parts.pop(tag, None)
            elif key == "underline":
                if value == "none":
                    parts.pop("u", None)
                else:
                    put("u", None if value == "single" else value)
            elif key == "size":
                put("sz", number_text(value))
            elif key == "color":
                put("color", value)
            elif key == "name":
                if parts.get("name") is None or parts["name"].get("val") != value:
                    put("name", value)
                    # a theme font's binding would show the theme's font, not the name, and
                    # the family class belongs to the face it named
                    for gone in ("scheme", "family", "charset"):
                        parts.pop(gone, None)
        out = etree.Element(_X + "font")
        for tag in FONT_ORDER:
            if tag in parts:
                out.append(parts.pop(tag))
        for node in parts.values():
            out.append(node)
        return out

    def _fill(self, wanted: "str | None") -> int:
        if wanted is None:
            return 0
        fill = etree.Element(_X + "fill")
        pattern = etree.SubElement(fill, _X + "patternFill", patternType="solid")
        etree.SubElement(pattern, _X + "fgColor", rgb=wanted)
        etree.SubElement(pattern, _X + "bgColor", indexed="64")
        return self.find_or_add("fills", fill)

    def _border(self, xf: Any, wanted: Mapping[str, tuple[str, str | None]],
                edges: frozenset[str]) -> Any | None:
        sides: dict[str, tuple[str, str | None]] = {}
        for key, value in wanted.items():
            if key == "all":
                for side in ("left", "right", "top", "bottom"):
                    sides.setdefault(side, value)
            elif key == "outline":
                for side in edges:
                    sides.setdefault(side, value)
            else:
                sides[key] = value
        if not sides:
            return None
        borders = self.items("borders")
        raw = xf.get("borderId", "0")
        number = int(raw) if raw.isdigit() else 0
        border = copy.deepcopy(borders[number]) if number < len(borders) \
            else etree.Element(_X + "border")
        present = {etree.QName(child).localname: child for child in border
                   if isinstance(child.tag, str)}
        for side, (style, rgb) in sides.items():
            node = etree.Element(_X + side)
            if style != "none":
                node.set("style", style)
                color = etree.SubElement(node, _X + "color")
                if rgb is None:
                    color.set("indexed", "64")     # Excel's automatic colour, measured
                else:
                    color.set("rgb", rgb)
            present[side] = node
        out = etree.Element(_X + "border")
        for name, value in border.attrib.items():
            out.set(name, value)
        for side in BORDER_SIDES:
            node = present.pop(side, None)
            out.append(node if node is not None else etree.Element(_X + side))
        for node in present.values():
            out.append(node)
        return out

    def _alignment(self, xf: Any, wanted: Mapping[str, Any]) -> None:
        node = xf.find(_X + "alignment")
        values = dict(node.attrib) if node is not None else {}
        for key, value in wanted.items():
            if key in ("horizontal", "vertical"):
                values[key] = value
            elif key == "wrap":
                if value:
                    values["wrapText"] = "1"
                else:
                    values.pop("wrapText", None)
            elif key == "shrink":
                if value:
                    values["shrinkToFit"] = "1"
                else:
                    values.pop("shrinkToFit", None)
            elif key == "indent":
                if value:
                    values["indent"] = str(value)
                    if values.get("horizontal", "general") == "general":
                        values["horizontal"] = "left"   # what Excel does (measured)
                else:
                    values.pop("indent", None)
            elif key == "rotation":
                if value:
                    values["textRotation"] = str(value)
                else:
                    values.pop("textRotation", None)
        if values.get("horizontal") == "general":
            del values["horizontal"]
        if node is not None:
            xf.remove(node)
        if values:
            fresh = etree.Element(_X + "alignment")
            for name in ALIGNMENT_ORDER:
                if name in values:
                    fresh.set(name, values.pop(name))
            for name, value in values.items():
                fresh.set(name, value)
            _insert_xf_child(xf, fresh)

    def _protection(self, xf: Any, wanted: Mapping[str, bool]) -> None:
        node = xf.find(_X + "protection")
        locked = True if node is None else _true(node.get("locked"), True)
        hidden = False if node is None else _true(node.get("hidden"), False)
        locked = wanted.get("locked", locked)
        hidden = wanted.get("hidden", hidden)
        if node is not None:
            xf.remove(node)
        if not locked or hidden:
            fresh = etree.Element(_X + "protection")
            if not locked:
                fresh.set("locked", "0")
            if hidden:
                fresh.set("hidden", "1")
            _insert_xf_child(xf, fresh)


def _true(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in ("1", "true")


def _part(xf: Any, aspect: str) -> Any:
    """One part of a cell format, as compared for its ``apply*`` flag."""
    if aspect in _ID_OF:
        return xf.get(_ID_OF[aspect], "0")
    node = xf.find(_X + aspect.lower())
    if aspect == "Protection":
        if node is None:
            return (True, False)
        return (_true(node.get("locked"), True), _true(node.get("hidden"), False))
    return () if node is None else tuple(sorted(node.attrib.items()))


def _insert_xf_child(xf: Any, child: Any) -> None:
    """``alignment`` before ``protection`` before ``extLst`` (18.8.45)."""
    order = ("alignment", "protection", "extLst")
    rank = order.index(etree.QName(child).localname)
    for existing in xf:
        name = etree.QName(existing).localname if isinstance(existing.tag, str) else ""
        if name in order and order.index(name) > rank:
            existing.addprevious(child)
            return
    xf.append(child)


# -- targets ----------------------------------------------------------------------------------------


@dataclass
class Target:
    """One area of a call: its sheet, the area, and whether it is whole columns or rows."""

    sheet: Any
    area: Area
    span: str
    text: str


def targets(workbook: Any, target: Any, problems: Problems, *,
            spans: tuple[str, ...] = ("cells", "columns", "rows")) -> list[Target]:
    """The areas ``target`` names (an address, or a list of them), each on a worksheet."""
    from ..address import AddressError, NameNotFound, RangeAddress, parse_address

    given = [target] if isinstance(target, str) else target
    if not isinstance(given, (list, tuple)) or not given:
        problems.add("bad_target", "target", "a range with its sheet, as Stock!A1:H1, or a "
                     "list of them")
        return []
    out: list[Target] = []
    for text in given:
        shown = str(text)[:80]
        try:
            parsed = parse_address(text) if isinstance(text, str) else None
        except (AddressError, ValueError) as error:
            problems.add("bad_target", shown, str(error)[:200])
            continue
        if not isinstance(parsed, RangeAddress) or parsed.sheet is None:
            problems.add("bad_target", shown, "a range with its sheet, as Stock!A1:H1, "
                         "Stock!D:D or Stock!3:3")
            continue
        if parsed.span not in spans:
            forms = {"cells": "cells, as Stock!A1:H1", "columns": "whole columns, as Stock!D:F",
                     "rows": "whole rows, as Stock!2:5"}
            problems.add("bad_target", shown, "this takes " + " or ".join(
                forms[s] for s in spans if s in forms))
            continue
        try:
            sheet = workbook.sheet_of(parsed)
        except (NameNotFound, KeyError, ValueError):
            problems.add("bad_target", shown, "no such sheet; sheets: "
                         + ", ".join(s.name for s in workbook.sheets[:20]))
            continue
        if sheet.type != "worksheet" or sheet.part is None:
            problems.add("not_a_worksheet", shown, "a chart sheet holds no cells")
            continue
        out.append(Target(sheet, parsed.area, parsed.span, parsed.text))
    return out


#: What a protected sheet still lets be done, by the call's action (measured, Excel for Mac
#: 16.106, 2026-11-08, AppleScript on a sheet protected with Excel's defaults and on one
#: with ``formatCells``, ``formatColumns`` and ``formatRows`` set to ``0``): freezing panes
#: and hiding gridlines always (a window's view, not the sheet's content); formatting cells,
#: column widths and row heights only where the protection's flag of that name is ``0``
#: (with the defaults Excel refused bold and left widths and heights as they were).
#: ``None`` (the default) is refused on any protected sheet.
PROTECTED_ALLOWS = ("view", "formatCells", "formatColumns", "formatRows")


def protection_allows(workbook: Any, sheet: Any, action: "str | tuple[str, ...] | None") -> bool:
    """Whether ``sheet``'s protection lets ``action`` (one of :data:`PROTECTED_ALLOWS`, or
    several, all needed) be done; ``True`` for a sheet that is not protected."""
    if not sheet.protected:
        return True
    if action is None:
        return False
    if action == "view":
        return True
    root = workbook.package.tree(sheet.part)
    node = root.find(_X + "sheetProtection") if root is not None else None
    if node is None:
        return True
    needed = (action,) if isinstance(action, str) else action
    return all((node.get(name) or "1").strip().lower() in ("0", "false") for name in needed)


def guard(workbook: Any, found: list[Target], problems: Problems, *,
          whole_sheet: bool = False,
          protected: "str | tuple[str, ...] | None" = None) -> None:
    """A protected sheet, and the host's write policy: the cells (or, ``whole_sheet``, the
    whole sheet, for a sheet's settings) must be writable; ``protected`` the action a
    protected sheet may still allow (:data:`PROTECTED_ALLOWS`)."""
    from .policy import PolicyError

    try:
        allowed = workbook._allowed()
    except PolicyError as error:
        problems.add("bad_policy", "policy", error.message)
        return
    for item in found:
        if item.sheet.protected and not protection_allows(workbook, item.sheet, protected):
            allowed_by = (f" or protect it allowing {protected}" if isinstance(protected, str)
                          and protected != "view" else "")
            problems.add("protected", item.text, "unprotect the sheet first "
                         f"(wb.unprotect_sheet){allowed_by}")
            continue
        if allowed is None:
            continue
        area = Area(Cell(1, 1), Cell(MAX_ROW, MAX_COLUMN)) if whole_sheet else item.area
        if not _covered(allowed, item.sheet.part, area):
            problems.add("not_writable", item.text,
                         "the host's write policy does not cover " +
                         ("the whole sheet" if whole_sheet else "these cells"))


def _covered(allowed: Any, part: str, area: Area) -> bool:
    """Whether every policy lets every cell of ``area`` be written: one of its areas holds
    the whole area, or (for an area of at most 100,000 cells) each cell is held."""
    for areas in allowed._areas:
        held = areas.get(part, ())
        if any(a.contains(area.first) and a.contains(area.last) for a in held):
            continue
        if area.rows * area.columns > 100_000:
            return False
        if not all(any(a.contains(cell) for a in held) for cell in area.cells()):
            return False
    return True


# -- formatting --------------------------------------------------------------------------------------


def format_cells(workbook: Any, target: Any, **options: Any) -> Formatted:
    """Format the cells ``target`` names (see the module); one step."""
    from . import EditError

    if workbook.strict:
        raise EditError("strict")
    problems = Problems()
    found = targets(workbook, target, problems)
    editor = workbook.editor
    if editor.cells.styles is None:
        raise EditError("bad_format", "the workbook has no stylesheet to add formats to")
    table = StyleTable(workbook)
    with editor.batch():
        # Checked inside the step, as set_cells does: a refusal leaves nothing changed.
        spec = read_spec(workbook, problems, styles=table, **options)
        if spec.empty and not problems.found:
            problems.add("bad_format", "options", "nothing to change: give font, fill, "
                         "border, alignment, number_format, protection, like or style")
        needs = {"cells": ("formatCells",), "columns": ("formatCells", "formatColumns"),
                 "rows": ("formatCells", "formatRows")}
        for item in found:
            guard(workbook, [item], problems, protected=needs[item.span])
        plan = [(item, list(_cells_of(editor, item))) for item in found]
        total = sum(len(cells) for _, cells in plan)
        if total > MAX_CELLS:
            problems.add("too_many_cells", found[0].text if found else "target",
                         f"{total:,} cells; at most {MAX_CELLS:,} in one call: format whole "
                         "columns (Stock!D:D) or rows instead")
        problems.raise_any()
        cells: dict[str, set[tuple[int, int]]] = {}
        changed = 0
        columns: list[str] = []
        rows: list[str] = []
        memo: dict[tuple[int, frozenset[str]], int] = {}

        def made(base: int, edges: frozenset[str]) -> int:
            key = (base, edges if "outline" in spec.border else frozenset())
            if key not in memo:
                memo[key] = table.derive(base, spec, key[1])
            return memo[key]

        for item, positions in plan:
            writer = editor.cells.sheet(item.sheet.stored_name)
            label = _label(item.sheet)
            if item.span == "columns":
                columns.append(item.text)
                _column_styles(workbook, writer, item, made)
            elif item.span == "rows":
                rows.append(item.text)
                for number in range(item.area.first.row, item.area.last.row + 1):
                    _row_style(writer, number, made)
            for cell in positions:
                edges = _edges(item, cell)
                base = writer.style(cell, inherited=True)
                if writer.set_style(cell, made(base, edges)):
                    changed += 1
                cells.setdefault(label, set()).add((cell.row, cell.column))
            editor.touched_format(item.sheet)
        workbook._forget()
    listed, _, _ = capped(cells)
    return Formatted(cells=listed, changed=changed, formats_added=table.added,
                     columns=tuple(columns), rows=tuple(rows))


def _label(sheet: Any) -> str:
    from .structure import _label as label

    return label(sheet)


def _cells_of(editor: Any, item: Target) -> Iterable[Cell]:
    """The cells a target formats: every cell of a range; the cells a whole column or row
    holds."""
    if item.span == "cells":
        yield from item.area.cells()
        return
    yield from editor.cells.sheet(item.sheet.stored_name).cells_in(item.area)


def _edges(item: Target, cell: Cell) -> frozenset[str]:
    area = item.area
    found = set()
    if cell.row == area.first.row:
        found.add("top")
    if cell.row == area.last.row:
        found.add("bottom")
    if cell.column == area.first.column:
        found.add("left")
    if cell.column == area.last.column:
        found.add("right")
    return frozenset(found)


def _row_style(writer: Any, number: int, made: Any) -> None:
    writer.set_row_style(number, made(writer.row_style(number), frozenset()))


def _column_styles(workbook: Any, writer: Any, item: Target, made: Any) -> None:
    from .layout import Columns

    columns = Columns(workbook, writer)
    for number in range(item.area.first.column, item.area.last.column + 1):
        base = columns.get(number).get("style", "0")
        style = made(int(base) if base.isdigit() else 0, frozenset())
        columns.set(number, number, style=str(style) if style else None,
                    width_default=True)
    columns.write()
