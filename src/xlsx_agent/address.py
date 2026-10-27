"""Addresses: the forms an agent names cells, ranges, tables and datasets with.

Every result names what it returns by address, and every tool takes the same forms
(ROADMAP.md, "Addresses")::

    Stock!F2                     a cell
    Stock!A1:H40                 a range;  Stock!A:A and Stock!2:5 are whole columns and rows
    Stock!                       a whole sheet ('Q3 sales'!, #3!), clipped to its used range
    'Q3 sales'!B2:B9             a quoted sheet: ' around the name, '' for a quote inside
    #3!F2                        the sheet whose sheetId is 3 (survives a rename)
    t:Products                   a table, a defined name over a range, or a declared region
    t:Stock!Prices               a defined name local to a sheet (Excel's own spelling)
    t:Stock!A1                   a detected region, named by its header's first cell
    t:Products[Price]            one column of a table, header row excluded
    t:Products/SKU-0042/Price    the row whose key is SKU-0042, column Price
    t:Products/'A/7'/Price       a key holding / or ' is quoted, as a sheet name is
    t:Products/#5                the fifth data row (a position, not an id)
    t:Products/#9:#16            data rows 9 to 16, the last included: a row range;
                                 t:Products/#9: row 9 to the last
    t:Products/*/Price           every row's Price: a selector (Workbook.expand gives each
                                 value by key); t:Products/*/* every value, t:Products/*
                                 every row, t:Products/SKU-0042/* each value of one row
    t:Products/{SKU-0042,SKU-0043}/Price
                                 the rows whose keys are listed, column Price: a key set
    d:by_category                a derived dataset; d:by_category[average],
                                 d:by_category/Hardware/average as for tables
    $step1                       a ref: a name a call gave to what it made

:func:`parse_address` reads any of them into a :class:`RangeAddress`,
:class:`TableAddress`, :class:`DatasetAddress` or :class:`RefAddress`, each of which has a
canonical :attr:`text` it formats back to; a canonical address parses and formats back to
itself.  Cells and areas are ooxml-edit's :class:`~ooxml_edit.sheets.Cell` and
:class:`~ooxml_edit.sheets.Area`, on Excel's grid (``A1`` to ``XFD1048576``).

What parsing accepts beyond the canonical form, and normalises:

* ``$`` in a cell reference (``Stock!$A$1`` is ``Stock!A1``), lowercase columns, a range
  given corner to corner the other way round (``B2:A1`` is ``A1:B2``), and a one-cell range
  (``A1:A1`` is ``A1``);
* quotes around a sheet name that needs none (``'Stock'!A1`` is ``Stock!A1``), and an
  unquoted name that :func:`needs_quotes` would quote but that is unambiguous here (no
  space, quote, ``!`` or any of ``[ ] : * ? / \\``; not starting with ``#``):
  ``Überblick!A1`` is ``'Überblick'!A1``;
* white space around the whole address.

**The canonical form.**  Every address the library returns -- a :class:`~xlsx_agent.read.Value`'s
``a``, a problem's or an error's address, a find result, a range, a tool's output -- is in
one form, written by one formatter (:func:`format_address`, and :func:`format_sheet`,
:func:`format_cell` and :func:`format_range` beneath it); :func:`canonical_address` turns
any spelling parsing accepts into it, so a caller compares addresses it holds with those
returned without copying these rules::

    canonical_address("stock!$b$4")        # 'stock!B4'
    canonical_address("'Stock'!B4:A1")     # 'Stock!A1:B4'
    canonical_address("Überblick!a1")      # "'Überblick'!A1"
    canonical_address("stock!b4", sheets=["Stock"])   # 'Stock!B4': the workbook's spelling

* **Sheet names.**  Bare when the name is ASCII letters, digits, ``_`` and ``.``, starts
  with a letter or ``_``, and cannot be read as a cell reference (``A1``, ``XFD9``) or an
  R1C1 one (``R1C1``, ``R``, ``C2``) or a boolean (``TRUE``, ``FALSE``, case aside):
  ``Stock``, ``Q3_sales``.  Any other name is quoted, ``'`` before and after, each ``'``
  inside doubled: ``'Q3 sales'``, ``'Bob''s'``, ``'2024'``, ``'A1'``, ``'R1C1'``,
  ``'Überblick'`` (:func:`needs_quotes`, :func:`quote_sheet`).  The name is spelled as the
  workbook spells it in what the library returns; :func:`canonical_address` keeps the
  spelling it is given unless it is given the workbook's names (``sheets``), as it does
  not know the workbook (the workbook matches sheet names case aside).  A sheet by id is
  ``#`` and its ``sheetId``, without leading zeros: ``#3``.
* **Cells.**  Column letters upper case, no ``$``, the row without leading zeros:
  ``B4``.
* **Ranges.**  First corner (top left) ``:`` last (bottom right): ``A1:B4``; a range of
  one cell is the cell (``B4``); whole columns ``A:C`` and rows ``2:5`` (one column or row
  too: ``A:A``, ``7:7``); a whole sheet is its sheet and ``!`` (``Stock!``).
* **Tables, names, regions and datasets** (``t:``, ``d:``): the name or id as given (the
  workbook's spelling, in what the library returns); a region's header cell as a cell
  above; a key or column quoted only when it must be (see below).
* **Refs**: ``$`` and the name.

Quoting is conservative, and is this library's rule, not a measurement of Excel: Excel
accepts a quoted name wherever it accepts a bare one, so the canonical form is always one
Excel reads, and pastes into a formula unchanged; but which further names Excel itself
leaves bare in the formulas it writes (names of non-ASCII letters, say) has not been
measured (the formula lexer, E2, is where it will be), so a formula's own text may write
bare a sheet this form quotes.  Compare through :func:`canonical_address`, never the
formula's text.

**Sheet names are not policed here.**  ECMA-376's ``sheet@name`` is an ``ST_Xstring``
whose only constraint is that it be unique in the workbook (ISO/IEC 29500-4, ``sml.xsd``,
``CT_Sheet``), so an address must be able to name whatever a file holds: any non-empty
name, quoted.  The characters Excel refuses in a sheet name and its length limit are
Excel's own rules, not the schema's; they matter when a sheet is added or renamed (E5),
and are measured from Excel then.

**Table, name and region names** (ECMA-376 Part 1, §18.17, "Names"; a table's
``displayName``, §18.5.1.2, has "the same character lengths and restrictions" as a defined
name).  The standard's grammar: a name starts with a letter, ``_`` or ``\\``, and goes on
with letters, the digits 0-9, ``_`` and ``.``; it is not ``TRUE`` or ``FALSE`` and not a
cell reference; names are not case-sensitive and are unique in the workbook; an
implementation "is encouraged to support" at least 255 characters.  :func:`is_name` is
that rule, and a name this library makes (:meth:`Workbook.declare_region
<xlsx_agent.workbook.Workbook.declare_region>`) must pass it.  In an address the rule is
wider, because a ``name`` attribute is an ST_Xstring and a file may hold a name the grammar
does not allow (``Products?``, say): a ``t:`` name is any non-empty text without white
space, without the characters the address itself uses (``[ ] / ! '``) and not starting
with ``#``.  A name outside that is in the workbook but has no ``t:`` address.

**Dataset ids** (``d:``) are ooxml-edit's (O2, :data:`ooxml_edit.tools.DATASET_ID`): a
letter or ``_``, then letters, digits, ``_``, ``.`` and ``-``; at most 64 characters; case
counts.  The session keeps datasets by that id, so an address takes no other.

**Sheet-scoped names.**  A defined name local to a sheet is written as Excel writes it,
sheet first: ``t:Stock!Prices``, ``t:'Q3 sales'!Rate``, ``t:#3!Rate``.  What follows the
``!`` is a region's header cell when it reads as one (``A1``; with ``:`` or ``$`` it is
taken as a cell or range, and refused if it is not one cell) and a name otherwise: a name
is never a cell reference.  A sheet-local name that looks like a cell past the grid
(``XFE1``, which the standard allows) has no address in this form.

**Whole sheets.**  A sheet and its ``!`` with nothing after it (``Stock!``, ``'Q3 sales'!``,
``#3!``) is the whole sheet: a :class:`RangeAddress` whose ``span`` is ``"sheet"`` and
whose area is the whole grid, clipped to the used range when read, as whole columns and
rows are.  It keeps R2's grammar -- a sheet, ``!``, then where on it -- with nothing for
"everywhere" (``Stock!A1:XFD1048576`` names the same cells, as a rectangle).
``t:Stock!`` is not a form: a detected region is named by a cell.

**Positions and ids.**  ``#3!`` names a sheet by id; ``/#5`` names a row by position.  A
key that begins with ``#`` is quoted, so the two never meet.

**Row ranges.**  ``/#9:#16`` names data rows 9 to 16 by position (from 1, as ``#5``; the
last included), ``/#9:`` row 9 to the last: a run of rows, in the table's order, as a key
set names rows by key (``/#9:#16/Price`` their one column, ``/#9:#16/*`` every column).  A
range given the other way round (``#16:#9``) is the same rows, and one of one row
(``#9:#9``) is that row (``#9``); ``#9:16`` is read as ``#9:#16``.  A row past the last is refused when the address is
resolved, never cut (``position_out_of_range``, the range there is as an option).

**Selectors and key sets.**  ``*`` in a row's place is every data row, in the table's order
when it is expanded (``t:Products/*/Price``: the cells of ``t:Products[Price]``;
``t:Products/*``: the table); ``*`` in a column's place is every column
(``t:Products/SKU-0042/*``: the row).  A selector reads as the cells it covers, and
:meth:`Workbook.expand <xlsx_agent.workbook.Workbook.expand>` gives each of them by key
(``t:Products/SKU-0042/Price``...), so that a caller can check that a set of addresses
covers every one.  A key set, ``{`` keys ``,`` ... ``}``, names the rows whose keys are
listed, in that order: what a derived value's sources are written as when its table has a
key column (:mod:`xlsx_agent.derive`).  A key that is ``*``, or begins with ``{``, is
quoted (``'*'``); in a set, so is a key holding ``,`` ``{`` or ``}``.  A set of one key is
that key; a key listed twice is listed once.  ``#5`` has no place in a set: a position is
not an id.

**No cell content in diagnostics** (ROADMAP.md, "Design rules", 7).  A key is cell
content, so an :class:`AddressError`'s message names the problem and where it is in the
address (a 1-based character position), never the address's text; sheet and table names
may appear.  The text is kept on the exception, for a program.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, replace
from typing import Iterable, Mapping, NoReturn, Union

from ooxml_edit.sheets import (MAX_COLUMN, MAX_ROW, Area, Cell, column_index,
                                        column_letters)
from ooxml_edit.tools import DATASET_ID, REF

#: One example of every form, by kind: what an :class:`AddressError` offers when the
#: address is none of them.
FORMS: dict[str, str] = {
    "cell": "Stock!F2",
    "range": "Stock!A1:H40",
    "columns": "Stock!A:C",
    "rows": "Stock!2:5",
    "sheet": "Stock!",
    "quoted sheet": "'Q3 sales'!B2:B9",
    "sheet by id": "#3!F2",
    "table": "t:Products",
    "sheet-scoped name": "t:Stock!Prices",
    "detected region": "t:Stock!A1",
    "column": "t:Products[Price]",
    "value by key": "t:Products/SKU-0042/Price",
    "row by position": "t:Products/#5",
    "row range": "t:Products/#9:#16",
    "every row": "t:Products/*/Price",
    "key set": "t:Products/{SKU-0042,SKU-0043}/Price",
    "dataset": "d:by_category",
    "ref": "$step1",
}

#: ``sheetId`` is an ``xsd:unsignedInt`` (ISO/IEC 29500-4, ``sml.xsd``, ``CT_Sheet``).
MAX_SHEET_ID = 2**32 - 1

#: The address prefixes of a table (or name, or region) and of a dataset.
TABLE_PREFIX = "t:"
DATASET_PREFIX = "d:"

#: How a :class:`RangeAddress` spans the grid: a rectangle of cells, whole columns
#: (``A:C``), whole rows (``2:5``), or the whole sheet (``Stock!``).
SPANS = ("cells", "columns", "rows", "sheet")

#: The whole grid: what a whole sheet's area is.
WHOLE_GRID = Area(Cell(1, 1), Cell(MAX_ROW, MAX_COLUMN))

_PLAIN_SHEET = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_CELL_LIKE = re.compile(r"[A-Za-z]{1,3}[0-9]+")
_R1C1_LIKE = re.compile(r"(?:[Rr][0-9]*(?:[Cc][0-9]*)?|[Cc][0-9]*)")
_BOOLEANS = ("TRUE", "FALSE")

# Characters a bare sheet name may not hold: they delimit, or Excel never allows them.
_NOT_IN_BARE_SHEET = frozenset("'![]:*?/\\")
# Characters a table, name, region or dataset id may not hold in an address: the address
# uses them (see the module's notes).
_NOT_IN_NAME = frozenset("[]/!'")

#: The most characters of a name this library makes: the length ECMA-376 encourages an
#: implementation to support (§18.17, "Names").
MAX_NAME = 255

_CELLS = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?::\$?([A-Za-z]{1,3})\$?([0-9]{1,7}))?")
_COLUMNS = re.compile(r"\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})")
_ROWS = re.compile(r"\$?([0-9]{1,7}):\$?([0-9]{1,7})")
_POSITION = re.compile(r"#([1-9][0-9]{0,6})")
_ROW_RANGE = re.compile(r"#([1-9][0-9]{0,6}):(?:#?([1-9][0-9]{0,6}))?")

_RANGE_FORMS = ("cell", "range", "columns", "rows", "sheet", "quoted sheet", "sheet by id")
_TABLE_FORMS = ("table", "sheet-scoped name", "detected region", "column", "value by key",
                "row by position", "row range", "every row", "key set")
_DATASET_ID_RULE = ("a dataset id is a letter or _, then letters, digits, _ . or -; at most "
                    "64 characters")
_DATASET_FORMS = ("d:by_category", "d:by_category[average]", "d:by_category/Hardware/average",
                  "d:by_category/#2", "d:by_category/#2:#5", "d:by_category/*/average")

#: The wildcard: every row, in a row's place; every column, in a column's.
ANY = "*"


# ------------------------------------------------------------------------------------------
# Errors
# ------------------------------------------------------------------------------------------


class AddressError(ValueError):
    """An address that is none of the forms, or a form with something wrong in it.

    The message says what is wrong and at which character (1-based), never the address's
    text: a key is cell content.  ``text`` is the address as given, ``position`` the
    character (``None`` when the address as a whole is wrong), and ``valid_options`` what
    would work instead -- a corrected address when there is one, else an example of each
    form the place takes (:data:`FORMS`)::

        try:
            parse_address("t:Stock!A1:F20")
        except AddressError as error:
            error.valid_options      # ['t:Stock!A1']
    """

    def __init__(self, message: str, *, text: str = "", position: int | None = None,
                 valid_options: Iterable[str] = ()) -> None:
        super().__init__(message)
        self.message = message
        self.text = text
        self.position = position
        self.valid_options = list(valid_options)


class NameNotFound(KeyError):
    """A sheet, sheet id or ref that names nothing in the workbook (or the session).

    ``what`` is ``"sheet"``, ``"sheet id"`` or ``"ref"``; ``name`` what was asked for;
    ``candidates`` every name there is, as an address writes it (``'Q3 sales'``, ``#3``,
    ``$step1``); ``valid_options`` the nearest of them (matched case aside), or all of
    them, up to 50, when none is near.  A :class:`KeyError`, as an unknown sheet always
    was; ``str()`` is the message without :class:`KeyError`'s quotes.
    """

    def __init__(self, message: str, *, what: str, name: object,
                 candidates: Iterable[str] = (), valid_options: Iterable[str] = ()) -> None:
        super().__init__(message)
        self.message = message
        self.what = what
        self.name = name
        self.candidates = list(candidates)
        self.valid_options = list(valid_options)

    def __str__(self) -> str:
        return self.message


def nearest(name: str, candidates: Iterable[str], *, limit: int = 50) -> list[str]:
    """The candidates nearest ``name``, matched case aside (at most five); every candidate
    (up to ``limit``) when none is near."""
    candidates = list(candidates)
    folded: dict[str, str] = {}
    for candidate in candidates:
        folded.setdefault(candidate.casefold(), candidate)
    close = difflib.get_close_matches(name.casefold(), list(folded), 5, 0.5)
    return [folded[match] for match in close] if close else candidates[:limit]


# ------------------------------------------------------------------------------------------
# Names
# ------------------------------------------------------------------------------------------


def is_name(name: object) -> bool:
    """Whether ``name`` is a name as ECMA-376 Part 1, §18.17 ("Names") defines one: a
    letter, ``_`` or ``\\`` first; then letters, the digits 0-9, ``_`` and ``.``; not
    ``TRUE`` or ``FALSE``; not a cell reference on the grid (``A1`` to ``XFD1048576``) nor
    an R1C1 one (``R1C1``, ``R``, ``C2``); at most :data:`MAX_NAME` characters.  A
    "letter" is a Unicode letter.  What a name this library makes must be; a file's own
    names may be more (see the module's notes)."""
    if not isinstance(name, str) or not name or len(name) > MAX_NAME:
        return False
    first = name[0]
    if not (first.isalpha() or first in "_\\"):
        return False
    if not all(ch.isalpha() or ch in "0123456789_." for ch in name[1:]):
        return False
    if name.upper() in _BOOLEANS or _R1C1_LIKE.fullmatch(name):
        return False
    return not _on_grid(name)


def is_address_name(name: object) -> bool:
    """Whether ``name`` can stand in a ``t:`` address (the wider rule of the module's
    notes): non-empty, no white space or control character, none of ``[ ] / ! '``, not
    starting with ``#``."""
    return (isinstance(name, str) and bool(name) and not name.startswith("#")
            and not any(ch.isspace() or ord(ch) < 32 or ch in _NOT_IN_NAME for ch in name))


def _on_grid(text: str) -> bool:
    """``text`` is an A1 reference to a cell on Excel's grid (``$`` allowed)."""
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})", text)
    return match is not None and column_index(match.group(1)) <= MAX_COLUMN \
        and 1 <= int(match.group(2)) <= MAX_ROW


# ------------------------------------------------------------------------------------------
# Sheet names
# ------------------------------------------------------------------------------------------


def needs_quotes(name: str) -> bool:
    """Whether a sheet name is quoted in an address: anything but ASCII letters, digits,
    ``_`` and ``.`` starting with a letter or ``_``, and anything that reads as a cell
    reference (``A1``, ``XFD9``, ``R1C1``, ``R``, ``C2``) or a boolean."""
    if not _PLAIN_SHEET.fullmatch(name):
        return True
    return bool(_CELL_LIKE.fullmatch(name) or _R1C1_LIKE.fullmatch(name)
                or name.upper() in _BOOLEANS)


def quote_sheet(name: str) -> str:
    """A sheet name as an address writes it: ``Stock``, ``'Q3 sales'``, ``'Bob''s'``."""
    if not name:
        raise ValueError("a sheet name is never empty")
    if needs_quotes(name):
        return "'" + name.replace("'", "''") + "'"
    return name


def format_sheet(name: str | None, id: int | None = None) -> str:
    """A sheet as an address writes it: its ``name`` quoted when it must be
    (:func:`quote_sheet`), or ``#`` and its ``id`` when it has no name: ``Stock``,
    ``'Q3 sales'``, ``#3``."""
    if name:
        return quote_sheet(name)
    if id is None:
        raise ValueError("a sheet is written by its name or its id")
    return f"#{id}"


_LETTERS: dict[int, str] = {}


def _letters(column: int) -> str:
    """A column's letters (1 is ``A``), remembered: what every cell address is written
    with."""
    letters = _LETTERS.get(column)
    if letters is None:
        number, out = column, ""
        while number:
            number, rest = divmod(number - 1, 26)
            out = chr(65 + rest) + out
        letters = _LETTERS[column] = out
    return letters


def _sheet_text(sheet: "SheetRef | str") -> str:
    return sheet if isinstance(sheet, str) else sheet.text


def format_cell(sheet: "SheetRef | str", row: int, column: int) -> str:
    """A cell's address: ``Stock!B4``.  ``sheet`` is a :class:`SheetRef`, or a sheet as an
    address writes it (:func:`format_sheet`); ``row`` and ``column`` count from 1."""
    return f"{_sheet_text(sheet)}!{_letters(column)}{row}"


def format_range(sheet: "SheetRef | str", area: "Area | Cell | None" = None,
                 span: str = "cells") -> str:
    """A cell, range or whole sheet's address, as :attr:`RangeAddress.text` writes it:
    ``Stock!A1:B4``, ``Stock!B4`` (a range of one cell), ``Stock!A:C`` and ``Stock!2:5``
    (``span`` ``"columns"`` or ``"rows"``: the area's columns or rows), ``Stock!`` (no
    ``area``, or ``span`` ``"sheet"``).  ``sheet`` as for :func:`format_cell`."""
    if isinstance(area, Cell):
        area = Area(area, area)
    if area is None or span == "sheet":
        where = ""
    elif span == "columns":
        where = f"{_letters(area.first.column)}:{_letters(area.last.column)}"
    elif span == "rows":
        where = f"{area.first.row}:{area.last.row}"
    elif span == "cells":
        first, last = area.first, area.last
        where = f"{_letters(first.column)}{first.row}"
        if last != first:
            where += f":{_letters(last.column)}{last.row}"
    else:
        raise ValueError(f"span is one of {SPANS}")
    return f"{_sheet_text(sheet)}!{where}"


@dataclass(frozen=True)
class SheetRef:
    """Which sheet: by ``name`` (as the workbook spells it) or by ``id`` (its ``sheetId``),
    exactly one of the two.  :meth:`Workbook.sheet_of
    <xlsx_agent.workbook.Workbook.sheet_of>` finds the sheet."""

    name: str | None = None
    id: int | None = None

    def __post_init__(self) -> None:
        if (self.name is None) == (self.id is None):
            raise ValueError("a sheet is named by name or by id, one of the two")
        if self.name is not None and (not isinstance(self.name, str) or not self.name):
            raise ValueError("a sheet name is a non-empty string")
        if self.id is not None and (isinstance(self.id, bool) or not isinstance(self.id, int)
                                    or not 0 <= self.id <= MAX_SHEET_ID):
            raise ValueError(f"a sheet id is a whole number from 0 to {MAX_SHEET_ID}")

    @property
    def text(self) -> str:
        """``Stock``, ``'Q3 sales'`` or ``#3``."""
        return f"#{self.id}" if self.id is not None else format_sheet(self.name)

    def __str__(self) -> str:
        return self.text


# ------------------------------------------------------------------------------------------
# The addresses
# ------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class RangeAddress:
    """A cell, a range or a whole sheet: ``Stock!F2``, ``Stock!A1:H40``, ``#3!A:A``,
    ``Stock!``.

    ``span`` is ``"cells"``, ``"columns"`` (``A:C``: ``area`` runs from row 1 to
    :data:`~ooxml_edit.sheets.MAX_ROW`), ``"rows"`` (``2:5``: column 1 to
    :data:`~ooxml_edit.sheets.MAX_COLUMN`) or ``"sheet"`` (``Stock!``: ``area`` is
    :data:`WHOLE_GRID`).  Whole columns, rows and sheets are clipped to the sheet's used
    range when read (:meth:`clipped`)."""

    sheet: SheetRef
    area: Area
    span: str = "cells"

    def __post_init__(self) -> None:
        if self.span not in SPANS:
            raise ValueError(f"span is one of {SPANS}")
        first, last = self.area.first, self.area.last
        if not (1 <= first.row <= last.row <= MAX_ROW
                and 1 <= first.column <= last.column <= MAX_COLUMN):
            raise ValueError("the area is not a rectangle on the sheet")
        if self.span == "columns" and (first.row, last.row) != (1, MAX_ROW):
            raise ValueError("whole columns run from row 1 to the last row")
        if self.span == "rows" and (first.column, last.column) != (1, MAX_COLUMN):
            raise ValueError("whole rows run from column A to the last column")
        if self.span == "sheet" and self.area != WHOLE_GRID:
            raise ValueError("a whole sheet is the whole grid, A1 to XFD1048576")

    @property
    def kind(self) -> str:
        """``"cell"``, ``"range"`` or ``"sheet"`` (a whole sheet)."""
        if self.span == "sheet":
            return "sheet"
        if self.span == "cells" and self.area.first == self.area.last:
            return "cell"
        return "range"

    @property
    def cell(self) -> Cell:
        """The cell of a one-cell address."""
        if self.kind != "cell":
            raise ValueError("a range is not one cell")
        return self.area.first

    @property
    def text(self) -> str:
        """The canonical address (see the module's notes): :func:`format_range`."""
        return format_range(self.sheet, self.area, self.span)

    def __str__(self) -> str:
        return self.text

    def on(self, sheet: SheetRef) -> "RangeAddress":
        """The same cells, with the sheet named as ``sheet`` names it (by id or by name)."""
        return RangeAddress(sheet, self.area, self.span)

    def clipped(self, used: Area | None) -> Area | None:
        """The cells meant, given the sheet's used range: whole columns, rows and sheets
        clipped to it (``None`` when they miss it, or the sheet is empty); a rectangle of cells as it
        is, used range or not."""
        if self.span == "cells":
            return self.area
        if used is None:
            return None
        first = Cell(max(self.area.first.row, used.first.row),
                     max(self.area.first.column, used.first.column))
        last = Cell(min(self.area.last.row, used.last.row),
                    min(self.area.last.column, used.last.column))
        if first.row > last.row or first.column > last.column:
            return None
        return Area(first, last)

    @classmethod
    def parse(cls, text: str) -> "RangeAddress":
        """``text`` as a cell, range or whole sheet; :class:`AddressError` for any other
        form."""
        return _expect(text, cls, _RANGE_FORMS)


class _Selection:
    """What table and dataset addresses share: a column, a row (by key or by position),
    or a value (a row and a column); or a selector (``*`` for every row, or every column),
    a key set (``keys``: several rows by key) or a row range (``row_range``: the first and
    last rows by position, from 1, the last ``None`` for to the last row)."""

    column: str | None
    key: str | None
    position: int | None
    keys: tuple[str, ...] | None
    any_row: bool
    any_column: bool
    row_range: tuple[int, int | None] | None

    def _check_selection(self) -> None:
        rows = [self.key is not None, self.position is not None, self.keys is not None,
                bool(self.any_row), self.row_range is not None]
        if sum(rows) > 1:
            raise ValueError("a row is named by key, by position, by a key set, by a row "
                             "range or by *, one of them")
        if self.row_range is not None:
            found = self.row_range
            if not isinstance(found, tuple) or len(found) != 2 or not all(
                    isinstance(n, int) and not isinstance(n, bool) and 1 <= n <= MAX_ROW
                    for n in (found[0], found[0] if found[1] is None else found[1])):
                raise ValueError(f"a row range is (first, last), whole numbers from 1 to "
                                 f"{MAX_ROW}, last None for to the last row")
            if found[1] is not None and found[1] <= found[0]:
                raise ValueError("a row range runs from its first row to a later last (one "
                                 "row is written as a position)")
        if self.any_column and (self.column is not None or not any(rows)):
            raise ValueError("* for every column follows a row (t:Products/SKU-0042/*), "
                             "in place of a column")
        if self.keys is not None:
            if isinstance(self.keys, (str, bytes)) or len(self.keys) < 2 or not all(
                    isinstance(k, str) and k for k in self.keys):
                raise ValueError("a key set holds two or more non-empty keys (one key is "
                                 "written as a key)")
            if len(set(self.keys)) != len(self.keys):
                raise ValueError("a key set lists each key once")
        if self.column is not None and (not isinstance(self.column, str) or not self.column):
            raise ValueError("a column name is a non-empty string")
        if self.key is not None and (not isinstance(self.key, str) or not self.key):
            raise ValueError("a key is a non-empty string")
        if self.position is not None and (isinstance(self.position, bool)
                                          or not isinstance(self.position, int)
                                          or not 1 <= self.position <= MAX_ROW):
            raise ValueError(f"a row position is a whole number from 1 to {MAX_ROW}")

    @property
    def _selected(self) -> str | None:
        if self.keys is not None or self.row_range is not None:
            return "values" if self.column is not None else "rows"
        if self.key is not None or self.position is not None:
            return "value" if self.column is not None else "row"
        return "column" if self.column is not None else None

    @property
    def is_whole(self) -> bool:
        """Names all of the table (or dataset): no column or row, or ``*`` for both
        (``t:Products``, ``t:Products/*``, ``t:Products/*/*``)."""
        return self._selected is None

    @property
    def is_selector(self) -> bool:
        """Holds a ``*`` (``t:Products/*/Price``, ``t:Products/SKU-0042/*``)."""
        return bool(self.any_row or self.any_column)

    def _selection_text(self) -> str:
        has_row = self.key is not None or self.position is not None or \
            self.keys is not None or self.any_row or self.row_range is not None
        if not has_row:
            return "" if self.column is None else f"[{_escape_column(self.column)}]"
        if self.any_row:
            row = ANY
        elif self.position is not None:
            row = f"#{self.position}"
        elif self.row_range is not None:
            first, last = self.row_range
            row = f"#{first}:" + ("" if last is None else f"#{last}")
        elif self.keys is not None:
            row = "{" + ",".join(_set_item(k) for k in self.keys) + "}"
        else:
            assert self.key is not None   # a row named by key: the one form left
            row = _segment(self.key, row=True)
        if self.any_column:
            column = "/" + ANY
        else:
            column = "" if self.column is None else "/" + _segment(self.column, row=False)
        return f"/{row}{column}"


@dataclass(frozen=True)
class TableAddress(_Selection):
    """A table, a defined name over a range or a declared region (``t:Products``), a
    defined name local to a sheet (``t:Stock!Prices``: ``scope`` is the sheet), or a
    detected region (``t:Stock!A1``: ``region`` is its header's first cell, ``name`` is
    ``None``); with, optionally, a column (``t:Products[Price]``), a row by key
    (``t:Products/SKU-0042``) or position (``t:Products/#5``), a row range
    (``t:Products/#9:#16``), or a value (``t:Products/SKU-0042/Price``).

    Parsing does not say whether the name is a table, a name or a region, nor whether the
    key and column exist: :meth:`Workbook.resolve <xlsx_agent.workbook.Workbook.resolve>`
    does."""

    name: str | None = None
    region: RangeAddress | None = None
    column: str | None = None
    key: str | None = None
    position: int | None = None
    scope: SheetRef | None = None
    keys: tuple[str, ...] | None = None
    any_row: bool = False
    any_column: bool = False
    row_range: tuple[int, int | None] | None = None

    def __post_init__(self) -> None:
        if (self.name is None) == (self.region is None):
            raise ValueError("a table address has a name or a region, one of the two")
        if self.name is not None:
            _check_name(self.name, "a table name")
            if self.scope is not None and _CELLS.fullmatch(self.name):
                raise ValueError("a sheet-scoped name is never a cell reference")
        elif self.scope is not None:
            raise ValueError("a detected region names its sheet in region, not in scope")
        if self.region is not None and self.region.kind != "cell":
            raise ValueError("a detected region is named by one cell, its header's first")
        self._check_selection()

    @property
    def kind(self) -> str:
        """``"table"`` or ``"region"`` (the whole of it), ``"column"``, ``"row"`` or
        ``"value"``; ``"rows"`` or ``"values"`` for a key set or a row range (with a
        column: values).  A
        selector is what it covers: ``t:Products/*/Price`` a column, ``t:Products/*`` the
        table, ``t:Products/SKU-0042/*`` a row."""
        return self._selected or ("region" if self.region is not None else "table")

    @property
    def is_region(self) -> bool:
        """Named by a header cell (``t:Stock!A1``), not by a name."""
        return self.region is not None

    @property
    def text(self) -> str:
        if self.name is None:
            assert self.region is not None   # a region's, or a name's (__post_init__)
            base = self.region.text
        elif self.scope is not None:
            base = f"{self.scope.text}!{self.name}"
        else:
            base = self.name
        return TABLE_PREFIX + base + self._selection_text()

    @property
    def whole(self) -> "TableAddress":
        """The table, name or region itself, without a column, row or selector."""
        return TableAddress(name=self.name, region=self.region, scope=self.scope)

    def __str__(self) -> str:
        return self.text

    @classmethod
    def parse(cls, text: str) -> "TableAddress":
        """``text`` as a ``t:`` address; :class:`AddressError` for any other form."""
        return _expect(text, cls, _TABLE_FORMS)


@dataclass(frozen=True)
class DatasetAddress(_Selection):
    """A derived dataset in the session (``d:by_category``), or a column, row or value of
    one, written as for tables: ``d:by_category[average]``,
    ``d:by_category/Hardware/average``, ``d:by_category/#2``, ``d:by_category/#2:#5``."""

    id: str
    column: str | None = None
    key: str | None = None
    position: int | None = None
    keys: tuple[str, ...] | None = None
    any_row: bool = False
    any_column: bool = False
    row_range: tuple[int, int | None] | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.id, str) or not DATASET_ID.fullmatch(self.id):
            raise ValueError(_DATASET_ID_RULE)
        self._check_selection()

    @property
    def kind(self) -> str:
        """``"dataset"`` (the whole of it), ``"column"``, ``"row"``, ``"value"``, or
        ``"rows"``/``"values"`` for a key set or a row range; a selector is what it
        covers, as for tables."""
        return self._selected or "dataset"

    @property
    def text(self) -> str:
        return DATASET_PREFIX + self.id + self._selection_text()

    def __str__(self) -> str:
        return self.text

    @classmethod
    def parse(cls, text: str) -> "DatasetAddress":
        """``text`` as a ``d:`` address; :class:`AddressError` for any other form."""
        return _expect(text, cls, ("dataset",))


@dataclass(frozen=True)
class RefAddress:
    """A ref, ``$step1``: a name an earlier call gave to an address it made (ooxml-edit's
    refs, :data:`ooxml_edit.tools.REF`).  ``rest`` is what follows a ``/`` after the name,
    as written (``$step1/SKU-0042/Price``); the tool layer puts the ref's address in its
    place before a handler sees it, and :meth:`resolve` does the same for a script."""

    name: str
    rest: str = ""

    def __post_init__(self) -> None:
        if not isinstance(self.name, str) or not REF.match(self.name):
            raise ValueError(f"a ref's name matches {REF.pattern}")

    @property
    def kind(self) -> str:
        return "ref"

    @property
    def text(self) -> str:
        return f"${self.name}" + (f"/{self.rest}" if self.rest else "")

    def __str__(self) -> str:
        return self.text

    def resolve(self, refs: Mapping[str, str]) -> "Address":
        """The address the ref stands for, with ``rest`` after it; :class:`NameNotFound`,
        with the known refs as ``valid_options``, when it names nothing."""
        if self.name not in refs:
            known = [f"${name}" for name in sorted(refs)]
            raise NameNotFound(f"no ref ${self.name}; the refs are {known}", what="ref",
                               name=self.name, candidates=known,
                               valid_options=nearest(f"${self.name}", known))
        return parse_address(refs[self.name] + (f"/{self.rest}" if self.rest else ""))

    @classmethod
    def parse(cls, text: str) -> "RefAddress":
        """``text`` as a ref; :class:`AddressError` for any other form."""
        return _expect(text, cls, ("ref",))


#: Any address :func:`parse_address` returns.
Address = Union[RangeAddress, TableAddress, DatasetAddress, RefAddress]


# ------------------------------------------------------------------------------------------
# Parsing
# ------------------------------------------------------------------------------------------


def parse_address(text: str) -> Address:
    """Read an address in any of the forms; :class:`AddressError` with the reason, and what
    would work, when it is none of them."""
    if not isinstance(text, str):
        raise AddressError("an address is a string", valid_options=FORMS.values())
    return _Parser(text).address()


def format_address(address: Address) -> str:
    """An address's canonical text (its :attr:`text`; see the module's notes)."""
    return address.text


def canonical_address(text: str, *, sheets: Iterable[str] | None = None) -> str:
    """Any spelling of an address that parsing accepts, in the canonical form every
    address the library returns is in (see the module's notes): ``"stock!$b$4"`` is
    ``"stock!B4"``, ``"'Stock'!B4:A1"`` is ``"Stock!A1:B4"``, ``"Überblick!a1"`` is
    ``"'Überblick'!A1"``.  ``sheets``, the workbook's sheet names (``[s.name for s in
    wb.sheets]``), spells a sheet named in the address as the workbook does, matched case
    aside as the workbook matches them; a name not among them is kept as given.  A
    canonical address is its own canonical form; :class:`AddressError` for text that is
    not an address."""
    address = parse_address(text)
    if sheets is None:
        return address.text
    spelled: dict[str, str] = {}
    for name in sheets:
        spelled.setdefault(name.casefold(), name)

    def respell(sheet: SheetRef) -> SheetRef:
        if sheet.name is None:
            return sheet
        return SheetRef(name=spelled.get(sheet.name.casefold(), sheet.name))

    if isinstance(address, RangeAddress):
        address = address.on(respell(address.sheet))
    elif isinstance(address, TableAddress):
        if address.region is not None:
            address = replace(address, region=address.region.on(respell(address.region.sheet)))
        elif address.scope is not None:
            address = replace(address, scope=respell(address.scope))
    return address.text


class _Parser:
    """A hand-written scan, one pass, no backtracking: an address may be long (a key can
    be any cell's text), and is untrusted."""

    def __init__(self, text: str) -> None:
        self.text = text
        self.body = text.strip()
        # Positions in messages count from the address as given.
        self.offset = len(text) - len(text.lstrip())
        # What an error offers when it has no correction: the forms of the place it is in.
        self.forms: list[str] = list(FORMS.values())

    def fail(self, message: str, at: int | None = None,
             options: Iterable[str] | None = None) -> NoReturn:
        position = None if at is None else self.offset + at + 1
        if position is not None:
            message = f"{message} (at character {position})"
        raise AddressError(message, text=self.text, position=position,
                           valid_options=self.forms if options is None else options)

    # -- dispatch ----------------------------------------------------------------------

    def address(self) -> Address:
        body = self.body
        if not body:
            self.fail("an address is empty")
        if body.startswith("$"):
            return self.ref()
        if body.startswith(TABLE_PREFIX):
            return self.table()
        if body.startswith(DATASET_PREFIX):
            return self.dataset()
        return self.sheet_range(0)

    # -- sheets and ranges -------------------------------------------------------------

    def sheet(self, at: int) -> tuple[SheetRef, int]:
        """The sheet starting at ``at``, and the index of the ``!`` after it."""
        body = self.body
        if body.startswith("'", at):
            name, i = [], at + 1
            while True:
                j = body.find("'", i)
                if j < 0:
                    self.fail("a quoted sheet name has no closing quote", at)
                name.append(body[i:j])
                if body.startswith("''", j):
                    name.append("'")
                    i = j + 2
                    continue
                i = j + 1
                break
            if i == at + 2:
                self.fail("a sheet name is never empty", at)
            if not body.startswith("!", i):
                self.fail("a quoted sheet name is followed by ! and a cell or range", i)
            return SheetRef(name="".join(name)), i
        bang = body.find("!", at)
        if bang < 0:
            self.fail_no_sheet(at)
        token = body[at:bang]
        if not token:
            self.fail("! follows a sheet name", at)
        if token.startswith("#"):
            digits = token[1:]
            if not (digits.isascii() and digits.isdigit()) or (len(digits) > 1 and digits[0] == "0"):
                self.fail("a sheet id is # and a whole number, as in #3!F2", at)
            if int(digits) > MAX_SHEET_ID:
                self.fail(f"a sheet id is at most {MAX_SHEET_ID}", at)
            return SheetRef(id=int(digits)), bang
        for k, ch in enumerate(token):
            if ch.isspace() or ch in _NOT_IN_BARE_SHEET or ord(ch) < 32:
                # The corrected address, for a plain range only: after a region's cell
                # may come a key, which is cell content.
                options = [quote_sheet(token) + body[bang:]] if at == 0 else None
                self.fail("a sheet name with spaces or punctuation is quoted: "
                          "'Q3 sales'!B2, 'Bob''s'!A1", at + k, options)
        return SheetRef(name=token), bang

    def fail_no_sheet(self, at: int) -> NoReturn:
        rest = self.body[at:]
        if at:
            self.fail("a detected region names its sheet and its header's first cell: "
                      "t:Stock!A1 or t:#3!A1", at)
        if _CELLS.fullmatch(rest) or _COLUMNS.fullmatch(rest) or _ROWS.fullmatch(rest):
            self.fail("a cell or range names its sheet: Stock!A1, 'Q3 sales'!A1 or #3!A1",
                      options=[FORMS["cell"], FORMS["quoted sheet"], FORMS["sheet by id"]])
        if at == 0 and rest and not any(ch in _NOT_IN_NAME or ch.isspace() for ch in rest) \
                and not rest.startswith("#"):
            self.fail("not an address; a table is written t:Name, a dataset d:name",
                      options=[TABLE_PREFIX + rest, DATASET_PREFIX + rest])
        self.fail("not an address: a cell or range is Sheet!A1, a table t:Name, "
                  "a dataset d:name, a ref $name", at)

    def where(self, at: int, end: int) -> tuple[Area, str]:
        """The cell, range, columns or rows in ``body[at:end]``."""
        spec = self.body[at:end]
        if not spec:
            self.fail("a sheet is followed by a cell or range, as in Stock!A1", at)
        match = _CELLS.fullmatch(spec)
        if match:
            first = self.cell(match.group(1), match.group(2), at)
            last = self.cell(match.group(3), match.group(4), at) if match.group(3) else first
            return Area.parse(f"{first.ref}:{last.ref}"), "cells"
        match = _COLUMNS.fullmatch(spec)
        if match:
            a, b = (self.column(match.group(k), at) for k in (1, 2))
            return Area(Cell(1, min(a, b)), Cell(MAX_ROW, max(a, b))), "columns"
        match = _ROWS.fullmatch(spec)
        if match:
            a, b = (self.row(match.group(k), at) for k in (1, 2))
            return Area(Cell(min(a, b), 1), Cell(max(a, b), MAX_COLUMN)), "rows"
        self.fail("not a cell or range: A1, A1:H40, A:C or 2:5", at)

    def column(self, letters: str, at: int) -> int:
        index = column_index(letters)
        if index > MAX_COLUMN:
            self.fail(f"column {letters.upper()} is past the last column, "
                      f"{column_letters(MAX_COLUMN)}", at)
        return index

    def row(self, digits: str, at: int) -> int:
        number = int(digits)
        if not 1 <= number <= MAX_ROW:
            self.fail(f"row {number} is outside the sheet's rows, 1 to {MAX_ROW}", at)
        return number

    def cell(self, letters: str, digits: str, at: int) -> Cell:
        return Cell(self.row(digits, at), self.column(letters, at))

    def sheet_range(self, at: int) -> RangeAddress:
        self.forms = [FORMS[form] for form in _RANGE_FORMS]
        sheet, bang = self.sheet(at)
        if bang + 1 == len(self.body):
            return RangeAddress(sheet, WHOLE_GRID, "sheet")
        area, span = self.where(bang + 1, len(self.body))
        return RangeAddress(sheet, area, span)

    # -- tables and datasets -----------------------------------------------------------

    def table(self) -> TableAddress:
        self.forms = [FORMS[form] for form in _TABLE_FORMS]
        body, start = self.body, len(TABLE_PREFIX)
        end = _name_end(body, start)
        if body.startswith(("'", "#"), start) or "!" in body[start:end]:
            sheet, bang = self.sheet(start)
            end = _name_end(body, bang + 1)
            token = body[bang + 1:end]
            if token and ":" not in token and "$" not in token \
                    and not _CELLS.fullmatch(token):
                name = self.name(bang + 1, end, "a sheet-scoped name", "t:Stock!Prices")
                return TableAddress(name=name, scope=sheet, **self.selection(end))
            if end == bang + 1:
                self.fail("a detected region is named by one cell, its header's first: "
                          "t:Stock!A1; a whole sheet has no t: form", end)
            area, span = self.where(bang + 1, end)
            if span != "cells" or area.first != area.last:
                options = None if span != "cells" else \
                    [TABLE_PREFIX + RangeAddress(sheet, Area(area.first, area.first)).text]
                self.fail("a detected region is named by one cell, its header's first: "
                          "t:Stock!A1", bang + 1, options)
            region = RangeAddress(sheet, area)
            return TableAddress(region=region, **self.selection(end))
        name = self.name(start, end, "a table name", "t:Products")
        return TableAddress(name=name, **self.selection(end))

    def dataset(self) -> DatasetAddress:
        self.forms = list(_DATASET_FORMS)
        start = len(DATASET_PREFIX)
        end = _name_end(self.body, start)
        name = self.name(start, end, "a dataset id", "d:by_category")
        if not DATASET_ID.fullmatch(name):
            self.fail(_DATASET_ID_RULE, start)
        return DatasetAddress(name, **self.selection(end))

    def name(self, start: int, end: int, what: str, example: str) -> str:
        name = self.body[start:end]
        if not name:
            self.fail(f"{what} follows the prefix, as in {example}", start)
        if name.startswith("#"):
            self.fail(f"{what} does not start with #", start)
        for k, ch in enumerate(name):
            if ch.isspace() or ord(ch) < 32 or ch in _NOT_IN_NAME:
                self.fail(f"{what} has no spaces and none of [ ] / ! '", start + k)
        return name

    def selection(self, at: int) -> dict:
        """``[column]``, ``/row`` or ``/row/column`` from ``at`` to the end, a row being a
        key, ``#n``, a row range ``#n:#m`` or ``#n:``, ``*`` or a key set ``{a,b}`` and a
        column a name or ``*``; as the address's fields."""
        body = self.body
        out: dict = {}
        if at == len(body):
            return out
        if body[at] == "[":
            column, end = self.bracketed(at)
            if end != len(body):
                self.fail("a column in [ ] ends the address; a value is written "
                          "t:Products/key/Column", end)
            return {"column": column}
        if body[at] != "/":
            self.fail("a name is followed by [column] or /row", at)
        if body.startswith("{", at + 1):
            keys, end = self.key_set(at + 1)
            if len(keys) == 1:
                out["key"] = keys[0]
            else:
                out["keys"] = tuple(keys)
        else:
            row, end, quoted = self.segment(at + 1)
            if not quoted and row == ANY:
                out["any_row"] = True
            elif not quoted and row.startswith("#"):
                match = _POSITION.fullmatch(row) or _ROW_RANGE.fullmatch(row)
                if match is None:
                    self.fail("a row position is # and a whole number from 1, as in /#5, "
                              "and a row range two of them, /#9:#16 (or /#9: to the last); "
                              "a key that starts with # is quoted", at + 1)
                numbers = [int(n) for n in match.groups() if n is not None]
                if max(numbers) > MAX_ROW:
                    self.fail(f"a row position is at most {MAX_ROW}", at + 1)
                if match.re is _POSITION or numbers[0] == numbers[-1] and len(numbers) == 2:
                    out["position"] = numbers[0]
                elif len(numbers) == 1:
                    out["row_range"] = (numbers[0], None)
                else:
                    out["row_range"] = (min(numbers), max(numbers))
            else:
                out["key"] = row
        if end == len(body):
            return out
        column, after, quoted = self.segment(end + 1)
        if after != len(body):
            self.fail("an address has at most a row and a column after the name", after)
        if not quoted and column == ANY:
            out["any_column"] = True
        else:
            out["column"] = column
        return out

    def key_set(self, at: int) -> tuple[list[str], int]:
        """A key set from the ``{`` at ``at``: keys, bare or quoted, between commas; the
        keys (each once, in order) and the index after the ``}``."""
        body = self.body
        keys: dict[str, None] = {}
        i = at + 1
        while True:
            if body.startswith("'", i):
                parts, j = [], i + 1
                while True:
                    k = body.find("'", j)
                    if k < 0:
                        self.fail("a quoted key has no closing quote", i)
                    parts.append(body[j:k])
                    if body.startswith("''", k):
                        parts.append("'")
                        j = k + 2
                        continue
                    j = k + 1
                    break
                if j == i + 2:
                    self.fail("a key is never empty", i)
                value = "".join(parts)
            else:
                j = i
                while j < len(body) and body[j] not in ",}":
                    if body[j] in "{/'":
                        self.fail("in a key set, a key holding / ' { , or } is quoted", j)
                    j += 1
                value = body[i:j]
                if not value:
                    self.fail("a key set lists keys between { and }, as {A1,B2}", i)
                if value != value.strip():
                    self.fail("a key with space at either end is quoted", i)
                if value == ANY:
                    self.fail("* is every row, not a key: write t:Products/*/Price, or "
                              "quote a key that is *", i)
                if value.startswith("#"):
                    self.fail("a key set lists keys, not positions; a key that starts "
                              "with # is quoted", i)
            keys.setdefault(value, None)
            if j >= len(body):
                self.fail("a key set has no closing }", at)
            if body[j] == ",":
                i = j + 1
                continue
            if body[j] != "}":
                self.fail("in a key set, a key is followed by , or }", j)
            end = j + 1
            if end != len(body) and body[end] != "/":
                self.fail("a key set is followed by / and a column, or the end", end)
            return list(keys), end

    def segment(self, at: int) -> tuple[str, int, bool]:
        """A key or column after a ``/``: bare up to the next ``/``, or quoted (``'A/7'``,
        ``''`` for a quote); its text, the index after it, and whether it was quoted."""
        body = self.body
        if body.startswith("'", at):
            parts, i = [], at + 1
            while True:
                j = body.find("'", i)
                if j < 0:
                    self.fail("a quoted key or column has no closing quote", at)
                parts.append(body[i:j])
                if body.startswith("''", j):
                    parts.append("'")
                    i = j + 2
                    continue
                i = j + 1
                break
            if i == at + 2:
                self.fail("a key or column is never empty", at)
            if i != len(body) and body[i] != "/":
                self.fail("a quoted key or column is followed by / or the end", i)
            return "".join(parts), i, True
        end = body.find("/", at)
        end = len(body) if end < 0 else end
        value = body[at:end]
        if not value:
            self.fail("a key or column is never empty", at)
        quote = value.find("'")
        if quote >= 0:
            self.fail("a key or column holding ' is quoted, with '' for the quote", at + quote)
        if value != value.strip():
            self.fail("a key or column with space at either end is quoted", at)
        return value, end, False

    def bracketed(self, at: int) -> tuple[str, int]:
        """A column in ``[ ]``, Excel's escapes inside (``'`` before ``[ ] # '``)."""
        body, i = self.body, at + 1
        out: list[str] = []

        while i < len(body):
            ch = body[i]
            if ch == "]":
                if not out:
                    self.fail("a column in [ ] is never empty", at)
                return "".join(out), i + 1
            if ch == "'":
                if i + 1 < len(body) and body[i + 1] in "[]#'":
                    out.append(body[i + 1])
                    i += 2
                    continue
                self.fail("in [ ], ' escapes one of [ ] # '", i)
            if ch == "[":
                self.fail("in [ ], a [ is written '[", i)
            out.append(ch)
            i += 1
        self.fail("a column in [ ] has no closing ]", at)

    # -- refs --------------------------------------------------------------------------

    def ref(self) -> RefAddress:
        name, slash, rest = self.body[1:].partition("/")
        if not REF.match(name):
            self.fail(f"a ref is $ and a name matching {REF.pattern}", 0, [FORMS["ref"]])
        if slash and not rest:
            self.fail("a / after a ref is followed by more of the address", len(self.body) - 1)
        return RefAddress(name, rest)


def _name_end(body: str, at: int) -> int:
    """Where a name (or a region's cell) ends: at its ``[`` or ``/``, or the end."""
    ends = [i for i in (body.find("[", at), body.find("/", at)) if i >= 0]
    return min(ends) if ends else len(body)


def _check_name(name: object, what: str) -> None:
    if not isinstance(name, str) or not name or name.startswith("#") or any(
            ch.isspace() or ord(ch) < 32 or ch in _NOT_IN_NAME for ch in name):
        raise ValueError(f"{what} is non-empty, does not start with #, and has no spaces "
                         "and none of [ ] / ! '")


def _segment(value: str, *, row: bool) -> str:
    """A key (``row``) or a column after a ``/``, quoted when it must be."""
    if ("/" in value or "'" in value or value != value.strip() or value == ANY
            or (row and value.startswith(("#", "{")))):
        return "'" + value.replace("'", "''") + "'"
    return value


def _set_item(value: str) -> str:
    """A key in a key set, quoted when it must be."""
    if any(ch in value for ch in "/',{}") or value != value.strip() or value == ANY \
            or value.startswith("#"):
        return "'" + value.replace("'", "''") + "'"
    return value


def _escape_column(name: str) -> str:
    """A column name inside ``[ ]``: ``'`` before each ``[ ] # '``, as Excel's structured
    references escape them."""
    return re.sub(r"([\[\]#'])", r"'\1", name)


def _expect(text: str, cls: type, forms: tuple[str, ...]):
    address = parse_address(text)
    if not isinstance(address, cls):
        raise AddressError(f"a {address.kind} address where {forms[0]} was expected",
                           text=text, valid_options=[FORMS[form] for form in forms])
    return address
