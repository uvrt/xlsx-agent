"""Addresses: the forms an agent names cells, ranges, tables and datasets with.

Every result names what it returns by address, and every tool takes the same forms
(ROADMAP.md, "Addresses")::

    Stock!F2                     a cell
    Stock!A1:H40                 a range;  Stock!A:A and Stock!2:5 are whole columns and rows
    'Q3 sales'!B2:B9             a quoted sheet: ' around the name, '' for a quote inside
    #3!F2                        the sheet whose sheetId is 3 (survives a rename)
    t:Products                   a table, a defined name over a range, or a declared region
    t:Stock!A1                   a detected region, named by its header's first cell
    t:Products[Price]            one column of a table, header row excluded
    t:Products/SKU-0042/Price    the row whose key is SKU-0042, column Price
    t:Products/'A/7'/Price       a key holding / or ' is quoted, as a sheet name is
    t:Products/#5                the fifth data row (a position, not an id)
    d:by_category                a derived dataset; d:by_category[average],
                                 d:by_category/Hardware/average as for tables
    $step1                       a ref: a name a call gave to what it made

:func:`parse_address` reads any of them into a :class:`RangeAddress`,
:class:`TableAddress`, :class:`DatasetAddress` or :class:`RefAddress`, each of which has a
canonical :attr:`text` it formats back to; a canonical address parses and formats back to
itself.  Cells and areas are ooxml-edit's :class:`~ooxml_edit.charts.workbook.Cell` and
:class:`~ooxml_edit.charts.workbook.Area`, on Excel's grid (``A1`` to ``XFD1048576``).

What parsing accepts beyond the canonical form, and normalises:

* ``$`` in a cell reference (``Stock!$A$1`` is ``Stock!A1``), lowercase columns, a range
  given corner to corner the other way round (``B2:A1`` is ``A1:B2``), and a one-cell range
  (``A1:A1`` is ``A1``);
* quotes around a sheet name that needs none (``'Stock'!A1`` is ``Stock!A1``), and an
  unquoted name that :func:`needs_quotes` would quote but that is unambiguous here (no
  space, quote, ``!`` or any of ``[ ] : * ? / \\``; not starting with ``#``):
  ``Überblick!A1`` is ``'Überblick'!A1``;
* white space around the whole address.

Quoting is conservative: a sheet name is left bare only when it is ASCII letters, digits,
``_`` and ``.``, starts with a letter or ``_``, and cannot be read as a cell reference
(``A1``, ``R1C1``) or a boolean.  Excel accepts a quoted name wherever it accepts a bare
one, so the canonical form is always one Excel reads; which further names Excel itself
leaves bare is for the formula lexer (E2) to measure.

**Sheet names are not policed here.**  ECMA-376's ``sheet@name`` is an ``ST_Xstring``
whose only constraint is that it be unique in the workbook (ISO/IEC 29500-4, ``sml.xsd``,
``CT_Sheet``), so an address must be able to name whatever a file holds: any non-empty
name, quoted.  The characters Excel refuses in a sheet name and its length limit are
Excel's own rules, not the schema's; they matter when a sheet is added or renamed (E5),
and are measured from Excel then.

**Positions and ids.**  ``#3!`` names a sheet by id; ``/#5`` names a row by position.  A
key that begins with ``#`` is quoted, so the two never meet.

**No cell content in diagnostics** (ROADMAP.md, "Design rules", 7).  A key is cell
content, so an :class:`AddressError`'s message names the problem and where it is in the
address (a 1-based character position), never the address's text; sheet and table names
may appear.  The text is kept on the exception, for a program.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass
from typing import Iterable, Mapping, Union

from ooxml_edit.charts.workbook import (MAX_COLUMN, MAX_ROW, Area, Cell, column_index,
                                        column_letters)
from ooxml_edit.tools import REF

#: One example of every form, by kind: what an :class:`AddressError` offers when the
#: address is none of them.
FORMS: dict[str, str] = {
    "cell": "Stock!F2",
    "range": "Stock!A1:H40",
    "columns": "Stock!A:C",
    "rows": "Stock!2:5",
    "quoted sheet": "'Q3 sales'!B2:B9",
    "sheet by id": "#3!F2",
    "table": "t:Products",
    "detected region": "t:Stock!A1",
    "column": "t:Products[Price]",
    "value by key": "t:Products/SKU-0042/Price",
    "row by position": "t:Products/#5",
    "dataset": "d:by_category",
    "ref": "$step1",
}

#: ``sheetId`` is an ``xsd:unsignedInt`` (ISO/IEC 29500-4, ``sml.xsd``, ``CT_Sheet``).
MAX_SHEET_ID = 2**32 - 1

#: The address prefixes of a table (or name, or region) and of a dataset.
TABLE_PREFIX = "t:"
DATASET_PREFIX = "d:"

#: How a :class:`RangeAddress` spans the grid: a rectangle of cells, whole columns
#: (``A:C``), or whole rows (``2:5``).
SPANS = ("cells", "columns", "rows")

_PLAIN_SHEET = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_CELL_LIKE = re.compile(r"[A-Za-z]{1,3}[0-9]+")
_R1C1_LIKE = re.compile(r"(?:[Rr][0-9]*(?:[Cc][0-9]*)?|[Cc][0-9]*)")
_BOOLEANS = ("TRUE", "FALSE")

# Characters a bare sheet name may not hold: they delimit, or Excel never allows them.
_NOT_IN_BARE_SHEET = frozenset("'![]:*?/\\")
# Characters a table, name, region or dataset id may not hold: the address uses them.
_NOT_IN_NAME = frozenset("[]/!'")

_CELLS = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?::\$?([A-Za-z]{1,3})\$?([0-9]{1,7}))?")
_COLUMNS = re.compile(r"\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})")
_ROWS = re.compile(r"\$?([0-9]{1,7}):\$?([0-9]{1,7})")
_POSITION = re.compile(r"#([1-9][0-9]{0,6})")

_RANGE_FORMS = ("cell", "range", "columns", "rows", "quoted sheet", "sheet by id")
_TABLE_FORMS = ("table", "detected region", "column", "value by key", "row by position")
_DATASET_FORMS = ("d:by_category", "d:by_category[average]", "d:by_category/Hardware/average",
                  "d:by_category/#2")


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
        return f"#{self.id}" if self.id is not None else quote_sheet(self.name)

    def __str__(self) -> str:
        return self.text


# ------------------------------------------------------------------------------------------
# The addresses
# ------------------------------------------------------------------------------------------


@dataclass(frozen=True)
class RangeAddress:
    """A cell or a range on one sheet: ``Stock!F2``, ``Stock!A1:H40``, ``#3!A:A``.

    ``span`` is ``"cells"``, ``"columns"`` (``A:C``: ``area`` runs from row 1 to
    :data:`~ooxml_edit.charts.workbook.MAX_ROW`) or ``"rows"`` (``2:5``: column 1 to
    :data:`~ooxml_edit.charts.workbook.MAX_COLUMN`).  Whole columns and rows are clipped to
    the sheet's used range when read (:meth:`clipped`)."""

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

    @property
    def kind(self) -> str:
        """``"cell"`` or ``"range"``."""
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
        first, last = self.area.first, self.area.last
        if self.span == "columns":
            where = f"{column_letters(first.column)}:{column_letters(last.column)}"
        elif self.span == "rows":
            where = f"{first.row}:{last.row}"
        else:
            where = self.area.ref
        return f"{self.sheet.text}!{where}"

    def __str__(self) -> str:
        return self.text

    def on(self, sheet: SheetRef) -> "RangeAddress":
        """The same cells, with the sheet named as ``sheet`` names it (by id or by name)."""
        return RangeAddress(sheet, self.area, self.span)

    def clipped(self, used: Area | None) -> Area | None:
        """The cells meant, given the sheet's used range: whole columns and rows clipped to
        it (``None`` when they miss it, or the sheet is empty); a rectangle of cells as it
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
        """``text`` as a cell or range; :class:`AddressError` for any other form."""
        return _expect(text, cls, _RANGE_FORMS)


class _Selection:
    """What table and dataset addresses share: a column, a row (by key or by position),
    or a value (a row and a column)."""

    column: str | None
    key: str | None
    position: int | None

    def _check_selection(self) -> None:
        if self.key is not None and self.position is not None:
            raise ValueError("a row is named by key or by position, not both")
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
        has_row = self.key is not None or self.position is not None
        if has_row:
            return "value" if self.column is not None else "row"
        return "column" if self.column is not None else None

    def _selection_text(self) -> str:
        has_row = self.key is not None or self.position is not None
        if not has_row:
            return "" if self.column is None else f"[{_escape_column(self.column)}]"
        row = f"#{self.position}" if self.position is not None else _segment(self.key, row=True)
        column = "" if self.column is None else "/" + _segment(self.column, row=False)
        return f"/{row}{column}"


@dataclass(frozen=True)
class TableAddress(_Selection):
    """A table, a defined name over a range or a declared region (``t:Products``), or a
    detected region (``t:Stock!A1``: ``region`` is its header's first cell, ``name`` is
    ``None``); with, optionally, a column (``t:Products[Price]``), a row by key
    (``t:Products/SKU-0042``) or position (``t:Products/#5``), or a value
    (``t:Products/SKU-0042/Price``).

    Parsing does not say whether the name is a table, a name or a region, nor whether the
    key and column exist: the workbook does, once tables are read (R4)."""

    name: str | None = None
    region: RangeAddress | None = None
    column: str | None = None
    key: str | None = None
    position: int | None = None

    def __post_init__(self) -> None:
        if (self.name is None) == (self.region is None):
            raise ValueError("a table address has a name or a region, one of the two")
        if self.name is not None:
            _check_name(self.name, "a table name")
        if self.region is not None and self.region.kind != "cell":
            raise ValueError("a detected region is named by one cell, its header's first")
        self._check_selection()

    @property
    def kind(self) -> str:
        """``"table"`` or ``"region"`` (the whole of it), ``"column"``, ``"row"`` or
        ``"value"``."""
        return self._selected or ("region" if self.region is not None else "table")

    @property
    def is_region(self) -> bool:
        """Named by a header cell (``t:Stock!A1``), not by a name."""
        return self.region is not None

    @property
    def text(self) -> str:
        base = self.name if self.name is not None else self.region.text
        return TABLE_PREFIX + base + self._selection_text()

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
    ``d:by_category/Hardware/average``, ``d:by_category/#2``."""

    id: str
    column: str | None = None
    key: str | None = None
    position: int | None = None

    def __post_init__(self) -> None:
        _check_name(self.id, "a dataset id")
        self._check_selection()

    @property
    def kind(self) -> str:
        """``"dataset"`` (the whole of it), ``"column"``, ``"row"`` or ``"value"``."""
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
    """An address's canonical text (its :attr:`text`)."""
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

    def fail(self, message: str, at: int | None = None, options: Iterable[str] | None = None):
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

    def fail_no_sheet(self, at: int) -> None:
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
            area, span = self.where(bang + 1, end)
            if span != "cells" or area.first != area.last:
                options = None if span != "cells" else \
                    [TABLE_PREFIX + RangeAddress(sheet, Area(area.first, area.first)).text]
                self.fail("a detected region is named by one cell, its header's first: "
                          "t:Stock!A1", bang + 1, options)
            region = RangeAddress(sheet, area)
            column, key, position = self.selection(end)
            return TableAddress(region=region, column=column, key=key, position=position)
        name = self.name(start, end, "a table name", "t:Products")
        column, key, position = self.selection(end)
        return TableAddress(name=name, column=column, key=key, position=position)

    def dataset(self) -> DatasetAddress:
        self.forms = list(_DATASET_FORMS)
        start = len(DATASET_PREFIX)
        end = _name_end(self.body, start)
        name = self.name(start, end, "a dataset id", "d:by_category")
        column, key, position = self.selection(end)
        return DatasetAddress(name, column=column, key=key, position=position)

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

    def selection(self, at: int) -> tuple[str | None, str | None, int | None]:
        """``[column]``, ``/key``, ``/#n``, ``/key/column`` or ``/#n/column`` from ``at`` to
        the end, as ``(column, key, position)``."""
        body = self.body
        if at == len(body):
            return None, None, None
        if body[at] == "[":
            column, end = self.bracketed(at)
            if end != len(body):
                self.fail("a column in [ ] ends the address; a value is written "
                          "t:Products/key/Column", end)
            return column, None, None
        if body[at] != "/":
            self.fail("a name is followed by [column] or /row", at)
        row, end, quoted = self.segment(at + 1)
        key = position = None
        if not quoted and row.startswith("#"):
            match = _POSITION.fullmatch(row)
            if match is None:
                self.fail("a row position is # and a whole number from 1, as in /#5; "
                          "a key that starts with # is quoted", at + 1)
            position = int(match.group(1))
            if position > MAX_ROW:
                self.fail(f"a row position is at most {MAX_ROW}", at + 1)
        else:
            key = row
        if end == len(body):
            return None, key, position
        column, end, _ = self.segment(end + 1)
        if end != len(body):
            self.fail("an address has at most a row and a column after the name", end)
        return column, key, position

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
        body, out, i = self.body, [], at + 1
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
    if ("/" in value or "'" in value or value != value.strip()
            or (row and value.startswith("#"))):
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
