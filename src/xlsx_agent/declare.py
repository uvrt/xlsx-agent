"""Tables declared from the caller's configuration, before ``describe``.

A workbook made from a template is known to its application: which sheets hold which
table, on which row its header is, which column is its key, what each column means, and
which columns a model should not see.  The application says so once, as plain data (its
own configuration, mapped to this), and the tables are then addressed, described, read,
queried and derived from like any other (``t:checks``)::

    wb.declare_tables({"tables": [
        {"name": "checks", "sheets": ["Area *"], "header_row": 3, "key": "Ref",
         "description": "One row per check, every area's sheet",
         "fields": [
             {"header": "Rating (1 to 5, in halves)", "id": "rating",
              "description": "The reviewer's rating"},
             {"header": "Reviewer notes", "hidden": True}]},
        {"name": "run", "type": "key_value", "sheets": "Run info"}]})
    wb.read("t:checks/C-07/rating")      # the value, with its own cell: 'Area 2'!D5
    wb.read("t:run/report_date/value")

Sheets the model must never see are declared alongside, ``{"sheets": "Notes", "hidden":
true}`` (see below), and so are terms it must never read, ``{"redact": ["Acme Corp"]}``;
a section's header can be found by its text, ``"header_row": {"match": ["Ref", "Item"]}``.

Declarations are the application's configuration, never a model's: no tool takes them.
The tool layer takes them when it opens a workbook (``xlsx_agent.tools.format(tables=...)``
for every workbook the format opens, or ``xlsx_agent.tools.declare(session, ...)`` for one
session's), and :meth:`Workbook.declare_tables
<xlsx_agent.workbook.Workbook.declare_tables>` in Python.

**A declaration** is a mapping (a list of them, or ``{"tables": [...]}``):

* ``name`` (required): the table's name, ``t:<name>``; a name as ECMA-376 defines one
  (:func:`~xlsx_agent.address.is_name`), no Excel table's or workbook-wide defined
  name's.  Declaring a name again replaces the table (or a region declared with it).
* ``sheets`` (required): a sheet name, or a list of them; each an exact name (matched
  case aside, as Excel matches them) or a glob (``*``, ``?``, ``[...]``, which no sheet
  name can hold: ``Area *``).  The sheets are taken in the workbook's order (the tabs'),
  each once.  An exact name that is no sheet, or a sheet that holds no cells (a chart
  sheet), is refused; a glob skips sheets that hold no cells, and matching none is
  refused.
* ``type``: ``"rows"`` (the default: a row per item) or ``"key_value"`` (see below).
* ``header_row``: the header's row on every sheet (1 by default; 0 for none, the default
  of a key-value table).  The header is one row; its text names the columns.  Every sheet's
  header must be the same text, column by column, as the first sheet's: one that differs
  is refused, naming both cells.  Or the header found by its text (see below):
  ``{"match": ["Ref", "Item"], "rows": "8:40"}``.
* ``first_row``: the first data row (the row under the header; without a header, the first
  row of the sheet's used range).
* ``last_row``: the last data row on every sheet.  Without it, the data rows run down to the
  row before the first row blank across the columns (or the used range's end), as a
  detected region's do; when a sheet holds more below that blank row, within the columns,
  the table says so (``rows_below``, at the first such cell), and ``last_row`` takes it in.
* ``columns``: the sheet columns, ``"A:F"`` (by default, the first sheet's header row from
  its first cell holding something to its last; without a header, the used range's
  columns).
* ``key``: the key column (by header or id): ``t:checks/C-07/rating``.  Every data row,
  on every sheet, must have a key (text, or a whole number) that no other row has: a key
  two rows share, on one sheet or two, is refused naming both cells (never the key); a row
  without one, naming its cell.
* ``description``: what the table holds, shown by ``describe``.
* ``fields``: what is said of some columns, each a mapping with ``header`` (its header
  text, matched case aside) or ``column`` (its letter, ``"D"``), and any of ``id`` (the
  name the column is addressed and shown by: ``t:checks[rating]``; its header still
  finds it), ``description`` (shown by ``describe`` and ``xlsx_profile``), ``hidden``
  (``true``: the column is hidden from the model) and ``clean`` (rules for its text: a
  leading label stripped, white space collapsed, a split into a list --
  ``{"strip": "^Lead:[^;]*;\\s*", "collapse_space": true, "split": ";"}``; see
  :mod:`xlsx_agent.clean`; never on the key column, ``clean_key``, nor on a hidden one).
  Columns not listed keep their header's name.
* ``optional``: ``true`` when a workbook without the sheets is no error: the table is then
  not declared.

**A header found by its text** (``"header_row": {"match": [...], "rows": "8:40"}``), for a
sheet of several sections: on each sheet, the one row whose text cells, in the table's
``columns`` (else the used range's), hold every text of ``match`` -- each a cell's whole
text, case and runs of white space aside, in any order, other cells besides -- within
``rows`` (first to last), or without it below the last row of the tables declared before it
in the same declarations on that sheet.  The data rows run from under it to the first blank
row or to the row before the next row that matches, whichever comes first (so ``first_row``
and ``last_row`` are not taken with it), and nothing past that next match is flagged
``rows_below``.  No row matching is refused (``no_header``); more than one in the rows
looked in too (``ambiguous_header``), the rows by number only (``valid_options``), with
what to do: give ``rows``, or more of the header's texts.  Decided: one row, not the
first, so a section never silently takes another's header; and whole texts, not parts of
them, so ``Ref`` does not find ``Reference``.

**A table over several sheets** reads as one: its data rows are each sheet's, sheet by
sheet in the workbook's order, its header the first sheet's.  It is one rectangle of its
own (:class:`~xlsx_agent.read.regions.Span`), and every value read from it -- by
``xlsx_read``, ``xlsx_query``, ``describe``'s samples, a derive's sources and left-out
cells, a dataset's sources, a comparison, a check -- carries its own cell's address
(``'Area 2'!D5``); a row by position counts across the sheets (``t:checks/#12``), and a
row by key finds its sheet.  Its range is each sheet's range, joined by ``,``.  **The
sheet is not a column**: a value is a cell, and a sheet's name is no cell, so there is no
``_sheet`` column to group or filter by.  Every address names its sheet; for one sheet's
figures, declare that sheet's table as well (``{"name": "checks_north", "sheets": "Area
North", ...}``).

**Hidden from the model**: a hidden column is left out of the table -- ``describe``,
``xlsx_read``, ``xlsx_query``, ``xlsx_derive``, ``xlsx_profile``, ``xlsx_compare``,
``xlsx_validate_data`` and a deck's or document's ``from`` never show it, and an address
that names it is refused as naming no column (its name is never offered).  Its cells are
withheld from a read of the sheet's cells (``xlsx_read`` of ``Area 1!A1:H40`` gives them as
``withheld``; :meth:`Workbook.read <xlsx_agent.workbook.Workbook.read>` gives them blank),
from ``find_text``, from a validation list drawn from them (its ``reason`` is
``withheld``), from the problems, and from a region detected or declared over them.  A
defined name over any of them is refused (``declared``, with the table as
``valid_options``) and neither listed by ``describe`` nor found by ``find_text``: a name
may say what its cells hold.  Another declared table that shows one of them is refused
(``hidden_elsewhere``): it hides the column too.  A column Excel hides (a hidden column of
the sheet) stays a column like any other unless it is declared hidden; its values say
``hidden``, as always.

*Below the table*: a hidden column is hidden from its header row (the first data row,
without one) to the sheet's last row, on each of the table's sheets -- or, when another
table is declared below it on that sheet (in the same declarations, or kept already), to
just above that table's header row (its first row, without one): the next section
(:attr:`Part.end <xlsx_agent.read.regions.Part.end>`).  The cells under the table's rows
are the column's too -- a row added after ``last_row``, a total, a note left under it,
everything between it and the next section -- and nothing else in a declaration says
where the column's content stops; only a declared table does (an Excel table or a
detected block below does not).  The cells above its header are not the table's (a
title, a block of run details beside it) and stay visible.  So another table, region or
name over that column within those rows is refused as one over the table's rows is,
while a section declared below may show it.  A table kept already is not bounded again
by one declared after it: declare the sections together.

*Formulas*: a visible cell's formula that reaches a withheld cell is given without its
text: its cached value as for any formula, ``formula_withheld`` set and no ``f``
(``xlsx_read`` with ``formulas``; :attr:`Value.formula_text
<xlsx_agent.read.types.Value.formula_text>` is ``None``).  It reaches one through a
reference over it (``=D2``, ``=SUM(Log!D:D)``, ``=SUM(2:2)``, on its sheet or another), a
structured reference to the hidden column or to the whole table (``ChecksTable[Note]``,
``[@Note]``, ``ChecksTable[[Item]:[Note]]``, ``ROWS(ChecksTable)``), a defined name over
one or whose formula reaches one, or ``INDIRECT`` or ``OFFSET``, whose reach its text does
not say (:meth:`Regions.formula_withheld
<xlsx_agent.read.regions.Regions.formula_withheld>`).  The text is withheld because it
says what the hidden column is: its header's name in a structured reference, a name over
it, a literal compared with its cells.  A shared formula's group is judged whole, its
references moved over the range it fills.  **The value is kept**: a visible cell is data
the workbook's owner shows, whatever it is computed from -- a count, a sum, a flag
computed from the notes -- and the reader cannot tell a copy (``=D2``) from a figure
computed from many cells.  A column whose formulas copy hidden content is to be declared
hidden as well.

**Over an Excel table**: a declared table that covers any cell of an Excel table (a
ListObject) replaces it.  ``describe`` lists the declared table alone, and ``t:<the Excel
table>`` -- with a column, a row, a key set or a wildcard, in every tool -- is refused
(``declared``, ``t:<the Excel table> lies in t:checks``, with ``t:checks`` as
``valid_options``), so a column the declared table hides is never reached by the Excel
table's name.  The Excel table's name is not listed as an alias: it would name the same
cells two ways, with different columns, and a table over several sheets is more than the
Excel table.  It stays taken (a declared table takes a name of its own).  The workbook's
``tables`` and ``open_document``'s count leave it out.

**Sheets hidden from the model**: ``{"sheets": "Notes", "hidden": true}`` in the same
declarations (``sheets`` a name, a glob or a list, as a table's; ``optional`` as a
table's; no other key) withholds whole sheets -- a table with every column hidden is
refused (``no_columns``), so this is the way to keep a sheet from a model.  The workbook
then has no such sheet (:meth:`Workbook.withhold_sheets
<xlsx_agent.workbook.Workbook.withhold_sheets>`): it is not among ``Workbook.sheets``, in
``describe``, ``open_document``'s count, ``find_text``, the problems or a comparison; an
address by its name or id, a ``t:`` address of an Excel table on it, of a defined name on
it, local to it or whose formula names it, and a declared table naming it are refused
exactly as for a sheet, table or name the workbook does not have -- its name is never
offered, and a glob of a table never matches it.  A drop-down drawn from it reads as one
from no sheet (``unknown_sheet``), and a link to a place on it is left out.  A formula on
another sheet that reaches it -- a reference, a run of sheets over it, its table, a name
over it, ``INDIRECT``/``OFFSET``, or its name anywhere in the text (a string an
``INDIRECT`` or ``HYPERLINK`` takes) -- keeps its value, not its text
(``formula_withheld``), as for a hidden column.  Decided: **the name is not listed** --
not even as "withheld": a sheet's name may itself say what it holds -- and a sheet named
that is not there is refused by its path only, without the name, since the error reaches
the model; a sheet's ``sheetId`` is not renumbered, so the ids shown may have a gap, as
after a sheet is deleted.  A sheet withheld stays withheld; a table or region kept already
that reads it is refused (``on_hidden_sheet``, naming the table, not the sheet).  Charts and
pivot tables are not listed by the catalogue, so none reveals it.

**Terms redacted from the model**: ``{"redact": ["Acme Corp", "re:Project\\s+\\w+"],
"with": "[redacted]"}`` in the same declarations (:mod:`xlsx_agent.redact`, where the
rules and their cost are) replaces each match in every text the workbook gives
(:meth:`Workbook.redact <xlsx_agent.workbook.Workbook.redact>`).  It is set first, before
any table of the same declarations: declared after a table, region, key or hidden sheet
is kept it is refused (``redact_late``), unless the rules are the ones set already.  The
caller's declarations still name sheets as the file stores them (a name, a glob) and find
headers, keys and a header's ``match`` by their stored text: each is redacted as the
workbook's text is before it is compared.  A table name, column id or description of the
caller's own that holds a match is refused (``redacted_term``, by path only), since the
model would read it; so is a refusal never given with a term (a sheet named that is not
there is named redacted).  Decided: names are shown redacted where an address can still
hold them -- a sheet (quoted), an Excel table's column (escaped) -- and left out where a
redacted name would be no name: an Excel table or defined name whose name holds a match
is not listed, its cells still read by address (an Excel table's as a block of its
sheet).  Several redactions apply in order.

**A key-value table** (``"type": "key_value"``) is a field column and a value column, and
any more (a confidence, a source...): ``t:run/report_date/value`` is the value of the row
whose field is ``report_date``.  Its first two columns are called ``field`` and ``value``
(an ``id`` renames them; a header's text finds them too), the rest by their header or
letter; its key is its first column.  ``describe`` shows it compactly, as its fields and
values.  A block of two columns with no header, whose first column is a different text
on every row (two rows at least) and whose first row holds something other than text, is
detected as one (``t:'Run info'!A1/report_date/value``); a wider one is declared, since a
first row of numbers over more columns is as likely a header of years.

**Validation**: a declaration that is not one of these, or that the workbook does not
fit, is refused with :class:`DeclarationError` -- its ``code``, ``path``
(``tables[0].fields[2].header``) and ``valid_options`` -- before any table is kept, so a
list of declarations is kept whole or not at all.  Never guessed past, and never with a
cell's content in the message.
"""

from __future__ import annotations

import fnmatch
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any, Mapping, Sequence

from ooxml_edit.sheets import Area, Cell, column_index, column_letters

from .address import NameNotFound, RangeAddress, SheetRef, is_name, nearest, quote_sheet
from .clean import CLEAN_KEYS, Clean, RuleError, parse_clean
from .redact import REDACT_KEYS, RedactError, Redaction, Rule, parse_rule
from .read.regions import KEY_VALUE_COLUMNS, Part, Region, Span, _header_names
from .read.types import Problem

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: What a declared table is.
TYPES = ("rows", "key_value")

#: The keys a table's declaration takes.
TABLE_KEYS = ("name", "sheets", "type", "header_row", "first_row", "last_row", "columns",
              "key", "description", "fields", "optional")

#: The keys a declaration of sheets withheld from the model takes.
SHEET_KEYS = ("sheets", "hidden", "optional")

#: The keys a header found by its text takes (``header_row: {"match": [...]}``).
MATCH_KEYS = ("match", "rows")

#: The most texts a header is found by.
MAX_MATCH = 50

#: The keys a field takes.
FIELD_KEYS = ("header", "column", "id", "description", "hidden", "clean")

#: The longest description (characters) a table or column takes: ``describe`` shows it.
MAX_DESCRIPTION = 300

#: The longest column id (characters).
MAX_ID = 64

#: A sheet's last row (ECMA-376 Part 1, §18.3.1.73: rows 1 to 1,048,576).
MAX_SHEET_ROW = 1048576

#: The characters that make a sheet name a glob: no sheet name holds them.
GLOB = "*?["


class DeclarationError(ValueError):
    """A declaration refused: ``code`` says why (``bad_declaration``, ``unknown_sheet``,
    ``no_sheets``, ``not_cells``, ``empty_sheet``, ``name_taken``, ``bad_name``,
    ``duplicate_name``, ``header_mismatch``, ``no_header``, ``bad_columns``, ``bad_rows``,
    ``no_rows``, ``unknown_column``, ``duplicate_field``, ``bad_id``, ``id_taken``,
    ``no_columns``, ``hidden_key``, ``no_key``, ``duplicate_key``, ``bad_pattern``,
    ``clean_key``, ``hidden_elsewhere``, ``ambiguous_header``, ``on_hidden_sheet``,
    ``redacted_term``, ``redact_late``), ``path`` where in the
    declarations (``tables[0].fields[2].header``), ``valid_options`` what would do instead.
    The message names sheets, cells, paths and what the declaration gave, never a cell's
    content."""

    def __init__(self, code: str, path: str, message: str, *,
                 valid_options: Sequence[str] = ()) -> None:
        super().__init__(f"declaration {path}: {message}" if path else f"declaration: {message}")
        self.code = code
        self.path = path
        self.message = str(self)
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class FieldSpec:
    """One field of a declaration, checked for its form."""

    path: str
    header: str | None = None
    column: int | None = None
    id: str | None = None
    description: str | None = None
    hidden: bool = False
    clean: "Clean | None" = None


@dataclass(frozen=True)
class HiddenSheets:
    """Sheets withheld from the model (``{"sheets": "Notes", "hidden": true}``), checked
    for their form."""

    path: str
    sheets: tuple[str, ...]
    optional: bool = False


@dataclass(frozen=True)
class RedactTerms:
    """Terms redacted from every text of the workbook (``{"redact": [...], "with":
    "[redacted]"}``, :mod:`xlsx_agent.redact`), checked and compiled."""

    path: str
    rule: Rule


@dataclass(frozen=True)
class TableSpec:
    """One table's declaration, checked for its form (not yet against a workbook).
    ``header_match`` the texts its header row is found by on each sheet (empty: the
    header is on ``header_row``), within the rows ``header_within`` (``None``: below the
    tables declared before it on that sheet)."""

    path: str
    name: str
    sheets: tuple[str, ...]
    type: str = "rows"
    header_row: int = 1
    first_row: int | None = None
    last_row: int | None = None
    columns: tuple[int, int] | None = None
    key: str | None = None
    description: str | None = None
    fields: tuple[FieldSpec, ...] = ()
    optional: bool = False
    header_match: tuple[str, ...] = ()
    header_within: tuple[int, int] | None = None


# ------------------------------------------------------------------------------------------
# The form
# ------------------------------------------------------------------------------------------


def parse_declarations(declarations: Any) -> "list[TableSpec | HiddenSheets | RedactTerms]":
    """``declarations`` (a mapping, a list of them, or ``{"tables": [...]}``) as
    :class:`TableSpec`, :class:`HiddenSheets` and :class:`RedactTerms`, checked for their form;
    :class:`DeclarationError` otherwise."""
    if isinstance(declarations, Mapping) and "tables" in declarations:
        extra = sorted(set(declarations) - {"tables"})
        if extra:
            raise DeclarationError("bad_declaration", extra[0], "unknown key; the "
                                   "declarations are {\"tables\": [...]}",
                                   valid_options=["tables"])
        items, base = declarations["tables"], "tables"
    elif isinstance(declarations, Mapping):
        items, base = [declarations], "tables"
    else:
        items, base = declarations, "tables"
    if isinstance(items, (str, bytes)) or not isinstance(items, Sequence):
        raise DeclarationError("bad_declaration", base, "a list of table declarations")
    specs = [_table(item, f"{base}[{n}]") for n, item in enumerate(items)]
    seen: dict[str, str] = {}
    for spec in specs:
        if not isinstance(spec, TableSpec):
            continue
        folded = spec.name.casefold()
        if folded in seen:
            raise DeclarationError("duplicate_name", f"{spec.path}.name", f"t:{spec.name} is "
                                   f"declared twice (first at {seen[folded]})")
        seen[folded] = spec.path
    return specs


def _table(item: Any, path: str) -> "TableSpec | HiddenSheets | RedactTerms":
    if not isinstance(item, Mapping):
        raise DeclarationError("bad_declaration", path, "a table's declaration is a mapping "
                               "with name, sheets and more", valid_options=list(TABLE_KEYS))
    if "redact" in item:
        try:
            return RedactTerms(path=path, rule=parse_rule(item))
        except RedactError as error:
            raise DeclarationError(error.code, f"{path}.{error.key}", error.message,
                                   valid_options=list(REDACT_KEYS)
                                   if error.code == "bad_declaration" else ()) from None
    if "hidden" in item:
        return _hidden_sheets(item, path)
    for key in item:
        if key not in TABLE_KEYS:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "unknown key",
                                   valid_options=nearest(str(key), TABLE_KEYS))
    for key in ("name", "sheets"):
        if key not in item:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "is required")
    name = item["name"]
    if not is_name(name):
        raise DeclarationError("bad_name", f"{path}.name", "a table's name is a name as "
                               "ECMA-376 defines one: a letter, _ or \\ first, then letters, "
                               "digits, _ and .; not TRUE, FALSE or a cell reference",
                               valid_options=["checks", "run_info"])
    sheets = item["sheets"]
    if isinstance(sheets, str):
        sheets = [sheets]
    if isinstance(sheets, (bytes, Mapping)) or not isinstance(sheets, Sequence) or not sheets \
            or not all(isinstance(s, str) and s for s in sheets):
        raise DeclarationError("bad_declaration", f"{path}.sheets", "a sheet name or glob "
                               "(Area *), or a list of them")
    kind = item.get("type", "rows")
    if kind not in TYPES:
        raise DeclarationError("bad_declaration", f"{path}.type", "unknown table type",
                               valid_options=list(TYPES))
    match: tuple[str, ...] = ()
    within = None
    if isinstance(item.get("header_row"), Mapping):
        match, within = _header_match(item["header_row"], f"{path}.header_row")
        for key in ("first_row", "last_row"):
            if key in item:
                raise DeclarationError("bad_rows", f"{path}.{key}", "a header found by its "
                                       "text has its data rows under it, to the first blank "
                                       "row or the next row it matches: no first_row or "
                                       "last_row; give header_row.rows to look in fewer rows")
        header_row = 1
    else:
        header_row = _row(item, "header_row", path, minimum=0,
                          default=1 if kind == "rows" else 0)
    first_row = _row(item, "first_row", path)
    last_row = _row(item, "last_row", path)
    if first_row is not None and header_row and first_row <= header_row:
        raise DeclarationError("bad_rows", f"{path}.first_row", "the first data row is "
                               "below the header row")
    top = first_row or header_row + 1
    if last_row is not None and last_row < top - 1:
        raise DeclarationError("bad_rows", f"{path}.last_row", "the last data row is at or "
                               "below the first (one above it: no data rows)")
    columns = None
    if "columns" in item:
        columns = _columns(item["columns"], f"{path}.columns")
    key = item.get("key")
    if key is not None and not (isinstance(key, str) and key):
        raise DeclarationError("bad_declaration", f"{path}.key", "a column's header or id")
    description = _description(item, path)
    fields = item.get("fields", ())
    if isinstance(fields, (str, bytes, Mapping)) or not isinstance(fields, Sequence):
        raise DeclarationError("bad_declaration", f"{path}.fields", "a list of fields")
    optional = item.get("optional", False)
    if not isinstance(optional, bool):
        raise DeclarationError("bad_declaration", f"{path}.optional", "true or false")
    return TableSpec(path=path, name=name, sheets=tuple(sheets), type=kind,
                     header_row=header_row, first_row=first_row, last_row=last_row,
                     columns=columns, key=key, description=description,
                     fields=tuple(_field(f, f"{path}.fields[{n}]") for n, f in enumerate(fields)),
                     optional=optional, header_match=match, header_within=within)


def _hidden_sheets(item: Mapping, path: str) -> HiddenSheets:
    """``{"sheets": ..., "hidden": true}``: sheets withheld from the model."""
    for key in item:
        if key not in SHEET_KEYS:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "a declaration of "
                                   "sheets hidden from the model takes sheets, hidden and "
                                   "optional only (hide a table's column with a field's "
                                   "hidden)", valid_options=list(SHEET_KEYS))
    if item["hidden"] is not True:
        raise DeclarationError("bad_declaration", f"{path}.hidden", "true: the sheets are "
                               "hidden from the model (leave a sheet out to show it)")
    sheets = item.get("sheets")
    if isinstance(sheets, str):
        sheets = [sheets]
    if isinstance(sheets, (bytes, Mapping)) or not isinstance(sheets, Sequence) or not sheets \
            or not all(isinstance(s, str) and s for s in sheets):
        raise DeclarationError("bad_declaration", f"{path}.sheets", "a sheet name or glob "
                               "(Notes *), or a list of them")
    optional = item.get("optional", False)
    if not isinstance(optional, bool):
        raise DeclarationError("bad_declaration", f"{path}.optional", "true or false")
    return HiddenSheets(path=path, sheets=tuple(sheets), optional=optional)


def _header_match(value: Mapping, path: str) -> tuple[tuple[str, ...], tuple[int, int] | None]:
    """``header_row: {"match": [...], "rows": "5:40"}``: the texts, and the rows."""
    for key in value:
        if key not in MATCH_KEYS:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "unknown key",
                                   valid_options=nearest(str(key), MATCH_KEYS))
    match = value.get("match")
    if isinstance(match, str):
        match = [match]
    if isinstance(match, (bytes, Mapping)) or not isinstance(match, Sequence) or not match \
            or len(match) > MAX_MATCH \
            or not all(isinstance(t, str) and _folded(t) for t in match):
        raise DeclarationError("bad_declaration", f"{path}.match", f"1 to {MAX_MATCH} of the "
                               "header's texts, each a cell's whole text (case and spaces "
                               "aside)", valid_options=["Ref", "Item"])
    within = None
    if "rows" in value:
        rows = value["rows"]
        parts = rows.split(":") if isinstance(rows, str) else []
        if len(parts) != 2 or not all(p.strip().isdigit() for p in parts) \
                or not 1 <= int(parts[0]) <= int(parts[1]) <= 1048576:
            raise DeclarationError("bad_rows", f"{path}.rows", "the rows to look in, first "
                                   "to last, as 5:40", valid_options=["5:40"])
        within = (int(parts[0]), int(parts[1]))
    return tuple(match), within


def _folded(text: str) -> str:
    """A header's text as it is matched: white space runs as one space, none at the ends,
    case aside."""
    return " ".join(text.split()).casefold()


def _field(item: Any, path: str) -> FieldSpec:
    if not isinstance(item, Mapping):
        raise DeclarationError("bad_declaration", path, "a field is a mapping with header "
                               "or column, and id, description or hidden",
                               valid_options=list(FIELD_KEYS))
    for key in item:
        if key not in FIELD_KEYS:
            raise DeclarationError("bad_declaration", f"{path}.{key}", "unknown key",
                                   valid_options=nearest(str(key), FIELD_KEYS))
    if ("header" in item) == ("column" in item):
        raise DeclarationError("bad_declaration", path, "a field names its column by header "
                               "or by column (its letter), one of the two",
                               valid_options=["header", "column"])
    header = item.get("header")
    if header is not None and not (isinstance(header, str) and header):
        raise DeclarationError("bad_declaration", f"{path}.header", "the header's text")
    column = None
    if "column" in item:
        column = _letter(item["column"], f"{path}.column")
    ident = item.get("id")
    if ident is not None and not (isinstance(ident, str) and ident and len(ident) <= MAX_ID
                                  and ident == ident.strip() and "\n" not in ident):
        raise DeclarationError("bad_id", f"{path}.id", f"a column id is a name of 1 to {MAX_ID} "
                               "characters, no space at its ends", valid_options=["rating"])
    hidden = item.get("hidden", False)
    if not isinstance(hidden, bool):
        raise DeclarationError("bad_declaration", f"{path}.hidden", "true or false")
    clean = None
    if "clean" in item:
        try:
            clean = parse_clean(item["clean"])
        except RuleError as error:
            where = f"{path}.clean" + (f".{error.key}" if error.key else "")
            raise DeclarationError(error.code, where, error.message,
                                   valid_options=list(CLEAN_KEYS)
                                   if error.code == "bad_declaration" else ()) from None
        if hidden:
            raise DeclarationError("bad_declaration", f"{path}.clean", "a column hidden from "
                                   "the model is never shown: its text is not cleaned")
    return FieldSpec(path=path, header=header, column=column, id=ident,
                     description=_description(item, path), hidden=hidden, clean=clean)


def _description(item: Mapping, path: str) -> str | None:
    text = item.get("description")
    if text is not None and not (isinstance(text, str) and text.strip()
                                 and len(text) <= MAX_DESCRIPTION):
        raise DeclarationError("bad_declaration", f"{path}.description", "a description is "
                               f"text of 1 to {MAX_DESCRIPTION} characters")
    return text


def _row(item: Mapping, key: str, path: str, *, minimum: int = 1,
         default: int | None = None) -> int | None:
    value = item.get(key, default)
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= 1048576:
        raise DeclarationError("bad_rows", f"{path}.{key}", f"a row number from {minimum}")
    return value


def _letter(value: Any, path: str) -> int:
    try:
        if isinstance(value, str) and value.isalpha() and value.isascii():
            number = column_index(value.upper())
            if 1 <= number <= 16384:
                return number
    except (ValueError, KeyError):
        pass
    raise DeclarationError("bad_columns", path, "a column's letters, A to XFD",
                           valid_options=["D"])


def _columns(value: Any, path: str) -> tuple[int, int]:
    if not isinstance(value, str) or value.count(":") != 1:
        raise DeclarationError("bad_columns", path, "the columns as letters, A:F",
                               valid_options=["A:F"])
    first, last = (_letter(part.strip(), path) for part in value.split(":"))
    if first > last:
        raise DeclarationError("bad_columns", path, "the first column comes first, A:F",
                               valid_options=["A:F"])
    return first, last


# ------------------------------------------------------------------------------------------
# Against the workbook
# ------------------------------------------------------------------------------------------


def declare_tables(workbook: "Workbook", declarations: Any) -> list[Region]:
    """Declare the tables of ``declarations`` on ``workbook`` (see the module's notes):
    every one checked against the workbook first, then all kept; the tables declared, in
    order (an optional one whose sheets are not there left out)."""
    specs = declarations if _specs(declarations) else parse_declarations(declarations)
    redacting = [spec for spec in specs if isinstance(spec, RedactTerms)]
    hiding = [spec for spec in specs if isinstance(spec, HiddenSheets)]
    specs = [spec for spec in specs if isinstance(spec, TableSpec)]
    redaction_before = workbook._redaction
    if redacting:
        _redact(workbook, redacting)
    try:
        specs = [_redacted_spec(workbook, spec) for spec in specs]
        before = workbook._withheld_sheets
        withheld = {spec.path: _withheld_names(workbook, spec) for spec in hiding}
        workbook._set_withheld_sheets(before.union(*withheld.values()))
        try:
            _check_hidden_sheets(workbook, specs, withheld)
            with workbook.regions.redeclaring([spec.name for spec in specs]):
                built: list[Region | None] = []
                for spec in specs:
                    built.append(build(workbook, spec, after=_below(built)))
            built = _bounded(workbook, built)
            _check_withheld(workbook, specs, built)
        except BaseException:
            workbook._set_withheld_sheets(before)   # kept whole or not at all
            raise
    except BaseException:
        if workbook._redaction is not redaction_before:
            workbook._set_redaction(redaction_before)
        raise
    return [workbook.regions.keep_table(region) for region in built if region is not None]


def _redact(workbook: "Workbook", specs: "Sequence[RedactTerms]") -> None:
    """Set the redaction ``specs`` declare on ``workbook`` (:meth:`Workbook.redact`):
    before anything else is declared, or the same rules again."""
    redaction = Redaction([spec.rule for spec in specs])
    try:
        workbook.redact(redaction)
    except ValueError:
        raise DeclarationError("redact_late", f"{specs[0].path}.redact", "terms are "
                               "redacted from the workbook's first declarations, before any "
                               "table, region or hidden sheet is kept: declare them first, "
                               "or open the workbook again") from None


def _redacted_spec(workbook: "Workbook", spec: TableSpec) -> TableSpec:
    """``spec`` with the texts it matches the workbook's by -- headers, the key, a header's
    match -- redacted as the workbook's are, so that it still finds them; a table name,
    column id or description of its own that holds a redacted term is refused
    (``redacted_term``), by its path only: the model would read it."""
    redaction = workbook._redaction
    if redaction is None:
        return spec

    def own(text: str | None, path: str) -> None:
        if text and redaction.holds(text):
            raise DeclarationError("redacted_term", path, "holds a term the workbook's "
                                   "redaction replaces, and the model would read it: write it "
                                   "without the term")

    own(spec.name, f"{spec.path}.name")
    own(spec.description, f"{spec.path}.description")
    fields = []
    for field in spec.fields:
        own(field.id, f"{field.path}.id")
        own(field.description, f"{field.path}.description")
        fields.append(replace(field, header=workbook._redacted(field.header))
                      if field.header else field)
    return replace(spec, fields=tuple(fields),
                   key=workbook._redacted(spec.key) if spec.key else spec.key,
                   header_match=tuple(workbook._redacted(t) for t in spec.header_match))


def _bounded(workbook: "Workbook", built: "list[Region | None]") -> "list[Region | None]":
    """``built`` with each part's hidden columns ending just above the header (the first
    row, without one) of the next table declared on the same sheet below it -- among
    those declared together, and those kept already -- so that a section below may show
    the column (:meth:`Span.withheld <xlsx_agent.read.regions.Span.withheld>`); with no
    such table, at the sheet's last row."""
    tops: dict[str, list[int]] = {}
    others = [r for r in workbook.regions.configured
              if r.name.casefold() not in {b.name.casefold() for b in built if b is not None}]
    for region in [*others, *built]:
        if region is None:
            continue
        if region.span is not None:
            for part in region.span.parts:
                top = part.header if part.header is not None else part.first
                tops.setdefault(part.sheet.casefold(), []).append(top)
        else:
            tops.setdefault(region.sheet.casefold(), []).append(region.area.first.row)
    out: list[Region | None] = []
    for region in built:
        if region is None or region.span is None or not region.span.hidden:
            out.append(region)
            continue
        parts = []
        for part in region.span.parts:
            below = [top for top in tops.get(part.sheet.casefold(), []) if top > part.last
                     and top > (part.header or 0)]
            parts.append(replace(part, end=min(below) - 1) if below else part)
        out.append(replace(region, span=replace(region.span, parts=tuple(parts))))
    return out


def _withheld_names(workbook: "Workbook", spec: HiddenSheets) -> set[str]:
    """The sheets ``spec`` withholds from the model, by name folded.  A sheet's name is
    never in a message: one that is no sheet is refused by its path only."""
    every = workbook._every_sheet
    shown = [s.name for s in workbook.sheets]
    out: set[str] = set()
    for n, wanted in enumerate(spec.sheets):
        where = f"{spec.path}.sheets" if len(spec.sheets) == 1 else f"{spec.path}.sheets[{n}]"
        if any(ch in wanted for ch in GLOB):
            pattern = wanted.casefold()
            found = {name for s in every if s.name and _globbed(s, pattern)
                     for name in (s.name.casefold(), s.stored_name.casefold())}
            if not found and not spec.optional:
                raise DeclarationError("no_sheets", where, "the glob matches no sheet",
                                       valid_options=shown[:50])
            out |= found
            continue
        folded = wanted.casefold()
        named = [s for s in every if folded in (s.name.casefold(), s.stored_name.casefold())]
        if named:
            out |= {named[0].name.casefold(), named[0].stored_name.casefold()}
        elif not spec.optional:
            raise DeclarationError("unknown_sheet", where, "no sheet has that name",
                                   valid_options=nearest(wanted, shown))
    return out


def _check_hidden_sheets(workbook: "Workbook", specs: "Sequence[TableSpec]",
                         withheld: "dict[str, set[str]]") -> None:
    """No table or region kept already (and not declared again now) reads a sheet now
    withheld from the model: refused (``on_hidden_sheet``), naming the table and the
    declaration, never the sheet."""
    names = {spec.name.casefold() for spec in specs}
    regions = workbook.regions
    kept = {id(r): r for r in [*regions.configured, *regions.declared]
            if r.name.casefold() not in names}
    for region in kept.values():
        sheets = [p.sheet for p in region.span.parts] if region.span is not None \
            else [region.sheet]
        for path, hidden in withheld.items():
            if any(sheet.casefold() in hidden for sheet in sheets):
                raise DeclarationError("on_hidden_sheet", f"{path}.sheets",
                                       f"{region.address.text} reads a sheet this hides from "
                                       "the model: declare it again without that sheet, or "
                                       "show the sheet")


def _below(built: "list[Region | None]") -> dict[str, int]:
    """The last row of the tables declared so far, by sheet (folded): a header found by
    its text is looked for below them."""
    out: dict[str, int] = {}
    for region in built:
        if region is None:
            continue
        if region.span is not None:
            parts = [(p.sheet, max(p.last, p.header or 0)) for p in region.span.parts]
        else:
            parts = [(region.sheet, region.area.last.row)]
        for sheet, last in parts:
            folded = sheet.casefold()
            out[folded] = max(out.get(folded, 0), last)
    return out


def _check_withheld(workbook: "Workbook", specs: "Sequence[TableSpec]",
                    built: "list[Region | None]") -> None:
    """No table shows a cell another hides from the model: refused (``hidden_elsewhere``),
    naming the column, its sheet and the table that hides it, so the declaration hides it
    too."""
    names = {spec.name.casefold() for spec in specs}
    hiding: list[tuple[str, str, Area]] = [
        (region.address.text, sheet, area) for name, region in
        ((r.name.casefold(), r) for r in workbook.regions.configured)
        if name not in names and region.span is not None
        for sheet, area in region.span.withheld()]
    hiding += [(region.address.text, sheet, area) for region in built
               if region is not None and region.span is not None
               for sheet, area in region.span.withheld()]
    for spec, region in zip(specs, built):
        if region is None:
            continue
        for sheet, area in _shown(region):
            for other, hidden_sheet, hidden in hiding:
                if other != region.address.text and hidden_sheet.casefold() == sheet.casefold() \
                        and _meet(area, hidden):
                    column = column_letters(hidden.first.column)
                    raise DeclarationError("hidden_elsewhere", f"{spec.path}.fields",
                                           f"column {column} of sheet {quote_sheet(sheet)} is "
                                           f"hidden from the model by {other}; hide it here "
                                           "too", valid_options=[column])


def _shown(region: Region) -> list[tuple[str, Area]]:
    """The cells a declared table shows, sheet by sheet, a rectangle per column run."""
    if region.span is None:
        return [(region.sheet, region.area)]
    out = []
    for sheet, area in region.span.areas():
        out += [(sheet, Area(Cell(area.first.row, column), Cell(area.last.row, column)))
                for column in region.span.columns]
    return out


def _meet(one: Area, two: Area) -> bool:
    return not (one.last.row < two.first.row or two.last.row < one.first.row
                or one.last.column < two.first.column or two.last.column < one.first.column)


def _specs(value: Any) -> bool:
    return isinstance(value, (list, tuple)) and bool(value) \
        and all(isinstance(v, (TableSpec, HiddenSheets, RedactTerms)) for v in value)


def build(workbook: "Workbook", spec: TableSpec,
          after: "Mapping[str, int] | None" = None) -> Region | None:
    """The table ``spec`` declares, checked against ``workbook``; ``None`` for an optional
    one whose sheets are not there.  ``after``: the last row of the tables declared before
    it, by sheet (folded), below which a header found by its text is looked for."""
    regions = workbook.regions
    path = spec.path
    folded = spec.name.casefold()
    taken = [t.address.text for t in regions.excel_tables if t.name.casefold() == folded]
    taken += [n.address.text for n in regions.names if n.address is not None
              and n.scope is None and n.name.casefold() == folded]
    if taken:
        raise DeclarationError("name_taken", f"{path}.name", f"{taken[0]} is already a table "
                               "or defined name; a declared table takes a name of its own")
    sheets = _sheets(workbook, spec)
    if not sheets:
        return None
    first_sheet = sheets[0]
    used = workbook.used_range(first_sheet)
    found = None
    if spec.header_match and used is not None:
        across = spec.columns or (used.first.column, used.last.column)
        found = _find_header(workbook, spec, first_sheet, used, *across, after or {})
    left, right = _extent(workbook, spec, first_sheet, used,
                          found[0] if found is not None else spec.header_row)
    parts: list[Part] = []
    flags: list[Problem] = []
    belows: list[dict[int, tuple[int, str]]] = []
    names: list[str] | None = None
    texts: list[str] | None = None
    for sheet in sheets:
        used = workbook.used_range(sheet)
        if used is None:
            raise DeclarationError("empty_sheet", f"{path}.sheets", f"sheet "
                                   f"{quote_sheet(sheet.name)} holds no cells")
        header = spec.header_row or None
        written = quote_sheet(sheet.name)
        stop = None
        if spec.header_match:
            header, stop = found if sheet is first_sheet and found is not None \
                else _find_header(workbook, spec, sheet, used, left, right, after or {})
        if header is not None:
            if header > used.last.row:
                raise DeclarationError("no_header", f"{path}.header_row", f"row {header} of "
                                       f"sheet {written} is below its last cell")
            row = _row_values(workbook, sheet, header, left, right)
            found = [v.v if v.t == "text" and isinstance(v.v, str) else "" for v in row]
            if texts is None:
                texts = found
                names = _header_names(row, left, flags, written)
            elif found != texts:
                at = next(i for i, (a, b) in enumerate(zip(found, texts)) if a != b)
                cell = f"{column_letters(left + at)}{header}"
                raise DeclarationError("header_mismatch", f"{path}.sheets", f"the header of "
                                       f"sheet {written} differs from sheet "
                                       f"{quote_sheet(first_sheet.name)}'s at {cell}: a table "
                                       "over several sheets has one header",
                                       valid_options=[f"{written}!{cell}",
                                                      f"{quote_sheet(first_sheet.name)}!{cell}"])
        first = spec.first_row or (header + 1 if header is not None else used.first.row)
        last, below = _last_row(workbook, spec, sheet, used, first, left, right, stop)
        belows.append(below)
        parts.append(Part(sheet.name, sheet.id, header, first, last))
    if names is None:
        names = [f"Column {column_letters(c)}" for c in range(left, right + 1)]
    if spec.type == "key_value":
        if right - left < 1:
            raise DeclarationError("bad_columns", f"{path}.columns", "a key-value table has "
                                   "two columns at least: a field and a value")
        if spec.header_row == 0:
            names[:2] = list(KEY_VALUE_COLUMNS)
    return _region(workbook, spec, parts, left, right, names, flags, belows)


def _sheets(workbook: "Workbook", spec: TableSpec) -> "list[Sheet]":
    path = f"{spec.path}.sheets"
    every = workbook.sheets
    chosen: set[int] = set()
    cells = ("worksheet", "macrosheet")
    for n, wanted in enumerate(spec.sheets):
        where = path if len(spec.sheets) == 1 else f"{path}[{n}]"
        if any(ch in wanted for ch in GLOB):
            pattern = wanted.casefold()
            found = [i for i, s in enumerate(every) if s.part is not None and s.type in cells
                     and _globbed(s, pattern)]
            if not found:
                if spec.optional:
                    continue
                raise DeclarationError("no_sheets", where, "the glob matches no sheet that "
                                       "holds cells", valid_options=[s.name for s in every][:50])
            chosen.update(found)
            continue
        try:
            sheet = workbook.sheet(wanted)
        except NameNotFound as error:
            stored = [s for s in every if s.stored_name.casefold() == wanted.casefold()]
            if stored:
                sheet = stored[0]
            elif spec.optional:
                continue
            else:
                raise DeclarationError("unknown_sheet", where, "no sheet "
                                       f"{workbook._redacted(wanted)!r}",
                                       valid_options=error.valid_options) from None
        if sheet.part is None or sheet.type not in cells:
            raise DeclarationError("not_cells", where, f"sheet {quote_sheet(sheet.name)} is "
                                   f"a {sheet.type or 'sheet of no known type'}: it holds no cells")
        chosen.add(next(i for i, s in enumerate(every) if s is sheet or s == sheet))
    return [every[i] for i in sorted(chosen)]


def _globbed(sheet: "Sheet", pattern: str) -> bool:
    """Whether a declaration's glob (folded) matches ``sheet``'s name, shown or stored."""
    return fnmatch.fnmatchcase(sheet.name.casefold(), pattern) \
        or fnmatch.fnmatchcase(sheet.stored_name.casefold(), pattern)


def _extent(workbook: "Workbook", spec: TableSpec, sheet: "Sheet",
            used: Area | None, header_row: int) -> tuple[int, int]:
    """The table's sheet columns: declared, or the first sheet's header's (its first cell
    holding something to its last), or its used range's."""
    if used is None:
        raise DeclarationError("empty_sheet", f"{spec.path}.sheets", f"sheet "
                               f"{quote_sheet(sheet.name)} holds no cells")
    if spec.columns is not None:
        return spec.columns
    if header_row and header_row <= used.last.row:
        row = _row_values(workbook, sheet, header_row, used.first.column, used.last.column)
        filled = [v.column for v in row if v.t != "blank"]
        if not filled:
            raise DeclarationError("no_header", f"{spec.path}.header_row", f"row "
                                   f"{header_row} of sheet {quote_sheet(sheet.name)} is "
                                   "blank; give the row of the header, or columns")
        return min(filled), max(filled)
    return used.first.column, used.last.column


def _find_header(workbook: "Workbook", spec: TableSpec, sheet: "Sheet", used: Area | None,
                 left: int, right: int, after: "Mapping[str, int]") -> tuple[int, int | None]:
    """The header row ``spec.header_match`` finds on ``sheet``: the one row, within
    ``header_within`` (else below the tables declared before it on the sheet), whose text
    cells in columns ``left`` to ``right`` hold every text of the match (each a cell's
    whole text, case and runs of white space aside; in any order, other cells besides).
    With it, the next row below that matches too, where the data rows stop (``None``:
    none).  None or several rows in the window: refused, the rows by number only."""
    path = f"{spec.path}.header_row"
    written = quote_sheet(sheet.name)
    if used is None:
        raise DeclarationError("empty_sheet", f"{spec.path}.sheets", f"sheet {written} holds "
                               "no cells")
    if spec.header_within is not None:
        top, bottom = spec.header_within
        where = f"rows {top}:{bottom}"
    else:
        top, bottom = after.get(sheet.name.casefold(), 0) + 1, MAX_SHEET_ROW
        where = f"rows {top}:{used.last.row}" if top > 1 else "any row"
    wanted = {_folded(text) for text in spec.header_match}
    matches: list[int] = []
    start = max(top, used.first.row)
    if start <= used.last.row:
        rows = workbook.read(RangeAddress(SheetRef(name=sheet.name),
                                          Area(Cell(start, left), Cell(used.last.row, right))))
        for values in rows:
            held = {_folded(v.v) for v in values if v.t == "text" and isinstance(v.v, str)}
            if wanted <= held:
                matches.append(values[0].row)
    inside = [row for row in matches if row <= bottom]
    if not inside:
        raise DeclarationError("no_header", path, f"no row of sheet {written} ({where}) holds "
                               "every text of header_row.match, each a cell's whole text; "
                               "check the texts, or header_row.rows")
    if len(inside) > 1:
        shown = ", ".join(str(row) for row in inside[:10])
        raise DeclarationError("ambiguous_header", path, f"{len(inside)} rows of sheet "
                               f"{written} ({where}) hold every text of header_row.match: "
                               f"{shown}; give header_row.rows, or more of the header's texts",
                               valid_options=[str(row) for row in inside[:50]])
    header = inside[0]
    stop = next((row for row in matches if row > header), None)
    return header, stop


def _row_values(workbook: "Workbook", sheet: "Sheet", row: int, left: int, right: int) -> list:
    return next(iter(workbook.read(RangeAddress(SheetRef(name=sheet.name),
                                                Area(Cell(row, left), Cell(row, right))))))


def _last_row(workbook: "Workbook", spec: TableSpec, sheet: "Sheet", used: Area, first: int,
              left: int, right: int, stop: int | None = None
              ) -> tuple[int, dict[int, tuple[int, str]]]:
    """The last data row of ``sheet``, and in each column the first cell holding something
    below the blank row that ended the rows, by sheet column: its row and address (none:
    nothing there, or ``last_row`` declared).  ``stop``: a row the rows end before (the
    next header a header found by its text matches), and what is below it not looked at."""
    if spec.last_row is not None:
        return min(spec.last_row, max(used.last.row, first - 1)), {}
    end = used.last.row if stop is None else min(used.last.row, stop - 1)
    if first > end:
        return first - 1, {}
    last = first - 1
    rows = workbook.read(RangeAddress(SheetRef(name=sheet.name),
                                      Area(Cell(first, left), Cell(end, right))))
    for values in rows:
        if all(v.t == "blank" for v in values):
            break
        last = values[0].row
    below: dict[int, tuple[int, str]] = {}
    for values in rows:
        for value in values:
            if value.t != "blank" and value.column not in below:
                below[value.column] = (value.row, value.a)
    return last, below


def _region(workbook: "Workbook", spec: TableSpec, parts: list[Part], left: int, right: int,
            names: list[str], flags: list[Problem],
            belows: "list[dict[int, tuple[int, str]]]") -> Region:
    path = spec.path
    width = right - left + 1
    ids: list[str | None] = [None] * width
    notes: list[str | None] = [None] * width
    hidden = [False] * width
    cleans: "list[Clean | None]" = [None] * width
    if spec.type == "key_value" and spec.header_row:
        ids[:2] = list(KEY_VALUE_COLUMNS)
    seen: dict[int, str] = {}
    for field in spec.fields:
        if field.column is not None:
            offset = field.column - left
            if not 0 <= offset < width:
                raise DeclarationError("unknown_column", f"{field.path}.column", f"column "
                                       f"{column_letters(field.column)} is outside the "
                                       f"table's, {column_letters(left)}:{column_letters(right)}")
        else:
            offset = next((i for i, name in enumerate(names)
                           if name.casefold() == field.header.casefold()), None)
            if offset is None:
                shown = [n for i, n in enumerate(names) if not _hidden_by(spec, i, left, names)]
                raise DeclarationError("unknown_column", f"{field.path}.header",
                                       f"no column has the header {field.header!r}",
                                       valid_options=nearest(field.header, shown))
        if offset in seen:
            raise DeclarationError("duplicate_field", field.path, f"column "
                                   f"{column_letters(left + offset)} is declared twice "
                                   f"(first at {seen[offset]})")
        seen[offset] = field.path
        if field.id is not None:
            ids[offset] = field.id
        notes[offset] = field.description
        hidden[offset] = field.hidden
        cleans[offset] = field.clean
    columns = [ids[i] or names[i] for i in range(width)]
    visible = [i for i in range(width) if not hidden[i]]
    if not visible:
        raise DeclarationError("no_columns", f"{path}.fields", "every column is hidden")
    taken: dict[str, int] = {}
    for i in visible:
        for name in dict.fromkeys([columns[i].casefold(), names[i].casefold()]):
            if name in taken and taken[name] != i:
                which = spec.fields[_field_of(spec, i, left, names)].path \
                    if _field_of(spec, i, left, names) is not None else f"{path}.fields"
                raise DeclarationError("id_taken", f"{which}.id", f"{columns[i]!r} is another "
                                       "column's name or header: a column is found by "
                                       "either, so each names one column")
            taken[name] = i
    if spec.type == "key_value" and (hidden[0] or hidden[1]):
        raise DeclarationError("hidden_key", f"{path}.fields", "a key-value table's field "
                               "and value columns are never hidden")
    key = None
    if spec.key is not None:
        wanted = spec.key.casefold()
        found = next((i for i in range(width)
                      if wanted in (columns[i].casefold(), names[i].casefold())), None)
        if found is None or hidden[found]:
            if found is not None:
                raise DeclarationError("hidden_key", f"{path}.key", "the key column is hidden "
                                       "from the model; keys are in addresses")
            raise DeclarationError("unknown_column", f"{path}.key", f"no column "
                                   f"{spec.key!r}", valid_options=nearest(
                                       spec.key, [columns[i] for i in visible]))
        key = columns[found]
    elif spec.type == "key_value":
        key = columns[0]
    if key is not None:
        at = next(i for i in range(width) if columns[i] == key)
        if cleans[at] is not None:
            field_path = spec.fields[_field_of(spec, at, left, names)].path
            raise DeclarationError("clean_key", f"{field_path}.clean", "the key column is "
                                   "not cleaned: an address by key names the key as its cell "
                                   "holds it, and two keys cleaned alike would name one row")
    sheet_columns = tuple(left + i for i in visible)
    # What the table says of its cells names no cell of a column hidden from the model:
    # a header flag there is dropped, and more below is flagged at its first visible cell.
    flags = [f for f in flags if f.address is None
             or _column_of(f.address) in (None, *sheet_columns)]
    for below in belows:
        shown = sorted((row, column, address) for column, (row, address) in below.items()
                       if column in sheet_columns)
        if shown:
            flags.append(Problem("rows_below", shown[0][2]))
    header_rows = 1 if spec.header_row else 0
    common = dict(kind="declared", name=spec.name, columns=tuple(columns[i] for i in visible),
                  header_rows=header_rows, key=key, flags=tuple(dict.fromkeys(flags)),
                  headers=tuple(names[i] for i in visible),
                  notes=tuple(notes[i] for i in visible), about=spec.description,
                  shape=spec.type, workbook=workbook)
    hidden_columns = tuple(left + i for i in range(width) if hidden[i])
    rules = tuple(cleans[i] for i in visible) if any(cleans) else ()
    if len(parts) == 1 and not hidden_columns and not rules:
        part = parts[0]
        top = part.header if part.header is not None else part.first
        if part.last < top:
            raise DeclarationError("no_rows", path, f"sheet {quote_sheet(part.sheet)} has no "
                                   "data rows, and the table no header row")
        region = Region(sheet=part.sheet, sheet_id=part.sheet_id,
                        area=Area(Cell(top, left), Cell(max(part.last, top), right)), **common)
    else:
        if not header_rows and not any(p.rows for p in parts):
            raise DeclarationError("no_rows", path, "the sheets have no data rows, and the "
                                   "table no header row")
        span = Span(name=f"t:{spec.name}", parts=tuple(parts), columns=sheet_columns,
                    header_rows=header_rows, left=left, right=right, hidden=hidden_columns,
                    clean=rules)
        region = Region(sheet=span.name, area=span.area, span=span, **common)
    if key is not None:
        if region.span is not None:
            with workbook.regions.reading(region.span):
                _check_keys(workbook, region, path)
        else:
            _check_keys(workbook, region, path)
    return region


def _column_of(address: str) -> int | None:
    """The sheet column of a cell's address (``'Area 1'!E3``), ``None`` for another form."""
    letters = address.rpartition("!")[2].rstrip("0123456789").lstrip("$")
    try:
        return column_index(letters) if letters.isalpha() else None
    except (ValueError, KeyError):
        return None


def _hidden_by(spec: TableSpec, offset: int, left: int, names: list[str]) -> bool:
    found = _field_of(spec, offset, left, names)
    return found is not None and spec.fields[found].hidden


def _field_of(spec: TableSpec, offset: int, left: int, names: list[str]) -> int | None:
    for n, field in enumerate(spec.fields):
        if field.column is not None and field.column - left == offset:
            return n
        if field.header is not None and field.header.casefold() == names[offset].casefold():
            return n
    return None


def _check_keys(workbook: "Workbook", region: Region, path: str) -> None:
    """Every data row has a key no other row has: refused otherwise, naming the cells."""
    index = workbook.regions.key_index(region)
    offset = region.column_index(region.key)

    def cell(position: int) -> str:
        row = region.data.first.row + position - 1
        column = region.area.first.column + offset
        return region.located(Area(Cell(row, column), Cell(row, column)))

    missing = [p for p, k in enumerate(index.keys, 1) if k is None]
    if missing:
        raise DeclarationError("no_key", f"{path}.key", f"{len(missing)} rows have no key in "
                               f"column {region.key!r} (blank, or not text or a whole "
                               f"number), the first at {cell(missing[0])}",
                               valid_options=[cell(p) for p in missing[:50]])
    if index.duplicates:
        rows = next(iter(index.duplicates.values()))
        count = len(index.duplicates)
        which = "a key" if count == 1 else f"{count} keys"
        are = "is" if count == 1 else "are"
        raise DeclarationError("duplicate_key", f"{path}.key", f"{which} of column "
                               f"{region.key!r} {are} on more than one row: "
                               f"{cell(rows[0])} and {cell(rows[1])} hold the same",
                               valid_options=[cell(p) for p in rows[:50]])


__all__ = ["DeclarationError", "FieldSpec", "HiddenSheets", "MAX_DESCRIPTION", "MAX_ID",
           "RedactTerms",
           "TableSpec", "TYPES", "build", "declare_tables", "parse_declarations"]
