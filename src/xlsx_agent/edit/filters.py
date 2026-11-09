"""AutoFilters and hidden rows and columns (production round 44, R55).

::

    wb.set_autofilter("Stock", "A1:F8", {"Item": ["Lamp", "Desk"], "Shelf": [None]})
    wb.set_autofilter("t:Orders", criteria={"Item": "Lamp"})     # a table's own filter
    wb.set_autofilter("Stock!A1:F8")                              # the buttons, no criteria
    wb.clear_autofilter("Stock")                  # the filter off: every row shown again
    wb.clear_autofilter("Stock", keep=True)       # Excel's Clear: criteria gone, buttons kept
    wb.hide_rows("Stock", "5:7"); wb.show_rows("Stock", [5, "9:10"])
    wb.hide_columns("Stock", "C:D"); wb.show_columns("Stock", "C")

Each is one undoable step (or part of the open batch), every problem refused before
anything changes, and changes no value.  ``set_autofilter`` sets the filter whole: its
``criteria`` replace those of every column of the range or table, so the columns to filter
are passed together in one call (Excel's own object model sets one column per call and
keeps the others' -- round 44's two-column probe -- which a mapping of columns does not
need).

**What Excel does, measured** (Excel for Mac 16.106, region Dutch (Belgium), 2026-11-08:
its object model's ``autofilter range`` with ``operator filter by value``, ``autofilter
mode`` false and ``show all data`` on a made-up sheet, saved and read; the probes are kept
with the round's notes):

* A sheet holds one AutoFilter: ``autoFilter ref`` (after ``sheetData``), a hidden name
  ``_xlnm._FilterDatabase`` local to the sheet naming the same range anchored
  (``Stock!$A$1:$F$8``), and while a column has criteria ``sheetPr filterMode="1"``.  The
  buttons alone (no criteria) are ``autoFilter ref`` and the name, no ``filterMode``.
* A table's filter is its own ``autoFilter`` in the table part: no name, no
  ``filterMode`` (measured), the rows it excludes hidden on the sheet all the same.
* A value list is ``filterColumn colId`` (the column's place in the range, from 0) with
  ``filters``: one ``filter val`` per value, sorted as text (``Desk``, ``Lamp``; ``12``
  before ``3``), ``blank="1"`` when the blank cells are kept (alone: ``filters
  blank="1"``).  Two columns are two ``filterColumn``; a row is shown when it passes all.
* A value is the cell's text as Excel shows it: text compared without regard to case
  (``Lamp`` kept ``lamp``), a number as its format shows it in the region (``1.5`` in
  ``0.00`` matched ``1,50`` in the Dutch region, not ``1.50`` nor ``1.5``; ``3`` and
  ``12`` in ``General`` matched as such), a formula by its cached value.  Here a number's
  shown text is the workbook's locale's (:attr:`Workbook.locale
  <xlsx_agent.Workbook.locale>`, :func:`xlsx_agent.shown.shown_text`); a number criterion
  (``12``) matches the cells holding that number, and the text written is theirs.
* Applying a filter hides the range's rows it excludes (``row hidden="1"``) and shows
  the others, rows hidden by hand among them (two hidden rows of the range were shown by
  the filter, and by the buttons alone; a hidden row below the range stayed hidden).
  Turning the filter off removes ``autoFilter`` and ``filterMode`` and shows every row of
  the range -- the ``_FilterDatabase`` name stays, as Excel left it; Excel's Clear (``show
  all data``, here ``keep=True``) leaves ``autoFilter ref`` without criteria.
* **Excel does not apply a saved filter when it opens a file**: a file whose criteria
  exclude rows that are not hidden opened, and was saved again, with every row shown,
  the criteria and ``filterMode`` kept (sheet and table filters alike).  So the rows a
  filter here excludes are hidden here, and the result says so (``hidden_by``); with
  ``hide=False`` only the criteria are written and the rows are left as they are, for
  Excel's Reapply (or a person) to hide.

**Literal values only, cached values where they hold.**  Which rows a criterion keeps is
read from the cells.  A cell whose shown text is not measured is refused
(``filter_unmeasured``): a date (Excel's filter took neither ``1/5/2026`` nor ``5-1-2026``
for a date shown so), a boolean or an error (shown in the application's language), a
number in a format :mod:`xlsx_agent.shown` does not render, or a formula's empty text.  A
formula is read by its cached value only while the workbook's caches hold as saved
(:attr:`Workbook.calculation_status <xlsx_agent.Workbook.calculation_status>`
``cached_unverified``); once any cell was written, or when the file asks to be recalculated
(``invalidated``), a criterion on a column holding formulas is refused
(``filter_formulas``) -- ``hide=False`` writes the criteria and leaves the hiding to Excel.

**What hiding rows changes.**  ``SUBTOTAL`` and ``AGGREGATE`` leave hidden rows out, so when
rows are hidden or shown and any sheet holds one (a table's totals row does), the workbook
is asked to recalculate on open (``fullCalcOnLoad``; ``calculation`` says so).

**Rows and columns by hand**: ``hide_rows``/``show_rows`` set and take away ``hidden`` on
the rows (a row with nothing else is written as a row element holding only that, and taken
away again when shown); columns as ``cols`` runs with ``hidden="1"`` and the sheet's
default width (Excel writes a width with a column's other attributes).  On a protected sheet
filtering needs the protection's ``autoFilter`` allowed, hiding rows ``formatRows`` and
columns ``formatColumns``, as Excel's Protect Sheet options say.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Sequence

from lxml import etree
from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, Area, Cell, column_letters, number_text

from ..address import AddressError, RangeAddress, format_range, parse_address, quote_sheet
from . import EditError
from .format import Problems, Target, guard
from .sheets import _X, find_sheet, placed

#: The most rows, columns or values listed in a result or a refusal.
MAX_LISTED = 50
#: The hidden name Excel gives a sheet's filter range (ECMA-376 Part 1, 18.2.6).
FILTER_NAME = "_xlnm._FilterDatabase"
#: What a result says hid the rows (:attr:`FilterChange.hidden_by`).
HIDDEN_HERE = "xlsx-agent"
NOT_APPLIED = ("Excel does not apply a saved filter when it opens a file (measured): the "
               "rows it excludes were hidden here")
LEFT_TO_EXCEL = ("hide=False: only the criteria were written; Excel does not apply them on "
                 "open (measured), so no row is hidden until Excel's Reapply")
#: CT_Workbook's children after ``definedNames`` (ECMA-376 Part 1, 18.2.27).
_AFTER_NAMES = ("calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
                "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects",
                "extLst")
#: Words in a formula that make its value depend on which rows are hidden.
_HIDDEN_AWARE = re.compile(rb"(?i)\b(?:SUBTOTAL|AGGREGATE)\s*\(")
#: A table's totals row function other than a custom formula (written as SUBTOTAL).
_TOTALS = re.compile(rb"""totalsRowFunction\s*=\s*["'](?!custom|none)""")


@dataclass(frozen=True)
class FilterChange:
    """What a filter call did: the filter's ``target`` (a sheet or ``t:Table``) and
    ``range``; the ``criteria`` written, by column header, each a list of the values
    written (``None`` for blank cells); the rows it excludes, hidden (``hidden``, as runs,
    ``"5"``, ``"7:9"``), and the rows it showed that were hidden (``shown``); who hid them (``hidden_by``: ``"xlsx-agent"``, or ``None`` with
    ``hide=False`` or when nothing is hidden); ``unmatched`` the values no cell matched;
    ``calculation`` the workbook's calculation status after; ``notes``."""

    action: str
    target: str
    range: str | None
    changed: bool = True
    criteria: dict[str, list[Any]] = field(default_factory=dict)
    hidden: tuple[str, ...] = ()
    shown: tuple[str, ...] = ()
    hidden_by: str | None = None
    unmatched: dict[str, list[Any]] = field(default_factory=dict)
    calculation: str = "cached_unverified"
    notes: tuple[str, ...] = ()

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"action": self.action, "target": self.target,
                               "range": self.range, "changed": self.changed}
        if self.criteria:
            out["criteria"] = {k: list(v[:limit]) for k, v in self.criteria.items()}
        for name in ("hidden", "shown"):
            runs = getattr(self, name)
            if runs:
                out[name] = list(runs[:limit])
                if len(runs) > limit:
                    out[name + "_cut"] = len(runs) - limit
        if self.hidden_by:
            out["hidden_by"] = self.hidden_by
        if self.unmatched:
            out["unmatched"] = {k: list(v[:limit]) for k, v in self.unmatched.items()}
        out["calculation"] = self.calculation
        if self.notes:
            out["notes"] = list(self.notes)
        return out


@dataclass(frozen=True)
class HiddenChange:
    """What ``hide_rows``/``show_rows``/``hide_columns``/``show_columns`` did: the rows or
    columns ``changed`` (as runs), those ``already`` so, the calculation status, notes."""

    action: str
    sheet: str
    changed: tuple[str, ...] = ()
    already: tuple[str, ...] = ()
    calculation: str = "cached_unverified"
    notes: tuple[str, ...] = ()

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        out: dict[str, Any] = {"action": self.action, "sheet": self.sheet,
                               "changed": list(self.changed[:limit])}
        if len(self.changed) > limit:
            out["changed_cut"] = len(self.changed) - limit
        if self.already:
            out["already"] = list(self.already[:limit])
        out["calculation"] = self.calculation
        if self.notes:
            out["notes"] = list(self.notes)
        return out


# -- finding the filter ----------------------------------------------------------------------


@dataclass
class _Filter:
    """Where a filter lives: the sheet, its range (``None`` while a sheet has none), and
    the element holding its ``autoFilter`` (the worksheet's or the table part's root)."""

    sheet: Any
    label: str
    target: str
    area: Area | None
    holder: Any
    holder_part: str
    table: Any = None
    headers: list[str] = field(default_factory=list)


def _sheet_label(sheet: Any) -> str:
    from .structure import _label

    return _label(sheet)


def _locate(workbook: Any, target: Any, ref: Any) -> _Filter:
    """The filter ``target`` (a sheet, ``t:Table``, or a range with its sheet) and ``ref``
    name; refused with the way out."""
    from .tables import _target

    if not isinstance(target, str) or not target.strip():
        raise EditError("filter_range", "the target is a sheet (Stock), a table (t:Orders) or "
                        "a range with its sheet (Stock!A1:F8)")
    text = target.strip()
    if text.startswith("t:") or ("!" not in text and _is_table(workbook, text, ref)):
        if ref is not None:
            raise EditError("filter_range", "a table's filter is over the table: leave ref "
                            "out", addresses=[text[:80]])
        table, _, _ = _target(workbook, text if text.startswith("t:") else "t:" + text)
        assert table is not None
        return _table_filter(workbook, table)
    if "!" in text and ref is None:
        try:
            parsed = parse_address(text)
        except (AddressError, ValueError):
            parsed = None
        if isinstance(parsed, RangeAddress) and parsed.sheet is not None \
                and parsed.span == "cells":
            sheet = workbook.sheet_of(parsed)
            return _sheet_filter(workbook, sheet, parsed.area)
    # A sheet is found by its name as given, spaces too (a sheet may be called "Stock "),
    # or as an address writes it ('Stock '!).
    sheet = find_sheet(workbook, target)
    area = None
    if ref is not None:
        if not isinstance(ref, str):
            raise EditError("filter_range", "ref is a range, as A1:F8", addresses=[text[:80]])
        try:
            area = Area.parse(ref.strip().replace("$", ""))
        except ValueError:
            raise EditError("filter_range", f"ref is a range, as A1:F8, not {ref[:40]!r}",
                            addresses=[text[:80]]) from None
    return _sheet_filter(workbook, sheet, area)


def _is_table(workbook: Any, text: str, ref: Any) -> bool:
    """Whether ``text`` names a table rather than a sheet.  A name that is both is the
    sheet when ``ref`` is given or the sheet has a filter of its own, else the table."""
    tables = {t.name.casefold() for t in workbook._table_parts()}
    if text.casefold() not in tables:
        return False
    try:
        sheet = find_sheet(workbook, text)
    except Exception:
        return True
    if ref is not None:
        return False
    root = workbook.package.tree(sheet.part) if sheet.part else None
    if root is not None and root.find(_X + "autoFilter") is not None:
        raise EditError("filter_range", f"{text!r} is a table and a sheet with a filter: say "
                        f"which, t:{text} or {quote_sheet(text)}", addresses=[text[:80]])
    return True


def _sheet_filter(workbook: Any, sheet: Any, area: Area | None) -> _Filter:
    if sheet.type != "worksheet" or sheet.part is None:
        raise EditError("not_a_worksheet", addresses=[sheet.stored_name])
    root = workbook.package.tree(sheet.part)
    label = _sheet_label(sheet)
    current = root.find(_X + "autoFilter")
    if area is None and current is not None:
        try:
            area = Area.parse(current.get("ref", ""), lenient=True)
        except ValueError:
            area = None
    return _Filter(sheet=sheet, label=label, target=label, area=area, holder=root,
                   holder_part=sheet.part)


def _table_filter(workbook: Any, facts: Any) -> _Filter:
    if facts.header_rows < 1:
        raise EditError("filter_range", f"table {facts.name} has no header row, so no filter",
                        addresses=[f"t:{facts.name}"])
    first = facts.area.first
    last = Cell(facts.area.last.row - facts.totals_rows, facts.area.last.column)
    if last.row <= first.row:
        raise EditError("filter_range", f"table {facts.name} has no data rows to filter",
                        addresses=[f"t:{facts.name}"])
    return _Filter(sheet=facts.sheet, label=_sheet_label(facts.sheet),
                   target=f"t:{facts.name}", area=Area(first, last), holder=facts.root,
                   holder_part=facts.part, table=facts, headers=list(facts.columns))


# -- reading cells ---------------------------------------------------------------------------


_UNMEASURED = object()


def _shown(workbook: Any, writer: Any, cell: Cell, *, formulas_valid: bool) -> Any:
    """What Excel's filter compares for ``cell``: ``None`` blank, ``("text", text)``,
    ``("number", value, shown)``; ``_UNMEASURED`` (with why) when it is not measured;
    ``("formula", ...)`` refusals are the caller's."""
    from ..shown import shown_text

    content = writer.content(cell)
    if content.formula is not None and not formulas_valid:
        return ("stale", None)
    kind = content.kind
    if kind == "blank":
        if content.formula is not None:
            return ("stale", None)          # a formula without a cached value
        return None
    if kind == "text":
        text = str(content.value)
        if text == "":
            return (_UNMEASURED, "empty text")
        return ("text", text)
    if kind == "number":
        code = workbook.styles.format(content.style).code
        shown = shown_text(content.value, code, workbook.locale)
        if shown is None:
            fmt = workbook.styles.format(content.style)
            why = "a date or time" if fmt.kind not in ("number", "general", "text", None) \
                else "a number whose shown text is not measured"
            return (_UNMEASURED, why)
        return ("number", content.value, shown)
    return (_UNMEASURED, "a boolean" if kind == "bool" else "an error")


def _header_text(workbook: Any, writer: Any, cell: Cell) -> str | None:
    content = writer.content(cell)
    if content.kind == "text":
        return str(content.value)
    if content.kind == "number":
        from ..shown import shown_text

        return shown_text(content.value, workbook.styles.format(content.style).code,
                          workbook.locale)
    return None


# -- criteria --------------------------------------------------------------------------------


def _columns(workbook: Any, found: _Filter, writer: Any) -> list[tuple[int, str]]:
    """Each column of the range: its sheet column number and its header's text (the
    table's column name; a sheet's header cell, or its letters when it has none)."""
    assert found.area is not None
    out = []
    for offset, number in enumerate(range(found.area.first.column,
                                          found.area.last.column + 1)):
        if found.table is not None:
            name = found.headers[offset] if offset < len(found.headers) else ""
        else:
            name = _header_text(workbook, writer, Cell(found.area.first.row, number)) or ""
        out.append((number, name))
    return out


def _resolve_column(key: Any, columns: list[tuple[int, str]], problems: Problems,
                    notes: list[str]) -> int | None:
    """The sheet column a criteria key names: a header's text (case aside), else the
    column's letters (``"C"``)."""
    headers = [name for _, name in columns]
    if isinstance(key, str) and key.strip():
        wanted = key.strip().casefold()
        for number, name in columns:
            if name.strip().casefold() == wanted:
                return number
        letters = key.strip().upper()
        if letters.isalpha() and len(letters) <= 3:
            try:
                number = Cell.parse(letters + "1").column
            except ValueError:
                number = 0
            if any(n == number for n, _ in columns):
                return number
    shown = [f"{name or '(no header)'} ({column_letters(n)})" for n, name in columns]
    problems.add("filter_column", str(key)[:40], "a criteria key is a column's header or its "
                 f"letters; the filter's columns are {shown[:MAX_LISTED]}"
                 + ("" if isinstance(key, str) else " (a number is not taken: 0- or 1-based "
                    "is not clear)"))
    del headers
    return None


@dataclass
class _Criterion:
    column: int
    header: str
    texts: list[str]
    numbers: list[float]
    blanks: bool
    given: list[Any]


def _criteria(criteria: Any, columns: list[tuple[int, str]], problems: Problems,
              notes: list[str]) -> list[_Criterion]:
    if criteria is None:
        return []
    if not isinstance(criteria, dict):
        problems.add("bad_value", "criteria", "criteria is {column: [values]}, as "
                     "{\"Item\": [\"Lamp\", \"Desk\"]}; None for blank cells")
        return []
    out: dict[int, _Criterion] = {}
    for key, values in criteria.items():
        number = _resolve_column(key, columns, problems, notes)
        if number is None:
            continue
        if isinstance(values, (str, int, float)) or values is None:
            values = [values]
        if not isinstance(values, (list, tuple)):
            problems.add("bad_value", str(key)[:40], "a column's criteria is a list of "
                         "values (text or numbers; None for blank cells)")
            continue
        if not values:
            notes.append(f"{key}: an empty list sets no criterion on that column")
            continue
        header = dict(columns)[number] or column_letters(number)
        if number in out:
            problems.add("bad_value", str(key)[:40], f"column {header} is named twice")
            continue
        item = _Criterion(number, header, [], [], False, list(values))
        for value in values:
            if value is None or value == "":
                item.blanks = True
            elif isinstance(value, bool):
                problems.add("filter_unmeasured", str(key)[:40], "a boolean criterion: how "
                             "Excel's filter shows TRUE and FALSE is the application's "
                             "language, not measured")
            elif isinstance(value, (int, float)):
                item.numbers.append(float(value))
            elif isinstance(value, str):
                item.texts.append(value)
            else:
                problems.add("bad_value", str(key)[:40], f"a criterion is text, a number or "
                             f"None, not {type(value).__name__}")
        out[number] = item
    return [out[n] for n in sorted(out)]


# -- the calls -------------------------------------------------------------------------------


def set_autofilter(workbook: Any, target: Any, ref: Any = None, criteria: Any = None, *,
                   hide: bool = True) -> FilterChange:
    """Set the AutoFilter of a sheet (over ``ref``, or the one it has) or of a table, with
    value-list ``criteria`` (see the module); the rows it excludes hidden unless
    ``hide`` is false."""
    if workbook.strict:
        raise EditError("strict")
    if not isinstance(hide, bool):
        raise EditError("bad_value", "hide is True (the rows the filter excludes hidden "
                        "here) or False (criteria only)")
    found = _locate(workbook, target, ref)
    if found.area is None:
        raise EditError("filter_range", f"{found.label} has no filter yet: give its range, "
                        f"as wb.set_autofilter({found.label!r}, \"A1:F20\") -- the header row "
                        "and the rows below it", addresses=[found.label])
    area = found.area
    shown_range = format_range(found.label, area)
    problems = Problems()
    notes: list[str] = []
    if found.table is None:
        _check_sheet_range(workbook, found, problems)
    guard(workbook, [Target(found.sheet, area, "cells", shown_range)], problems,
          protected="autoFilter")
    writer = workbook.editor.cells.sheet(found.sheet.stored_name)
    if writer is None:
        raise EditError("not_a_worksheet", addresses=[shown_range])
    columns = _columns(workbook, found, writer)
    wanted = _criteria(criteria, columns, problems, notes)
    formulas_valid = workbook.calculation_status != "invalidated"
    rows = range(area.first.row + 1, area.last.row + 1)
    keep: dict[int, bool] = {}
    written: dict[int, list[str]] = {}
    unmatched: dict[str, list[Any]] = {}
    if wanted:
        stale: list[str] = []
        unmeasured: dict[str, list[str]] = {}
        seen: dict[int, dict[str, str]] = {c.column: {} for c in wanted}
        matched: dict[int, set[int]] = {c.column: set() for c in wanted}
        for row in rows:
            passes = True
            for item in wanted:
                cell = Cell(row, item.column)
                read = _shown(workbook, writer, cell, formulas_valid=formulas_valid)
                address = format_range(found.label, Area(cell, cell))
                if read is not None and read[0] == "stale":
                    stale.append(address)
                    continue
                if read is not None and read[0] is _UNMEASURED:
                    unmeasured.setdefault(read[1], []).append(address)
                    continue
                ok = _passes(item, read, seen[item.column], matched[item.column])
                passes = passes and ok
            keep[row] = passes
        if stale and hide:
            why = ("the workbook was written since it was opened" if _written(workbook)
                   else "the file asks to be recalculated when it opens")
            problems.add("filter_formulas", stale[0], f"{len(stale)} formula cell(s) in the "
                         f"filtered columns, whose cached values may be stale ({why}): pass "
                         "hide=False to write the criteria and leave hiding the rows to "
                         "Excel's Reapply, or filter on a column of values")
        if hide:
            for why, cells in unmeasured.items():
                problems.add("filter_unmeasured", cells[0], f"{len(cells)} cell(s) hold {why}, "
                             "whose text Excel's filter compares is not measured here: pass "
                             "hide=False to write the criteria and leave hiding the rows to "
                             "Excel, or filter on another column")
        for item in wanted:
            values, missing = _written_values(workbook, item, seen[item.column],
                                              matched[item.column], problems)
            written[item.column] = values
            if missing:
                unmatched[item.header] = missing
    problems.raise_any()
    if not hide:
        keep = {}
    return _apply(workbook, found, area, wanted, written, keep, hide, unmatched, notes)


def _check_sheet_range(workbook: Any, found: _Filter, problems: Problems) -> None:
    area = found.area
    assert area is not None
    shown = format_range(found.label, area)
    if area.last.row <= area.first.row:
        problems.add("filter_range", shown, "a filter is its header row and at least one "
                     "row below it")
    from .tables import _meets

    clash = [t for t in workbook._table_parts()
             if t.sheet == found.sheet.name and t.area is not None and _meets(t.area, area)]
    if clash:
        problems.add("filter_table", shown, "the range overlaps a table, which has its own "
                     "filter: " + ", ".join(f"wb.set_autofilter('t:{t.name}', ...)"
                                            for t in clash[:5]))


def _passes(item: _Criterion, read: Any, seen: dict[str, str], matched: set[int]) -> bool:
    """Whether the cell read passes ``item``; the shown texts of cells a value matched
    are kept (``seen``: folded value to the text written), and which values matched."""
    if read is None:
        return item.blanks
    if read[0] == "text":
        folded = read[1].casefold()
        for index, text in enumerate(item.texts):
            if text.casefold() == folded:
                seen.setdefault(folded, read[1])
                matched.add(index)
                return True
        return False
    _, value, shown = read
    folded = shown.casefold()
    for index, text in enumerate(item.texts):
        if text.casefold() == folded:
            seen.setdefault(folded, shown)
            matched.add(index)
            return True
    for index, number in enumerate(item.numbers):
        if float(value) == number:
            seen.setdefault(folded, shown)
            matched.add(len(item.texts) + index)
            return True
    return False


def _written_values(workbook: Any, item: _Criterion, seen: dict[str, str],
                    matched: set[int], problems: Problems) -> tuple[list[str], list[Any]]:
    """The ``filter val`` texts of a column, sorted as Excel sorted them (as text), and
    the values no cell matched."""
    from ..shown import shown_text

    out: dict[str, str] = {}
    missing: list[Any] = []
    for index, text in enumerate(item.texts):
        folded = text.casefold()
        out.setdefault(folded, seen.get(folded, text))
        if index not in matched:
            missing.append(text)
    for index, number in enumerate(item.numbers):
        if len(item.texts) + index in matched:
            continue
        missing.append(int(number) if number.is_integer() else number)
        shown = shown_text(int(number) if number.is_integer() else number, "General",
                           workbook.locale)
        if shown is None:
            problems.add("bad_value", item.header, f"{number!r} matches no cell, and how "
                         "Excel would write it is not measured: give it as text")
            continue
        out.setdefault(shown.casefold(), shown)
    for folded, text in seen.items():
        out.setdefault(folded, text)
    return sorted(out.values(), key=lambda t: (t.casefold(), t)), missing


def _written(workbook: Any) -> bool:
    editor = workbook._editor
    return editor is not None and bool(editor.written())


def _apply(workbook: Any, found: _Filter, area: Area, wanted: list[_Criterion],
           written: dict[int, list[str]], keep: dict[int, bool], hide: bool,
           unmatched: dict[str, list[Any]], notes: list[str]) -> FilterChange:
    editor = workbook.editor
    package = workbook.package
    shown_range = format_range(found.label, area)
    hidden_rows: list[int] = []
    shown_rows: list[int] = []
    before = package.read(found.holder_part)
    with editor.batch():
        writer = editor.cells.sheet(found.sheet.stored_name)
        node = found.holder.find(_X + "autoFilter")
        old_area = None
        if node is not None:
            try:
                old_area = Area.parse(node.get("ref", ""), lenient=True)
            except ValueError:
                old_area = None
        if node is None or old_area is None or old_area.ref != area.ref:
            fresh = etree.Element(_X + "autoFilter")
            if node is not None:
                node.addprevious(fresh)
                found.holder.remove(node)
            elif found.table is not None:
                found.holder.insert(0, fresh)
            else:
                placed(found.holder, fresh)
            node = fresh
            if old_area is not None and old_area.ref != area.ref and hide:
                # the old filter off: its rows shown again, as turning it off shows them
                for row in range(old_area.first.row + 1, old_area.last.row + 1):
                    if not area.first.row < row <= area.last.row and _set_hidden(writer, row,
                                                                                   False):
                        shown_rows.append(row)
                notes.append(f"the filter over {format_range(found.label, old_area)} was "
                             "replaced: a sheet has one")
        node.set("ref", area.ref)
        for child in list(node):
            if child.tag == _X + "filterColumn":
                node.remove(child)
        anchor = node.find(_X + "sortState")
        for item in wanted:
            column = etree.Element(_X + "filterColumn")
            column.set("colId", str(item.column - area.first.column))
            filters = etree.SubElement(column, _X + "filters")
            if item.blanks:
                filters.set("blank", "1")
            for text in written[item.column]:
                etree.SubElement(filters, _X + "filter").set("val", text)
            if anchor is not None:
                anchor.addprevious(column)
            else:
                extensions = node.find(_X + "extLst")
                if extensions is not None:
                    extensions.addprevious(column)
                else:
                    node.append(column)
        package.mark_dirty(found.holder_part)
        if found.table is None:
            _filter_mode(workbook, found.sheet, bool(wanted))
            _filter_name(workbook, found.sheet, area)
        rows_changed: list[int] = []
        if hide:
            for row in range(area.first.row + 1, area.last.row + 1):
                passes = keep.get(row, True)
                if not passes:
                    hidden_rows.append(row)
                if _set_hidden(writer, row, not passes):
                    rows_changed.append(row)
                    if passes:
                        shown_rows.append(row)
        rows_changed = sorted(set(rows_changed) | set(shown_rows))
        if rows_changed:
            _touched_rows(editor, found.sheet, rows_changed)
            _recalculate(workbook)
        else:
            _touched_rows(editor, found.sheet, [area.first.row])
        workbook._forget()
    changed = package.read(found.holder_part) != before or bool(rows_changed)
    criteria_out = {item.header: (list(written[item.column]) + ([None] if item.blanks else []))
                    for item in wanted}
    if not hide and wanted:
        notes.append(LEFT_TO_EXCEL)
    elif hidden_rows:
        notes.append(NOT_APPLIED)
    if not changed:
        notes.append("the filter is already so")
    return FilterChange("set_autofilter", found.target, shown_range, changed=changed,
                        criteria=criteria_out, hidden=tuple(_runs(sorted(hidden_rows))),
                        shown=tuple(_runs(sorted(shown_rows))),
                        hidden_by=HIDDEN_HERE if hidden_rows else None, unmatched=unmatched,
                        calculation=workbook.calculation_status, notes=tuple(notes))


def clear_autofilter(workbook: Any, target: Any, *, keep: bool = False) -> FilterChange:
    """Turn a sheet's filter off (its range's rows shown, ``autoFilter`` and
    ``filterMode`` removed, the hidden name kept as Excel keeps it), or with ``keep`` clear
    its criteria and keep its buttons (Excel's Clear); a table's criteria cleared (its
    buttons stay).  No filter, or none with criteria: ``changed`` false."""
    if workbook.strict:
        raise EditError("strict")
    found = _locate(workbook, target, None)
    if found.area is None:
        return FilterChange("clear_autofilter", found.target, None, changed=False,
                            calculation=workbook.calculation_status,
                            notes=("the sheet has no filter",))
    area = found.area
    shown_range = format_range(found.label, area)
    problems = Problems()
    guard(workbook, [Target(found.sheet, area, "cells", shown_range)], problems,
          protected="autoFilter")
    problems.raise_any()
    node = found.holder.find(_X + "autoFilter")
    has_criteria = node is not None and node.find(_X + "filterColumn") is not None
    remove = found.table is None and not keep
    notes: list[str] = []
    if found.table is not None and not keep:
        notes.append("a table keeps its filter buttons: its criteria were cleared")
    editor = workbook.editor
    package = workbook.package
    shown_rows: list[int] = []
    before = package.read(found.holder_part)
    with editor.batch():
        writer = editor.cells.sheet(found.sheet.stored_name)
        if node is not None:
            if remove:
                found.holder.remove(node)
            else:
                for child in list(node):
                    if child.tag == _X + "filterColumn":
                        node.remove(child)
            package.mark_dirty(found.holder_part)
        if found.table is None:
            _filter_mode(workbook, found.sheet, False)
        for row in range(area.first.row + 1, area.last.row + 1):
            if _set_hidden(writer, row, False):
                shown_rows.append(row)
        if shown_rows:
            _touched_rows(editor, found.sheet, shown_rows)
            _recalculate(workbook)
        else:
            _touched_rows(editor, found.sheet, [area.first.row])
        workbook._forget()
    changed = package.read(found.holder_part) != before or bool(shown_rows)
    if not changed:
        notes.append("the filter has no criteria and no row of it is hidden"
                     if not remove else "the sheet has no filter")
    elif not has_criteria and not remove:
        notes.append("the filter had no criteria")
    return FilterChange("clear_autofilter", found.target, shown_range, changed=changed,
                        shown=tuple(_runs(shown_rows)),
                        calculation=workbook.calculation_status, notes=tuple(notes))


# -- the parts a filter touches ----------------------------------------------------------------


def _filter_mode(workbook: Any, sheet: Any, on: bool) -> None:
    """``sheetPr filterMode="1"`` while the sheet's filter has criteria (measured), none
    otherwise; a ``sheetPr`` left with nothing is removed."""
    root = workbook.package.tree(sheet.part)
    node = root.find(_X + "sheetPr")
    if on:
        if node is None:
            node = etree.Element(_X + "sheetPr")
            placed(root, node)
        if node.get("filterMode") != "1":
            node.set("filterMode", "1")
            workbook.package.mark_dirty(sheet.part)
    elif node is not None and "filterMode" in node.attrib:
        del node.attrib["filterMode"]
        if not len(node) and not node.attrib:
            root.remove(node)
        workbook.package.mark_dirty(sheet.part)


def _sheet_index(workbook: Any, sheet: Any) -> int:
    from .sheets import _sheet_nodes

    for index, node in enumerate(_sheet_nodes(workbook)):
        if node.get("name") == sheet.stored_name:
            return index
    return -1


def _filter_name(workbook: Any, sheet: Any, area: Area) -> None:
    """The hidden name ``_xlnm._FilterDatabase`` local to the sheet, naming the range
    anchored, as Excel writes it (made, or set to the range)."""
    root = workbook.package.tree(workbook.part)
    index = str(_sheet_index(workbook, sheet))
    text = (f"{quote_sheet(sheet.stored_name)}!${column_letters(area.first.column)}"
            f"${area.first.row}:${column_letters(area.last.column)}${area.last.row}")
    names = root.find(_X + "definedNames")
    if names is None:
        names = etree.Element(_X + "definedNames")
        later = next((child for child in root if isinstance(child.tag, str)
                      and etree.QName(child).localname in _AFTER_NAMES), None)
        if later is not None:
            later.addprevious(names)
        else:
            root.append(names)
    for node in names:
        if node.get("name", "").casefold() == FILTER_NAME.casefold() \
                and node.get("localSheetId") == index:
            if node.text != text:
                node.text = text
                workbook.package.mark_dirty(workbook.part)
            return
    node = etree.SubElement(names, _X + "definedName")
    node.set("name", FILTER_NAME)
    node.set("localSheetId", index)
    node.set("hidden", "1")
    node.text = text
    workbook.package.mark_dirty(workbook.part)


def _set_hidden(writer: Any, row: int, hidden: bool) -> bool:
    """Set or take away row ``row``'s ``hidden``; whether it changed.  A row element made
    for it holds only that, and one left with nothing is removed (as the cell writer
    does for a row height)."""
    node = writer._row(row, create=hidden)
    if node is None:
        return False
    flag = (node.get("hidden") or "").strip().lower() in ("1", "true")
    if flag == hidden:
        return False
    if hidden:
        node.set("hidden", "1")
    else:
        del node.attrib["hidden"]
        if len(node) == 0 and not set(node.attrib) - {"r", "spans"}:
            node.getparent().remove(node)
            writer._order = None
    writer.workbook.package.mark_dirty(writer.part)
    return True


def _touched_rows(editor: Any, sheet: Any, rows: Iterable[int]) -> None:
    """The step checks these rows of the sheet (no cell of them was written)."""
    edit = editor._edit
    assert edit is not None
    edit.sheets.add(sheet.part)
    edit.rows.setdefault(sheet.part, set()).update(rows)


def _recalculate(workbook: Any) -> None:
    """Ask for a full calculation on open when a formula's value depends on which rows
    are hidden (``SUBTOTAL``, ``AGGREGATE``) anywhere in the workbook, a table's totals
    row's functions among them."""
    if _hidden_aware(workbook):
        workbook.editor.cells.request_full_calculation()


def _hidden_aware(workbook: Any) -> bool:
    package = workbook.package
    dirty = package.dirty_parts
    for sheet in workbook._every_sheet:
        if sheet.part is None:
            continue
        if sheet.part in dirty:
            root = package.tree(sheet.part)
            if root is not None and any(_HIDDEN_AWARE.search((node.text or "").encode())
                                        for node in root.iter(_X + "f")):
                return True
        elif _HIDDEN_AWARE.search(package.read(sheet.part) or b""):
            return True
        for rel in package.relationships(sheet.part).values():
            if rel.type.rpartition("/")[2] == "table" and rel.target_part is not None:
                data = package.read(rel.target_part) or b""
                if _TOTALS.search(data) or _HIDDEN_AWARE.search(data):
                    return True
    return False


def _runs(numbers: Sequence[int]) -> list[str]:
    out: list[str] = []
    start = previous = None
    for number in numbers:
        if previous is not None and number == previous + 1:
            previous = number
            continue
        if start is not None:
            out.append(str(start) if start == previous else f"{start}:{previous}")
        start = previous = number
    if start is not None:
        out.append(str(start) if start == previous else f"{start}:{previous}")
    return out


def _column_runs(numbers: Sequence[int]) -> list[str]:
    return [":".join(column_letters(int(n)) for n in run.split(":")) for run in _runs(numbers)]


# -- rows and columns by hand ----------------------------------------------------------------


def _lines(value: Any, what: str, problems: Problems) -> list[int]:
    """Row numbers or column numbers from ``5``, ``"5:7"``, ``"C"``, ``"C:D"`` or a list."""
    items = value if isinstance(value, (list, tuple)) else [value]
    out: set[int] = set()
    limit = MAX_ROW if what == "rows" else MAX_COLUMN
    for item in items:
        try:
            if isinstance(item, bool):
                raise ValueError
            if isinstance(item, int):
                low = high = item
            elif isinstance(item, str) and item.strip():
                first, _, last = item.strip().replace("$", "").partition(":")
                low, high = _line(first, what), _line(last or first, what)
            else:
                raise ValueError
        except ValueError:
            form = "5, \"5:7\" or a list of them" if what == "rows" else \
                "\"C\", \"C:D\", 3 or a list of them"
            problems.add("bad_position", str(item)[:40], f"{what} are given as {form}")
            continue
        if low > high:
            low, high = high, low
        if not 1 <= low <= high <= limit:
            problems.add("bad_position", str(item)[:40], f"{what} are 1 to {limit:,}")
            continue
        if len(out) + high - low + 1 > 100_000:
            problems.add("too_many_cells", str(item)[:40], f"at most 100,000 {what} in one "
                         "call")
            break
        out.update(range(low, high + 1))
    if not items:
        problems.add("bad_position", what, f"give the {what}")
    return sorted(out)


def _line(text: str, what: str) -> int:
    text = text.strip()
    if what == "rows":
        if not text.isdigit():
            raise ValueError(text)
        return int(text)
    if text.isdigit():
        return int(text)
    if not text.isalpha():
        raise ValueError(text)
    return Cell.parse(text.upper() + "1").column


def set_rows_hidden(workbook: Any, sheet: Any, rows: Any, hidden: bool) -> HiddenChange:
    """Hide or show rows of a sheet by hand (see the module)."""
    action = "hide_rows" if hidden else "show_rows"
    found, problems = _sheet_for(workbook, sheet)
    numbers = _lines(rows, "rows", problems)
    if found is not None and numbers:
        area = Area(Cell(numbers[0], 1), Cell(numbers[-1], MAX_COLUMN))
        guard(workbook, [Target(found, area, "rows", _sheet_label(found))], problems,
              protected="formatRows")
    problems.raise_any()
    assert found is not None
    editor = workbook.editor
    changed: list[int] = []
    with editor.batch():
        writer = editor.cells.sheet(found.stored_name)
        if writer is None:
            raise EditError("not_a_worksheet", addresses=[_sheet_label(found)])
        for number in numbers:
            if _set_hidden(writer, number, hidden):
                changed.append(number)
        if changed:
            _touched_rows(editor, found, changed)
            _recalculate(workbook)
        workbook._forget()
    already = sorted(set(numbers) - set(changed))
    notes = () if changed else (f"the rows are already {'hidden' if hidden else 'shown'}",)
    return HiddenChange(action, _sheet_label(found), tuple(_runs(changed)),
                        tuple(_runs(already)), workbook.calculation_status, notes)


def set_columns_hidden(workbook: Any, sheet: Any, columns: Any, hidden: bool) -> HiddenChange:
    """Hide or show columns of a sheet (see the module)."""
    from .layout import Columns

    action = "hide_columns" if hidden else "show_columns"
    found, problems = _sheet_for(workbook, sheet)
    numbers = _lines(columns, "columns", problems)
    if found is not None and numbers:
        area = Area(Cell(1, numbers[0]), Cell(MAX_ROW, numbers[-1]))
        guard(workbook, [Target(found, area, "columns", _sheet_label(found))], problems,
              protected="formatColumns")
    problems.raise_any()
    assert found is not None
    editor = workbook.editor
    changed: list[int] = []
    with editor.batch():
        writer = editor.cells.sheet(found.stored_name)
        if writer is None:
            raise EditError("not_a_worksheet", addresses=[_sheet_label(found)])
        runs = Columns(workbook, writer)
        for number in numbers:
            now = (runs.get(number).get("hidden") or "").strip().lower() in ("1", "true")
            if now == hidden:
                continue
            runs.set(number, number, hidden="1" if hidden else None, width_default=hidden)
            if not hidden and runs.get(number) == {"width": number_text(runs.default_width())}:
                runs.set(number, number, width=None)     # only the width hiding gave it
            changed.append(number)
        runs.write()
        if changed:
            editor.touched_format(found)
            _recalculate(workbook)
        workbook._forget()
    already = sorted(set(numbers) - set(changed))
    notes = () if changed else (f"the columns are already {'hidden' if hidden else 'shown'}",)
    return HiddenChange(action, _sheet_label(found), tuple(_column_runs(changed)),
                        tuple(_column_runs(already)), workbook.calculation_status, notes)


def _sheet_for(workbook: Any, sheet: Any) -> "tuple[Any, Problems]":
    problems = Problems()
    if workbook.strict:
        raise EditError("strict")
    found = find_sheet(workbook, sheet)
    if found.type != "worksheet" or found.part is None:
        raise EditError("not_a_worksheet", addresses=[found.stored_name])
    return found, problems


__all__ = ["FILTER_NAME", "FilterChange", "HiddenChange", "clear_autofilter",
           "set_autofilter", "set_columns_hidden", "set_rows_hidden"]
