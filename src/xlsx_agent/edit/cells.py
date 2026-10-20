"""Typed cell writes (0.2, E3): ``wb.set_cells({...})``, every refusal before any change.

::

    wb.set_cells({
        "Stock!B15": "Spare fan",                       # text: always literal, "=x" too
        "Stock!D15": 12,                                # a number
        "Stock!G15": True,                              # a boolean
        "Stock!H15": {"date": "2026-10-20"},            # a date (or a Python date)
        "Stock!B4": None,                               # a blank: the value goes, the style stays
        "Stock!J15": {"formula": "=D15*E15"},           # a formula only when given as one
    })

**Values.**  Text is always text: ``"=D15"`` is stored as those four characters, and
nothing that looks like a number or a date is read as one.  Only ``{"formula": "=..."}``
(or an ooxml-edit :class:`~ooxml_edit.sheets.Formula`) is a formula; it is written with no
cached value, so it reads ``absent`` until the application has calculated it.  A number is
an ``int`` or a ``float``, finite, and a whole number that a double holds exactly (beyond
2^53 one is refused: write it as text, as an id).  ``None`` is a blank: the value (and any
formula) goes, the cell's format stays; empty text is refused for the same reason (it is no
value).  ``{"value": v}`` is ``v``.  Nothing is coerced: a blank is never 0, and text over
32,767 characters is refused, never cut.

**Dates.**  ``{"date": "2026-10-20"}`` or a Python ``date`` is the day's serial in the
workbook's date system (1900 or 1904, ``workbookPr date1904``).  A date string is ISO 8601
(``YYYY-MM-DD``) or written in the date order of the workbook's locale
(:attr:`Workbook.locale <xlsx_agent.Workbook.locale>`, the shared
:class:`ooxml_common.locales.Locale`: ``20/10/2026`` under ``dmy``) with a four-digit year;
a two-digit year, a locale without a date order, a time in the string, or a day no
calendar has is refused.  A ``datetime`` with a time zone is refused (a cell has none: a
time zone is never dropped silently); a naive one is taken as its date when its time is
midnight, and refused otherwise: times and durations come later.  A ``time`` and a
``timedelta`` are refused.  The cell must show a date: its number format (its own, or for a
new cell the one its row or column gives it) is one whose code shows a day
(:meth:`ooxml_edit.sheets.Styles.is_date`, and xlsx-agent's reading of the
language-dependent built-in ids in the workbook's locale, :mod:`xlsx_agent.read.styles`);
a format that shows only a time, or none that shows a day, is refused naming the cell,
unless the call gives one: ``{"date": "2026-10-20", "format": "yyyy-mm-dd"}`` (a code that
shows a day), which the cell is given, reusing an identical format.

**Guards** (each a refusal of the whole call, :class:`~xlsx_agent.edit.EditError` with the
cells concerned): writing over a cell that holds a formula, unless the call says
``replace_formulas=True``; a cell of an array formula's block or a data table, and a cell
inside a merged area other than its first; an Excel table's header cell (its text is the
column's name); a cell named twice.  A neighbour's formula is never touched.

**Shared formulas.**  Writing over the master of a shared formula group (the cell that
holds the group's text) hands the group to its next member, its formula moved there by
:func:`xlsx_agent.formula.shift` (:meth:`ooxml_edit.sheets.Worksheet.release_shared`): every
other member keeps its formula.  A member written leaves its group alone.  A master whose
formula the lexer cannot move safely is refused (``shared_formula``), never guessed.

**Atomic.**  Every value is checked, and every guard asked, before anything changes; then
the writes run in one step (:meth:`Workbook.batch <xlsx_agent.Workbook.batch>`, or the
open batch), so a write that fails anyway rolls back every one.  The first write drops the
calculation chain and sets ``fullCalcOnLoad``: every formula's cached value in the workbook
reads ``invalidated`` from then on (:mod:`xlsx_agent.edit`).
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping

from ooxml_common.locales import DATE_ORDERS
from ooxml_edit.dates import is_date_code
from ooxml_edit.sheets import MAX_TEXT, Cell, Formula

from ..formula import ShiftError, shift, tokens
from ..read.styles import classify_format
from ..read.types import _letters

#: A whole number beyond this is not held exactly by a double (Excel's numbers).
MAX_EXACT = 2 ** 53

#: The most refusals one error lists.
MAX_LISTED = 50

_ISO_DAY = re.compile(r"([0-9]{4})-([0-9]{2})-([0-9]{2})")
_LOCAL_DAY = re.compile(r"([0-9]{1,4})([/.-])([0-9]{1,2})\2([0-9]{1,4})")


@dataclass
class Planned:
    """One cell's write, checked: what the writer is given."""

    address: str
    sheet: Any
    cell: Cell
    value: Any
    kind: str
    old: str
    #: The format the cell is given: an index, or ``None`` to keep its own.
    style: int | None = None
    #: A date's format code to give the cell (resolved inside the step).
    format: str | None = None
    #: The cell is the master of a shared formula group, handed on before it is written.
    release: bool = False


@dataclass(frozen=True)
class Written:
    """What :meth:`Workbook.set_cells <xlsx_agent.Workbook.set_cells>` did: the cells
    written (``cells``, addresses, in the order given; ``changed`` how many of them changed),
    each one's type before (``old``: ``text``, ``number``, ``bool``, ``error``, ``blank``
    or ``formula``), the shared formula groups handed on (``released``: the old masters'
    addresses), and the workbook's calculation status after (``invalidated`` once anything
    was written).  Never a value."""

    cells: tuple[str, ...]
    changed: int
    old: dict[str, str] = field(default_factory=dict)
    released: tuple[str, ...] = ()
    calculation: str = "cached_unverified"

    def as_dict(self, limit: int = MAX_LISTED) -> dict[str, Any]:
        """As plain data, at most ``limit`` cells listed (the rest counted)."""
        shown = self.cells[:limit]
        out: dict[str, Any] = {"written": len(self.cells), "changed": self.changed,
                               "cells": [{"a": a, "old": self.old.get(a)} for a in shown]}
        if len(self.cells) > limit:
            out["cells_cut"] = len(self.cells) - limit
        if self.released:
            out["released"] = list(self.released[:limit])
        out["calculation"] = self.calculation
        return out


def set_cells(workbook: Any, cells: Any, *, replace_formulas: bool = False) -> Written:
    """Write ``cells`` (see the module): a mapping of address to value, or ``(address,
    value)`` pairs.  :class:`~xlsx_agent.edit.EditError` before anything changes for any
    value refused or guard met, every refused cell listed (at most :data:`MAX_LISTED`)."""
    from . import EditError

    items = list(cells.items()) if isinstance(cells, Mapping) else list(cells)
    if not items:
        raise EditError("bad_value", "no cells to write")
    editor = workbook.editor
    changed = 0
    released: list[str] = []
    with editor.batch():
        # Checked inside the step: a refusal caught in an open batch still rolls it back,
        # as any write that fails there does (xlsx_agent.edit, "Atomic").
        planned = plan(workbook, items, replace_formulas=replace_formulas)
        cells_of: dict[str, set[Cell]] = {}
        for write in planned:
            cells_of.setdefault(write.sheet.part, set()).add(write.cell)
        for write in planned:
            writer = editor.cells.sheet(write.sheet.stored_name)
            content = writer.content(write.cell)
            if content.formula_type == "shared" and content.formula:
                # The master of a group: handed to a member this call does not write.
                moved = writer.release_shared(write.cell, _translate,
                                              written=cells_of[write.sheet.part])
                if moved:
                    released.append(write.address)
                    editor.touched(write.sheet, moved)
            style = write.style
            if write.format is not None:
                styles = editor.cells.styles
                base = writer.style(write.cell, inherited=True)
                style = styles.with_number_format(base, write.format)
            if editor.write(write.sheet, write.cell, write.value, style=style):
                changed += 1
    return Written(cells=tuple(w.address for w in planned), changed=changed,
                   old={w.address: w.old for w in planned}, released=tuple(released),
                   calculation=workbook.calculation_status)


def _translate(text: str, rows: int, columns: int) -> str:
    return shift(text, rows, columns)


# -- checking ----------------------------------------------------------------------------------


class _Refusals:
    """The refusals met, by code, in the order met."""

    def __init__(self) -> None:
        self.found: list[tuple[str, str, str | None]] = []

    def add(self, code: str, address: str, detail: str | None = None) -> None:
        self.found.append((code, address, detail))

    def raise_any(self) -> None:
        from . import EditError

        if not self.found:
            return
        code, _, detail = self.found[0]
        same = [address for c, address, _ in self.found if c == code]
        others = len(self.found) - len(same)
        if others:
            detail = (f"{detail}; " if detail else "") + f"{others} other refusal(s)"
        raise EditError(code, detail, addresses=same,
                        problems=[(c, a) for c, a, _ in self.found[:MAX_LISTED]])


def plan(workbook: Any, items: "list[tuple[Any, Any]]", *,
         replace_formulas: bool = False) -> list[Planned]:
    """Every write checked against the values allowed and the guards (see the module),
    nothing changed: the writes, or :class:`~xlsx_agent.edit.EditError`."""
    from ..address import AddressError, NameNotFound, RangeAddress, parse_address
    from . import EditError
    from .structure import _label

    if workbook.strict:
        raise EditError("strict")
    refusals = _Refusals()
    out: list[Planned] = []
    seen: dict[tuple[str, int, int], str] = {}
    sheets: dict[str, _SheetFacts] = {}
    editor = workbook.editor
    for item in items:
        if not isinstance(item, (tuple, list)) or len(item) != 2:
            raise EditError("bad_address", "cells are a mapping of address to value, or "
                            "(address, value) pairs")
        address, given = item
        if not isinstance(address, str):
            raise EditError("bad_address", f"an address is text, not {type(address).__name__}")
        try:
            parsed = parse_address(address)
        except (AddressError, ValueError):
            refusals.add("bad_address", address.strip()[:80])
            continue
        if not isinstance(parsed, RangeAddress) or parsed.kind != "cell" \
                or parsed.sheet is None:
            refusals.add("bad_address", address.strip()[:80],
                         "a cell with its sheet, as Stock!D10")
            continue
        try:
            sheet = workbook.sheet_of(parsed)
        except (NameNotFound, KeyError, ValueError):
            refusals.add("bad_address", address.strip()[:80], "no such sheet")
            continue
        cell = parsed.area.first
        label = f"{_label(sheet)}!{_letters(cell.column)}{cell.row}"
        if sheet.type != "worksheet" or sheet.part is None:
            refusals.add("not_a_worksheet", label)
            continue
        key = (sheet.part, cell.row, cell.column)
        if key in seen:
            refusals.add("duplicate_cell", label)
            continue
        seen[key] = label
        facts = sheets.get(sheet.part)
        if facts is None:
            writer = editor.cells.sheet(sheet.stored_name)
            if writer is None:
                refusals.add("not_a_worksheet", label)
                continue
            facts = sheets[sheet.part] = _SheetFacts(workbook, sheet, writer)
        write = _value(workbook, facts, label, sheet, cell, given, refusals)
        if write is None:
            continue
        if write.style is None and write.format is None and write.kind != "blank":
            # A new cell takes its row's or column's format, as one typed there does.
            inherited = facts.writer.style(cell, inherited=True)
            if inherited != facts.writer.style(cell):
                write.style = inherited
        if _guards(facts, write, refusals, replace_formulas):
            out.append(write)
    for part, facts in sheets.items():
        if facts.masters:
            _check_releases(facts, {w.cell for w in out if w.sheet.part == part}, refusals)
    refusals.raise_any()
    return out


class _SheetFacts:
    """What the guards ask of one sheet, read once per call: its merges and array blocks
    (the sheet's scan), its tables' header rows."""

    def __init__(self, workbook: Any, sheet: Any, writer: Any) -> None:
        self.workbook = workbook
        self.writer = writer
        found = workbook.scan(sheet)
        self.merges = list(found.merges)
        self.arrays = list(found.arrays)
        #: The masters of shared formula groups the call writes, with their groups.
        self.masters: list = []
        self.headers = []
        for table in writer.tables():
            area = table.area
            if area is not None and table.header_rows:
                self.headers.append((area.first.row, area.first.row + table.header_rows - 1,
                                     area.first.column, area.last.column))


def _value(workbook: Any, facts: _SheetFacts, label: str, sheet: Any, cell: Cell,
           given: Any, refusals: _Refusals) -> Planned | None:
    """The value to write, checked; ``None`` (with the refusal noted) when it is refused."""
    content = facts.writer.content(cell)
    old = "formula" if content.formula is not None else content.kind
    date_format = None
    if isinstance(given, Mapping):
        keys = set(given)
        kinds = keys & {"value", "formula", "date"}
        extra = keys - {"value", "formula", "date", "format"}
        if len(kinds) != 1 or extra or ("format" in keys and kinds != {"date"}):
            refusals.add("bad_value", label, 'one of "value", "formula" or "date" (a date '
                         'with an optional "format")')
            return None
        (kind,) = kinds
        if kind == "value":
            given = given["value"]
            if isinstance(given, Mapping):
                refusals.add("bad_value", label, "a value is text, a number, a boolean or "
                             "None")
                return None
        elif kind == "formula":
            text = given["formula"]
            problem = _formula_problem(text)
            if problem is not None:
                refusals.add("bad_formula", label, problem)
                return None
            return Planned(label, sheet, cell, Formula(text), "formula", old)
        else:
            date_format = given.get("format")
            if date_format is not None:
                problem = _format_problem(date_format)
                if problem is not None:
                    refusals.add("date_format", label, problem)
                    return None
            day = _day(given["date"], workbook, label, refusals)
            if day is None:
                return None
            return _dated(workbook, facts, label, sheet, cell, day, old, date_format,
                          refusals)
    if isinstance(given, Formula):
        return Planned(label, sheet, cell, given, "formula", old)
    if given is None:
        return Planned(label, sheet, cell, None, "blank", old)
    if isinstance(given, bool):
        return Planned(label, sheet, cell, given, "bool", old)
    if isinstance(given, (int, float)):
        if isinstance(given, float) and not math.isfinite(given):
            refusals.add("bad_number", label, "not a finite number")
            return None
        if isinstance(given, int) and abs(given) > MAX_EXACT:
            refusals.add("bad_number", label, "a whole number past 2^53 is not held exactly: "
                         "write it as text")
            return None
        return Planned(label, sheet, cell, given, "number", old)
    if isinstance(given, str):
        if len(given) > MAX_TEXT:
            refusals.add("text_too_long", label, f"{len(given):,} characters")
            return None
        if not given:
            refusals.add("bad_value", label, "empty text is no value: write None for a blank")
            return None
        return Planned(label, sheet, cell, given, "text", old)
    if isinstance(given, (_dt.date, _dt.datetime)):
        day = _day(given, workbook, label, refusals)
        if day is None:
            return None
        return _dated(workbook, facts, label, sheet, cell, day, old, None, refusals)
    if isinstance(given, (_dt.time, _dt.timedelta)):
        refusals.add("date_time", label, "times and durations are not written yet")
        return None
    refusals.add("bad_value", label, f"not a value a cell holds: {type(given).__name__}")
    return None


def _day(given: Any, workbook: Any, label: str, refusals: _Refusals) -> _dt.date | None:
    """A date given as a Python date or datetime, or as text (see the module)."""
    if isinstance(given, _dt.datetime):
        if given.tzinfo is not None:
            refusals.add("date_time", label, "a time zone a cell cannot hold: give the date "
                         "in the zone the workbook means, without tzinfo")
            return None
        if given.time() != _dt.time(0):
            refusals.add("date_time", label, "a time of day: times are not written yet; give "
                         "the date alone")
            return None
        return given.date()
    if isinstance(given, _dt.date):
        return given
    if not isinstance(given, str):
        refusals.add("bad_date", label, "a date is text (YYYY-MM-DD) or a Python date")
        return None
    text = given.strip()
    match = _ISO_DAY.fullmatch(text)
    if match is not None:
        year, month, day = (int(g) for g in match.groups())
    else:
        local = _LOCAL_DAY.fullmatch(text)
        order = getattr(workbook.locale, "date_order", None)
        if local is None or order not in DATE_ORDERS:
            refusals.add("bad_date", label, "a date as YYYY-MM-DD, or in the locale's date "
                         "order with a four-digit year")
            return None
        parts = dict(zip(order, (local.group(1), local.group(3), local.group(4))))
        if len(parts["y"]) != 4 or len(parts["m"]) > 2 or len(parts["d"]) > 2:
            refusals.add("bad_date", label, "ambiguous: a four-digit year in the locale's "
                         f"date order ({order}), or YYYY-MM-DD")
            return None
        year, month, day = int(parts["y"]), int(parts["m"]), int(parts["d"])
    try:
        return _dt.date(year, month, day)
    except ValueError:
        refusals.add("bad_date", label, "no calendar has that day")
        return None


def _dated(workbook: Any, facts: _SheetFacts, label: str, sheet: Any, cell: Cell,
           day: _dt.date, old: str, date_format: str | None,
           refusals: _Refusals) -> Planned | None:
    first = _dt.date(1904, 1, 1) if workbook.date1904 else _dt.date(1900, 1, 1)
    if day < first:
        refusals.add("bad_date", label, f"before {first.isoformat()}, the first day of the "
                     "workbook's date system: write it as text")
        return None
    style = facts.writer.style(cell, inherited=True)
    own = facts.writer.style(cell)
    if date_format is not None:
        if workbook.editor.cells.styles is None:
            refusals.add("date_format", label, "the workbook has no cell formats to give it "
                         "one: write the date as text")
            return None
        return Planned(label, sheet, cell, day, "date", old, format=date_format)
    shows = _shows_day(workbook, style)
    if shows is not True:
        refusals.add("date_format", label, shows)
        return None
    return Planned(label, sheet, cell, day, "date", old,
                   style=style if style != own else None)


_GIVE = '{"date": ..., "format": "yyyy-mm-dd"}'


def _shows_day(workbook: Any, style: int) -> "bool | str":
    """``True`` when cell format ``style`` shows a day; else what to do (the refusal's
    detail)."""
    styles = workbook.editor.cells.styles
    if styles is None:
        return f"the workbook has no cell formats; give one: {_GIVE}"
    if styles.is_date(style) is False:
        return f"give one: {_GIVE}"
    kind = workbook.styles.format(style).kind
    if kind in ("date", "datetime"):
        return True
    if kind == "time":
        return f"it shows only a time; give a date format: {_GIVE}"
    return f"it is not known to show one in the workbook's language; give one: {_GIVE}"


def _format_problem(code: Any) -> str | None:
    if not isinstance(code, str) or not code.strip():
        return "a format is a number format code, as yyyy-mm-dd"
    if not is_date_code(code) or classify_format(code)[0] not in ("date", "datetime"):
        return "the format given does not show a day"
    return None


def _formula_problem(text: Any) -> str | None:
    if not isinstance(text, str):
        return "a formula is text, as =D10*2"
    body = text[1:] if text.startswith("=") else text
    if not body.strip():
        return "a formula needs text; write None for a blank"
    if len(body) > MAX_TEXT:
        return "longer than a cell holds"
    try:
        found = tokens(body)
    except ShiftError as error:
        return str(error)
    depth = 0
    for token in found:
        if token.kind == "other":
            depth += token.text.count("(") - token.text.count(")")
            if depth < 0:
                break
    if depth != 0:
        return "its parentheses do not pair"
    return None


def _guards(facts: _SheetFacts, write: Planned, refusals: _Refusals,
            replace_formulas: bool) -> bool:
    """Whether the write passes the guards (see the module); a refusal noted otherwise."""
    cell, label = write.cell, write.address
    content = facts.writer.content(cell)
    if content.formula_type in ("array", "dataTable") or any(
            area.contains(cell) for area in facts.arrays):
        refusals.add("array_formula", label)
        return False
    for area in facts.merges:
        if area.contains(cell) and cell != area.first:
            refusals.add("merged_cell", label, f"write {_letters(area.first.column)}"
                         f"{area.first.row}, the merged area's first cell")
            return False
    for top, bottom, left, right in facts.headers:
        if top <= cell.row <= bottom and left <= cell.column <= right:
            refusals.add("table_header", label)
            return False
    if content.formula is not None:
        if not replace_formulas:
            refusals.add("formula_cell", label)
            return False
        if content.formula_type == "shared" and content.formula:
            group = facts.writer.shared_group(cell)
            if group is not None and group.master == cell:
                write.release = True
                facts.masters.append((write, group))
    return True


def _check_releases(facts: _SheetFacts, written: set[Cell], refusals: _Refusals) -> None:
    """Every master written: its formula moved to each member the call leaves (what the
    group is handed on with), checked before anything changes."""
    for write, group in facts.masters:
        try:
            for member in group.members:
                if member not in written:
                    shift(group.text, member.row - write.cell.row,
                          member.column - write.cell.column)
        except ShiftError as error:
            refusals.add("shared_formula", write.address, str(error))
