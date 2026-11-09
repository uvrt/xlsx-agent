"""Editing a workbook (0.2, E1): one history, atomic steps, bounded reports, a checked save.

Every change to a workbook goes through its :class:`Editor` (``wb.editor``, made when first
needed): ooxml-edit's :class:`~ooxml_edit.history.History` over the workbook's package, and
ooxml-edit's lossless cell writer (:mod:`ooxml_edit.sheets`) on the same package::

    with wb.batch() as edit:          # one step: all of it, or none of it
        wb.set_cells({...})           # the writes (E3: xlsx_agent.edit.cells)
    edit.report                       # EditReport: cells by range, parts, calculation, issues
    wb.undo(); wb.to_bytes() == original      # True: undo to the start is the bytes opened
    wb.redo(); wb.to_bytes() == edited        # True
    wb.save("stock.xlsx")             # reopened and checked before it is written

**Atomic.**  A batch is one undo step and one version.  Anything that raises inside it
rolls the whole package back to where the batch began -- cells, shared strings and styles,
``calcPr``, the calculation chain with its relationship and content-type entry -- and the
version stays as it was, so nothing derived from the workbook (a tool session's datasets,
keyed by version) goes stale.  The calculation status is read from the package
(:attr:`Workbook.calculation_status <xlsx_agent.Workbook.calculation_status>`), so it is
rolled back with it.  An error caught inside the batch does not make the rest of it stand:
the write that failed marks the batch, and the batch rolls back on leaving and raises
:class:`EditError` ``rolled_back``.  Only :meth:`Editor.savepoint` (what the tool layer's
``batch`` runs each op in) rolls back a part of a batch and lets the rest stand.

**Writing invalidates.**  The first write that changes a cell drops the calculation chain
and sets ``fullCalcOnLoad``: every formula's cached value in the workbook may now be stale
(workbook-wide and conservative: there is no dependency graph).  Each written cell is read
back through the writer at once; one that does not hold what was given refuses the batch
(``readback``).  The shared strings' counts are brought up to date as each step closes
(:meth:`ooxml_edit.sheets.Workbook.finish`).

**Checked as it closes, for what it touched.**  Every step's report carries the structure
check (:mod:`.structure`) of the sheets it wrote: on the rows whose cells it wrote
(:attr:`Edit.rows`), or whole for a sheet it changed otherwise (:attr:`Edit.whole`:
formatting, layout, rows and columns).  Its guards read a sheet's merges, array formulas
and validation rules, and a block's cells, from the tree the cell writer works on, never
typing the sheet's values (round 44, lane P): a step on a sheet of 50,000 rows costs what
it touches.  The parts it changed are the history's own states compared.

**History.**  :meth:`Editor.undo` and :meth:`Editor.redo` move one step; equal versions mean
equal content, and undoing every step gives the bytes that were opened.  In a tool session a
step can also be taken back by sheet (``undo``'s ``scope``: a sheet's name) or by caller,
while no later step changed a part it changed: a step's sheet is its own, and the workbook
part, shared strings, styles, calculation chain, the workbook's relationships and the
content types are every sheet's (:func:`undo_scope`, :func:`undo_shared`), so a step that
changed one of them is taken back only by a plain undo, which never leaves a later change
without what it built on.

**Saving.**  :func:`serialise` writes the package as the kind asked for (``xlsx``,
``xlsm``, ``xltx``, ``xltm``: the main part's content type set to it, nothing else), and
refuses a kind without macros while the workbook holds a VBA project (``vba_project``).
:func:`verify` reopens those bytes as a new workbook and reads every cell written since the
workbook was opened back through xlsx-agent's own reader: its type and value must be what
the writer holds (``readback``, with the addresses), and the structure check
(:mod:`.structure`) is run on what was saved.

**Guarded** (X4).  The host's write policy (:mod:`.policy`), cells by key and new rows in
a declared table's capacity (:mod:`.keys`), expectations of what a cell holds, a revision
guard and protected sheets (:mod:`.cells`) are checked before anything changes;
:meth:`Editor.write` asks the policy and a sheet's protection again of every write,
whatever planned it.

**No content.**  Reports, errors and log records (the ``xlsx_agent.edit`` logger) carry
versions, counts, codes, part names and addresses: never a value, a text or a formula.
"""

from __future__ import annotations

import datetime as _dt
import logging
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any, Iterable, Iterator

from ooxml_common.kinds import KINDS, REL_VBA_PROJECT, main_content_type
from ooxml_edit.history import History
from ooxml_edit.sheets import Cell, Formula, RichText, date_serial
from ooxml_edit.sheets import Workbook as CellWriter

from ..address import format_cell
from ..read.strings import decode_xstring
from ..read.types import Problem
from .report import MAX_ISSUES, EditReport, capped, compress
from .structure import CODES as STRUCTURE_CODES
from .structure import _label, check_structure

log = logging.getLogger("xlsx_agent.edit")

#: Every :attr:`EditError.code`.
EDIT_CODES = {
    "rolled_back": "a write in the batch failed; the whole batch was rolled back",
    "readback": "a cell did not read back as written",
    "not_a_worksheet": "the sheet holds no cells to write",
    "strict": "a strict (purl.oclc.org) workbook is not edited until Excel's behaviour on "
              "one is measured",
    "in_batch": "undo and redo wait until the batch ends",
    "vba_project": "the workbook holds a VBA project, which this kind cannot carry: save "
                   "it as .xlsm (or .xltm)",
    "bad_kind": "not a workbook kind",
    # set_cells (E3, xlsx_agent.edit.cells): each refused before anything changes
    "bad_address": "not a cell address with its sheet, as Stock!D10",
    "duplicate_cell": "a cell is named twice in one call",
    "bad_value": "not a value a cell is written with",
    "bad_number": "not a number a cell holds exactly",
    "text_too_long": "text longer than a cell holds (32,767 characters); it is never cut",
    "bad_formula": "not a formula that can be written",
    "bad_date": "not a date that can be written",
    "date_time": "a time, a duration or a time zone, which are not written yet",
    "date_format": "the cell's number format does not show a date",
    "formula_cell": "the cell holds a formula: pass replace_formulas to write over it",
    "array_formula": "the cell is part of an array formula or a data table, written as a "
                     "whole in Excel",
    "merged_cell": "the cell is inside a merged area and not its first cell",
    "table_header": "the cell is a table's header, whose text is its column's name",
    "shared_formula": "the cell holds a shared formula whose references cannot be moved "
                      "safely to the group's other cells",
    # guards (X4, xlsx_agent.edit.policy, .keys, .cells): each refused before anything changes
    "protected": "the sheet is protected: nothing on it is written",
    "not_writable": "the host's write policy does not let this cell be written",
    "formulas_not_allowed": "the host's write policy writes no formula here",
    "bad_policy": "the host's write policy names what the workbook does not have",
    "not_expected": "the cell does not hold what the call expected",
    "revision": "the workbook is not at the version the call expected",
    "positional": "a table's row is written by its key, never by its place",
    "no_key_column": "the table has no key column declared to write by key",
    "missing_key": "no row has that key",
    "duplicate_key": "two rows share that key",
    "no_capacity": "the table has no capacity for new rows",
    "capacity_full": "the table's capacity for new rows is used up",
    "capacity_row": "the cell is in a table's free capacity, where a row is started by key",
    # the contract's keys and rows (X6, xlsx_agent.edit.keys)
    "formula_key": "the table's key column holds formulas: keys are values the application "
                   "supplies",
    "not_capacity": "the row allocated is outside the table's capacity",
    "data_row": "the row allocated is already one of the table's data rows",
    "not_free": "the row allocated is not empty",
    "allocated_twice": "a row or a key is allocated twice",
    "existing_key": "a row already has the key allocated a new row",
    "capacity_gap": "the rows allocated leave a blank row below the table's data rows",
    "not_allocated": "the write policy takes new rows only where allocate names them",
    # formatting and layout (E7, round 44: xlsx_agent.edit.format, .layout, .fit)
    "bad_format": "not a formatting option or value that can be written",
    "bad_target": "not cells, columns, rows or a sheet that can be formatted",
    "no_cell_style": "the workbook has no cell style of that name",
    "too_many_cells": "more cells than one formatting call takes",
    "unmeasured_font": "a cell's font has no measured widths to fit a column to",
    "wrong_password": "not the sheet's password",
    # a declared field derived from another (round 26, xlsx_agent.declare's from_field)
    "derived_column": "the column is read from another column's cells by the table's own "
                      "rules: write that column",
    # inserting, deleting and moving rows and columns (round 44, xlsx_agent.edit.grid)
    "bad_position": "not a row or column (or run of them) on the sheet",
    "array_split": "the change would cut through an array formula, which Excel changes only "
                   "as a whole",
    "off_grid": "cells that hold something would be pushed off the sheet",
    "table_header_row": "a table's header row would be deleted, which Excel refuses",
    "table_move": "columns of a table, or rows out of or across a table's edge, would be "
                  "moved, which Excel refuses",
    "table_rows": "rows of two tables side by side would be inserted, deleted or moved at "
                  "once, which Excel refuses",
    "merge_split": "part of a merged area would be moved, which Excel refuses",
    "pivot_table": "the change would cut through a pivot table",
    "not_measured": "a dependent whose change Excel makes is not measured yet: nothing is "
                    "written rather than half of it",
    # sheets (E5, R52: xlsx_agent.edit.sheets)
    "no_sheet": "the workbook has no such sheet",
    "bad_sheet_name": "not a name Excel gives a sheet",
    "sheet_exists": "another sheet has that name, as Excel compares names",
    "bad_index": "not a place among the sheets",
    "bad_state": "not a sheet's state",
    "structure_locked": "the workbook's structure is protected",
    "referred": "something refers to what the change would remove",
    "last_sheet": "the workbook's last sheet",
    "last_visible": "the workbook's last visible sheet",
    # merged cells (R56: xlsx_agent.edit.merges)
    "merge_range": "not a range of more than one cell on a sheet",
    "merge_overlap": "the range overlaps a merged area",
    "merge_table": "the range overlaps a table, whose cells Excel does not merge",
    "merge_array": "the range overlaps an array formula",
    "merge_content": "cells other than the range's first hold content, which a merge drops",
    # notes (R57: xlsx_agent.edit.notes) and document properties (R62: .properties)
    "threaded_comment": "the cell holds a threaded comment, which is not written or "
                        "removed here",
    "bad_property": "not a document property, or not a value it takes",
    # a table written, rows appended (round 44, xlsx_agent.edit.tables)
    "not_empty": "the cell holds a value or a formula, which a block is not written over "
                 "unless asked",
    "table_overlap": "a table or an AutoFilter is there, which a block is never written over",
    "bad_table_name": "not a name Excel gives a table",
    "bad_table_style": "not a table style the workbook has",
    "totals_referenced": "something refers to the table's totals row by address, and the "
                         "row would move down from under it",
}

#: The content type of a VBA project part ([MS-OFFMACRO2]; what Excel writes).
CT_VBA_PROJECT = "application/vnd.ms-office.vbaProject"


class EditError(ValueError):
    """An edit or a save refused.  :attr:`code` is one of :data:`EDIT_CODES`;
    :attr:`addresses` the cells concerned (at most 20), never their content;
    :attr:`problems` every refusal of a call that met several, as ``(code, address)`` (at
    most 50; the first code's are :attr:`addresses`); :attr:`held` what each cell an
    expectation was not met at holds, as its type (``not_expected``; at most 50), never its
    value."""

    def __init__(self, code: str, detail: str | None = None, *,
                 addresses: "tuple[str, ...] | list[str]" = (),
                 problems: "tuple[tuple[str, str], ...] | list[tuple[str, str]]" = (),
                 held: "dict[str, str] | None" = None) -> None:
        if code not in EDIT_CODES:
            raise ValueError(f"not an edit error code: {code!r}")
        self.code = code
        self.addresses = tuple(addresses)[:20]
        self.problems = tuple(problems)[:50]
        self.held = dict(list((held or {}).items())[:50])
        message = EDIT_CODES[code]
        if detail:
            message = f"{message}: {detail}"
        if self.addresses:
            message += f" ({', '.join(self.addresses)})"
        super().__init__(message)


@dataclass
class Edit:
    """What :meth:`Editor.batch` yields: :attr:`report` once the batch has closed."""

    report: EditReport | None = None
    #: The cells written, by sheet label: (row, column).
    cells: dict[str, set[tuple[int, int]]] = field(default_factory=dict)
    #: The sheet parts written.
    sheets: set[str] = field(default_factory=set)
    #: The rows of each sheet part whose cells were written (:meth:`Editor.write`,
    #: :meth:`Editor.touched`): the cells the step's check reads.
    rows: dict[str, set[int]] = field(default_factory=dict)
    #: The sheet parts changed otherwise (formatting, layout, rows or columns inserted,
    #: deleted or moved): checked whole.
    whole: set[str] = field(default_factory=set)
    #: Why the batch must roll back: an error raised inside it, caught there.
    failed: str | None = None


@dataclass(frozen=True)
class SaveReport:
    """What a checked save found: the kind and size saved, how many written cells were
    read back, and the structure issues of what was saved (codes with addresses)."""

    kind: str
    size: int
    verified: int
    issues: tuple[Problem, ...] = ()

    def as_dict(self) -> dict[str, Any]:
        return {"kind": self.kind, "size": self.size, "verified": self.verified,
                "issues": [{"code": p.code, "address": p.address, "count": p.count}
                           for p in self.issues[:MAX_ISSUES]]}


class Editor:
    """A workbook's history and cell writer (see the module).  Made by
    :attr:`Workbook.editor <xlsx_agent.Workbook.editor>`."""

    def __init__(self, workbook: Any) -> None:
        self.workbook = workbook
        self.package = workbook.package
        self.history = History(self.package)
        self._cells: CellWriter | None = None
        self._depth = 0
        self._savepoints = 0
        self._edit: Edit | None = None
        #: Every cell written since the workbook was opened, by sheet part: what a save
        #: reads back (a superset after an undo, which is harmless: the writer is compared).
        self._written: dict[str, set[tuple[int, int]]] = {}
        #: The last step's report.
        self.last: EditReport | None = None

    @property
    def cells(self) -> CellWriter:
        """ooxml-edit's writer over the workbook's package."""
        if self._cells is None:
            self._cells = CellWriter(self.package)
        return self._cells

    @property
    def version(self) -> int:
        return self.history.version

    @property
    def in_batch(self) -> bool:
        return self._depth > 0

    # -- steps -------------------------------------------------------------------------------

    @contextmanager
    def batch(self) -> Iterator[Edit]:
        """One step: everything inside is kept together, or rolled back together."""
        if self._depth:
            assert self._edit is not None   # the open batch's
            self._depth += 1
            try:
                yield self._edit
            except Exception as error:
                if not self._savepoints and self._edit is not None:
                    self._edit.failed = type(error).__name__
                raise
            finally:
                self._depth -= 1
            return
        edit = Edit()
        before = self.history.version
        self._edit, self._depth = edit, 1
        try:
            # The history's own states before and after the step say which parts it
            # changed: no state of the package is taken twice.
            with self.history.watch() as change, self.history.batch():
                yield edit
                if edit.failed is not None:
                    raise EditError("rolled_back", f"a {edit.failed} was caught inside it")
                if self._cells is not None:
                    self._cells.finish()
        except Exception as error:
            self.workbook._forget()
            log.info("edit rolled back; version %d kept (%s)", before, type(error).__name__)
            raise
        finally:
            self._edit, self._depth = None, 0
        self.workbook._forget()
        edit.report = self.last = self._report(edit, before, change)
        if edit.report.changed:
            log.info("edit %d -> %d: %d cell(s), %d part(s), %d issue(s)", before,
                     edit.report.version, edit.report.cell_count, len(edit.report.parts),
                     len(edit.report.issues))
        return

    @contextmanager
    def savepoint(self) -> Iterator[None]:
        """Inside a batch: if what runs inside raises, it is rolled back alone, and the
        batch stays open with what came before (the tool layer's ``batch`` runs each op in
        one).  Outside a batch it is a :meth:`batch`."""
        if not self._depth:
            with self.batch():
                yield
            return
        snapshot = self.package.snapshot()
        edit = self._edit
        assert edit is not None   # the open batch's
        cells = {sheet: set(found) for sheet, found in edit.cells.items()}
        rows = {part: set(found) for part, found in edit.rows.items()}
        sheets, whole = set(edit.sheets), set(edit.whole)
        self._savepoints += 1
        try:
            yield
        except Exception:
            self.package.restore(snapshot)
            edit.cells, edit.sheets, edit.rows, edit.whole = cells, sheets, rows, whole
            self.workbook._forget()
            raise
        finally:
            self._savepoints -= 1

    def undo(self) -> bool:
        """Back one step; whether there was one."""
        return self._move("undo")

    def redo(self) -> bool:
        """Forward one undone step; whether there was one."""
        return self._move("redo")

    def _move(self, direction: str) -> bool:
        if self._depth:
            raise EditError("in_batch")
        before = self.history.version
        done = getattr(self.history, direction)()
        if done:
            self.workbook._forget()
            log.info("%s %d -> %d", direction, before, self.history.version)
        return done

    # -- writing -----------------------------------------------------------------------------

    def write(self, sheet: Any, cell: Cell, value: Any, *, style: int | None = None) -> bool:
        """Write one cell of ``sheet`` (an :class:`xlsx_agent.Sheet`) in a step of its own,
        or in the open batch; whether it changed.  ``value`` as
        :meth:`ooxml_edit.sheets.Worksheet.write` takes it.  The first change invalidates
        every formula's cached value (see the module); the cell is read back at once."""
        label = _label(sheet)
        address = format_cell(label, cell.row, cell.column)
        if sheet.type != "worksheet" or sheet.part is None:
            raise EditError("not_a_worksheet", addresses=[address])
        if self.workbook.strict:
            raise EditError("strict")
        if sheet.protected:
            raise EditError("protected", addresses=[address])
        with self.batch() as edit:
            writer = self.cells.sheet(sheet.stored_name)
            if writer is None:
                raise EditError("not_a_worksheet", addresses=[address])
            self._permitted(sheet, cell, value, writer, address)
            if not writer.write(cell, value, style=style):
                return False
            self.cells.drop_calc_chain()
            self.cells.request_full_calculation()
            edit.cells.setdefault(label, set()).add((cell.row, cell.column))
            edit.sheets.add(sheet.part)
            edit.rows.setdefault(sheet.part, set()).add(cell.row)
            self._written.setdefault(sheet.part, set()).add((cell.row, cell.column))
            if not _holds(writer.content(cell), value, self.cells.date1904):
                raise EditError("readback", addresses=[address])
        return True

    def _permitted(self, sheet: Any, cell: Cell, value: Any, writer: Any,
                   address: str) -> None:
        """The host's write policy asked again of one write, whatever planned it
        (:mod:`.policy`): the cell, a formula, a formula written over."""
        from .policy import PolicyError

        try:
            allowed = self.workbook._allowed()
        except PolicyError as error:
            raise EditError("bad_policy", error.message) from None
        if allowed is None:
            return
        if not allowed.allows(sheet.part, cell):
            raise EditError("not_writable", addresses=[address])
        if not allowed.formulas and isinstance(value, Formula) \
                or not allowed.replace_formulas and writer.content(cell).formula is not None:
            raise EditError("formulas_not_allowed", addresses=[address])

    def touched(self, sheet: Any, cells: "Iterable[Cell]") -> None:
        """Record ``cells`` of ``sheet`` as changed in the open batch (a shared formula
        group handed on: its members' formulas), for the report and the save's read-back."""
        edit = self._edit
        if edit is None:
            raise RuntimeError("cells are recorded inside a batch")
        label = _label(sheet)
        rows = edit.rows.setdefault(sheet.part, set())
        for cell in cells:
            edit.cells.setdefault(label, set()).add((cell.row, cell.column))
            rows.add(cell.row)
            self._written.setdefault(sheet.part, set()).add((cell.row, cell.column))
        edit.sheets.add(sheet.part)

    def touched_format(self, sheet: Any) -> None:
        """Record ``sheet`` as changed in the open batch by formatting or layout, which
        writes no value (:mod:`.format`, :mod:`.layout`): its structure is checked when the
        step closes, and no calculation is invalidated."""
        edit = self._edit
        if edit is None:
            raise RuntimeError("a sheet is recorded inside a batch")
        edit.sheets.add(sheet.part)
        edit.whole.add(sheet.part)

    # -- reports -----------------------------------------------------------------------------

    def _report(self, edit: Edit, before: int, change: Any) -> EditReport:
        version = self.history.version
        if version == before:
            return EditReport(before=before, version=version,
                              calculation=self.workbook.calculation_status)
        parts = tuple(sorted(change.parts))
        cells, cut, count = capped(edit.cells)
        # The step's check: a sheet whose cells alone were written is checked on the rows
        # written (structure.check_structure's ``rows``); one changed otherwise, whole.
        rows = {part: edit.rows.get(part, set()) for part in edit.sheets
                if part not in edit.whole}
        issues = check_structure(self.workbook, sheets=edit.sheets, rows=rows)
        return EditReport(before=before, version=version, cells=cells, cell_count=count,
                          cells_cut=cut, parts=parts,
                          calculation="invalidated" if count
                          else self.workbook.calculation_status,
                          issues=tuple(issues[:MAX_ISSUES]),
                          issues_cut=max(0, len(issues) - MAX_ISSUES))

    # -- saving ------------------------------------------------------------------------------

    def written(self) -> dict[str, set[tuple[int, int]]]:
        """The cells written since the workbook was opened, by sheet part."""
        return {part: set(cells) for part, cells in self._written.items()}


# -- helpers -----------------------------------------------------------------------------------


def _holds(content: Any, value: Any, date1904: bool) -> bool:
    """Whether the writer's ``content`` of a cell is what writing ``value`` asked for."""
    if isinstance(value, Formula):
        if content.formula != value.text or content.formula_type != "normal":
            return False
        value = value.value
        if value is None:
            return content.kind == "blank"
    elif content.formula is not None:
        return False
    if value is None:
        return content.kind == "blank"
    if isinstance(value, bool):
        return content.kind == "bool" and content.value is value
    if isinstance(value, (_dt.date, _dt.time)):
        value = date_serial(value, date1904=date1904)
    if isinstance(value, (int, float)):
        return content.kind == "number" and float(content.value) == float(value)
    if isinstance(value, RichText):
        value = value.text
    if isinstance(value, str):
        return content.kind == "text" and content.value == value
    return False


def has_vba_project(workbook: Any) -> bool:
    """Whether the workbook holds a VBA project: a part of its type, or one its workbook
    part relates to as one."""
    package = workbook.package
    for rel in package.relationships(workbook.part).values():
        if rel.type == REL_VBA_PROJECT and not rel.is_external and rel.target_part \
                and package.has_part(rel.target_part):
            return True
    return any(package.content_type(part) == CT_VBA_PROJECT for part in package.part_names)


#: The kinds a workbook is saved as.
SAVE_KINDS = tuple(name for name, kind in KINDS.items() if kind.application == "excel")


def serialise(workbook: Any, kind: str | None = None) -> bytes:
    """The workbook as a file of ``kind`` (its own when ``None``): the package as it is,
    with the main part's content type set to the kind's when that differs.  A kind
    without macros is refused while the workbook holds a VBA project."""
    kind = kind or workbook.kind
    if kind not in SAVE_KINDS:
        raise EditError("bad_kind", f"one of {', '.join(SAVE_KINDS)}")
    if not KINDS[kind].macro_enabled and has_vba_project(workbook):
        raise EditError("vba_project")
    if kind == workbook.kind:
        return workbook.to_bytes()
    package = workbook.package
    types = package.content_types_with(workbook.part, main_content_type(kind))
    return package.to_bytes({"[Content_Types].xml": types})


def verify(workbook: Any, data: bytes, kind: str) -> SaveReport:
    """Reopen ``data`` and read every cell written since ``workbook`` was opened back
    through xlsx-agent's reader: :class:`EditError` ``readback`` (with the addresses)
    when one does not hold what the writer holds.  The report counts the cells read back
    and lists the structure issues of what was saved."""
    from ..workbook import Workbook

    reopened = Workbook.open(data)
    if reopened.kind != kind:
        raise EditError("readback", f"the file saved is a {reopened.kind}, not a {kind}")
    written = workbook.editor.written() if workbook._editor is not None else {}
    by_part = {sheet.part: sheet for sheet in reopened._every_sheet if sheet.part}
    wrong: list[str] = []
    verified = 0
    for part, positions in written.items():
        sheet = by_part.get(part)
        writer = workbook.editor.cells.sheet(sheet.stored_name) if sheet is not None else None
        if sheet is None or writer is None:
            wrong.append(part)
            continue
        label = _label(sheet)
        # Only the cells written are typed (the pass places every cell, as a read does):
        # a sheet of many rows is read back for what was written, not typed whole.
        reader = reopened._reader(sheet)
        read = {}
        for row, column, hidden, node in reopened._stream(sheet, reader._walk):
            if (row, column) in positions:
                read[(row, column)] = reader.value(row, column, hidden, node)
        for row, column in sorted(positions):
            verified += 1
            if not _agrees(writer.content(Cell(row, column)), read.get((row, column))):
                wrong.append(format_cell(label, row, column))
    if wrong:
        log.warning("save refused: %d cell(s) did not read back", len(wrong))
        raise EditError("readback", "the file was not handed over", addresses=wrong)
    issues = tuple(check_structure(reopened))
    return SaveReport(kind=kind, size=len(data), verified=verified, issues=issues)


def _agrees(content: Any, value: Any) -> bool:
    """Whether xlsx-agent's reading of a saved cell (``value``, ``None`` when it read no
    cell there) agrees with the writer's ``content`` of it: the same formula (when the
    reader gives its text), type and value."""
    formula = content.formula is not None
    if value is None:
        return content.kind == "blank" and not formula
    if value.formula != formula:
        return False
    if formula and content.formula_type == "normal" and value.formula_text is not None \
            and value.formula_text != decode_xstring(content.formula):
        return False
    kind = content.kind
    if kind == "blank":
        return value.t == "blank"
    if kind == "text":       # the writer's as stored, an ST_Xstring; the reader's decoded
        return value.t in ("text", "list") and (value.original if value.cleaned
                                                else value.v) == decode_xstring(content.value)
    if kind == "bool":
        return value.t == "bool" and value.v is content.value
    if kind == "error":
        return value.t == "error" and value.v == content.value
    number = value.serial if value.serial is not None else value.v
    return value.t in ("number", "percent", "date", "time", "datetime") \
        and isinstance(number, (int, float)) and not isinstance(number, bool) \
        and float(number) == float(content.value)


def undo_scope(workbook: Any, scope: str) -> list[str]:
    """The parts a sheet's undo scope owns: the sheet's part and its relationships part
    (``undo``'s ``scope`` is a sheet's name, as an address writes it or as stored)."""
    from ooxml_edit.opc import rels_path_for
    from ooxml_edit.tools import ToolError

    try:
        sheet = workbook.sheet(scope.strip().rstrip("!").strip("'").replace("''", "'"))
    except (KeyError, ValueError):
        sheet = None
    if sheet is None or sheet.part is None:
        raise ToolError("not_found", f"no sheet {scope!r} to undo in", field="scope",
                        valid_options=[s.name for s in workbook.sheets][:50])
    return [sheet.part, rels_path_for(sheet.part)]


def undo_shared(workbook: Any) -> list[str]:
    """The parts every sheet's changes share: the workbook part and its relationships,
    the shared strings, the styles, the calculation chain and the content types."""
    from ooxml_edit.opc import rels_path_for

    out = [workbook.part, rels_path_for(workbook.part), "[Content_Types].xml"]
    for rel in workbook.package.relationships(workbook.part).values():
        if not rel.is_external and rel.target_part and rel.type.rpartition("/")[2] in (
                "sharedStrings", "styles", "calcChain"):
            out.append(rel.target_part)
    return out


from .policy import PolicyError, TableEntry, WritePolicy  # noqa: E402

__all__ = ["CT_VBA_PROJECT", "EDIT_CODES", "Edit", "EditError", "EditReport", "Editor",
           "PolicyError", "SAVE_KINDS", "STRUCTURE_CODES", "SaveReport", "TableEntry",
           "WritePolicy", "check_structure", "compress", "has_vba_project", "serialise",
           "undo_scope", "undo_shared", "verify"]
