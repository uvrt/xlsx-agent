"""Every cell compared by address, formulas by their text (0.2, X5): the integrity side of
:func:`~xlsx_agent.compare.compare` (``cells=True``)::

    diff = Workbook.compare(before, after, cells=True, policy=policy)
    diff.cells.counts()      # {"formula_to_literal": 1, "value_changed": 12}
    diff.cells.ranges("formula_to_literal")   # ({"Stock": ["F12"]}, 0, 1)
    diff.cells.outside       # the changes the policy does not allow, by address
    diff.cells.structure     # names, validations, merges, conditional formats changed
    diff.cells.as_dict()     # bounded: counts, capped ranges per kind, structure, outside

What a host asks before accepting a workbook back (a client's offline review, an agent's
step): did only the cells it allowed change, and did any formula change, or become a
value, while the cached values still look the same?  The tables compare
(:mod:`xlsx_agent.compare`) answers by key and by cached value; this one by address, over
every cell of every sheet the two workbooks share, and looks at the formulas themselves.

**Sheets** are paired as the compare pairs them (by ``sheetId``, else by name); a sheet in
one workbook only is in :attr:`Diff.sheets <xlsx_agent.compare.Diff.sheets>`, its cells not
compared.  Each change is named by the newer workbook's sheet.  What is withheld from the
model (a sheet, a declared table's hidden column: :mod:`xlsx_agent.declare`) is not
compared, as nothing else reads it: compare workbooks opened without withholding to cover
every cell.

**Kinds** (:data:`KINDS`), for a cell at one address in both:

* ``formula_changed``: a formula in both, not the same formula;
* ``formula_to_literal``: a formula in ``a``, a value without one in ``b``;
* ``literal_to_formula``: a value in ``a``, a formula in ``b``;
* ``formula_removed``: a formula in ``a``, nothing in ``b`` (no value, no formula);
* ``formula_added``: nothing in ``a``, a formula in ``b``;
* ``value_changed``: no formula in either, and the stored value differs (another value or
  type, or a value added or cleared);
* ``cache_changed``: the same formula in both, another cached value -- what a
  recalculation does, never a write: it is counted, and no policy is asked about it.

**Structure** (X6), listed apart from the cells (:attr:`CellsDiff.structure`,
:data:`STRUCTURE_KINDS`), on the sheets both workbooks have: defined names
(``name_added``, ``name_removed``, ``name_changed``: by name and scope, case aside; changed
when the refers-to differs as a formula does, normalised alike, or the hidden flag; a name
of a sheet withheld from the model is none); data validations, standard and ``x14`` as
one (``validation_added``, ``_removed``, ``_changed``: by the cells a rule covers, its
``sqref`` with its ranges sorted; changed when its type, operator -- ``between`` when the
file gives none, for the types that compare --, ``allowBlank`` or formulas differ, the
formulas normalised; its messages, and which form stores it, are not compared); merged
areas (``merge_added``, ``merge_removed``); conditional formats (``conditional_format_added``,
``_removed``, ``_changed``: by ``sqref``; changed when a rule's type, attributes, formulas
or definition -- a colour scale's stops, a data bar's -- differ, in priority order; the
priorities' numbers, the format a rule applies (``dxfId``, a style) and the ids linking a
rule to its ``x14`` part are not compared).  The sheets themselves (added, removed,
renamed) are :attr:`Diff.sheets <xlsx_agent.compare.Diff.sheets>`.  A change says its kind
and where (a name, or the sheet and its ranges, at most :data:`SQREF_CHARS` characters);
its formulas only with ``formula_text``.

**The same formula.**  Formulas are compared as :func:`~xlsx_agent.formula.normalise` reads
them (case and whitespace the grammar gives no meaning left out; ECMA-376 Part 1, §18.17).
A shared formula's member (``<f t="shared" si="0"/>``, ECMA-376 Part 1, §18.3.1.40) is
compared as its expanded text -- its group's formula moved from the group's first cell to
the member (:func:`~xlsx_agent.formula.shift`) -- so a group split into the same formulas
written one by one is no change.  An array formula is compared with the range it fills
(its ``ref``), and the other cells of its block as the block's (their values are its
results); a data table (``t="dataTable"``) by its attributes.  Values: text as stored (a
shared string, an inline string and a ``str`` cell holding one text are one value; rich
text's runs are not compared), numbers as numbers, booleans, errors and ISO dates as
written.  Styles are not compared.

**Against a policy** (:mod:`xlsx_agent.edit.policy`): ``policy`` (or ``a``'s own policies,
when it was opened with one) is resolved against ``a``, the workbook as it was, and each
change but ``cache_changed`` is *outside* it when its cell is not writable, when it writes
a formula under ``formulas: false`` (``formula_changed``, ``literal_to_formula``,
``formula_added``), or replaces one under ``replace_formulas: false``
(``formula_changed``, ``formula_to_literal``, ``formula_removed``).  A sheet added,
removed or renamed is outside any policy (``sheets_outside``), and so is every structural
change (``outside``'s ``structure``): a policy lets cell values be written, nothing else.

**Bounded, and no content.**  :meth:`CellsDiff.as_dict` gives counts by kind and each
kind's cells as ranges, capped as an edit report's are (:data:`~xlsx_agent.edit.report.
MAX_RANGES` per kind, the rest counted), the structural changes counted by kind and the
first :data:`~xlsx_agent.edit.report.MAX_RANGES` of them by kind and place (``cut`` the
rest), and the same for the changes outside the policy.
A change carries its address and kind; its formulas' text only when the caller asks
(``formula_text=True``), and never a formula a redaction matches or that reaches a cell
hidden from the model (``withheld``).  Nothing here is logged.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
from typing import TYPE_CHECKING, Any, Iterable, Iterator

from ooxml_edit.sheets import Area, Cell

from .address import quote_sheet
from .edit.report import MAX_RANGES, capped
from .formula import ShiftError, normalise, shift
from .read.stream import _F, _IS, _V, _area, ConditionalFormat, Validation
from .read.strings import SharedStrings, item_text
from .read.types import _letters

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: Every kind of change a cell can show (see the module's notes).
KINDS = ("formula_changed", "formula_to_literal", "literal_to_formula", "formula_removed",
         "formula_added", "value_changed", "cache_changed")
#: The kinds a write makes: what a policy is asked about.
WRITES = KINDS[:-1]
#: The kinds that write a formula, and those that replace one.
FORMULA_WRITTEN = ("formula_changed", "literal_to_formula", "formula_added")
FORMULA_REPLACED = ("formula_changed", "formula_to_literal", "formula_removed")
#: Every kind of structural change (see the module's notes), apart from the cells': the
#: sheets' own changes are :attr:`Diff.sheets <xlsx_agent.compare.Diff.sheets>`.
STRUCTURE_KINDS = ("name_added", "name_removed", "name_changed", "validation_added",
                   "validation_removed", "validation_changed", "merge_added", "merge_removed",
                   "conditional_format_added", "conditional_format_removed",
                   "conditional_format_changed")
#: The most characters of a range list (``sqref``) a structural change gives; longer ones
#: are cut with ``…``.
SQREF_CHARS = 120
#: The sheet types that hold cells.
_CELL_SHEETS = ("worksheet", "macrosheet", "dialogsheet")


@dataclass(frozen=True)
class CellEdit:
    """One cell that differs: ``sheet`` as an address writes it (the newer workbook's
    name), ``row``, ``column``, ``kind`` (:data:`KINDS`); ``a`` and ``b`` each side's
    formula text when the compare was asked for it (``None`` for a side without one, or
    ``withheld``)."""

    sheet: str
    row: int
    column: int
    kind: str
    a: str | None = field(default=None, repr=False)
    b: str | None = field(default=None, repr=False)
    withheld: bool = False

    @property
    def address(self) -> str:
        return f"{self.sheet}!{_letters(self.column)}{self.row}"

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"at": self.address, "kind": self.kind}
        if self.withheld:
            out["withheld"] = True
        for side in ("a", "b"):
            text = getattr(self, side)
            if text is not None:
                out[side] = text
        return out


@dataclass(frozen=True)
class StructureChange:
    """One structural change: ``kind`` (:data:`STRUCTURE_KINDS`), ``sheet`` as an address
    writes it (the newer workbook's name; ``None`` for a workbook-wide name), ``at`` what
    changed -- a defined name, or the cells a rule or a merge covers (``D10:D200``) --
    and, when the compare was asked for formulas, ``a`` and ``b`` each side's formulas (a
    name's refers-to, a rule's formulas; empty for a side without it), ``withheld`` when a
    redaction matches one or it reaches a cell hidden from the model."""

    kind: str
    sheet: str | None
    at: str
    a: tuple[str, ...] = field(default=(), repr=False)
    b: tuple[str, ...] = field(default=(), repr=False)
    withheld: bool = False

    @property
    def address(self) -> str:
        return self.at if self.sheet is None else f"{self.sheet}!{self.at}"

    def as_dict(self, *, formulas: bool = True) -> dict[str, Any]:
        out: dict[str, Any] = {"kind": self.kind, "at": self.address}
        if formulas:
            if self.withheld:
                out["withheld"] = True
            for side in ("a", "b"):
                if getattr(self, side):
                    out[side] = list(getattr(self, side))
        return out


@dataclass(frozen=True)
class CellsDiff:
    """Every shared sheet's cells compared (see the module's notes): ``sheets`` the sheets
    compared, as the newer workbook names them; ``edits`` the cells that differ, sheet by
    sheet, in row then column order; ``outside`` those a policy does not allow (``None``
    without a policy) and ``sheets_outside`` the sheets added, removed or renamed then;
    ``formula_text`` whether the edits carry their formulas."""

    sheets: tuple[str, ...]
    edits: tuple[CellEdit, ...]
    outside: "tuple[CellEdit, ...] | None" = None
    sheets_outside: int = 0
    formula_text: bool = False
    structure: "tuple[StructureChange, ...]" = ()

    @property
    def same(self) -> bool:
        return not self.edits and not self.structure

    def structure_counts(self) -> dict[str, int]:
        """The structural changes counted by kind, in :data:`STRUCTURE_KINDS` order."""
        found: dict[str, int] = {}
        for change in self.structure:
            found[change.kind] = found.get(change.kind, 0) + 1
        return {kind: found[kind] for kind in STRUCTURE_KINDS if kind in found}

    def counts(self, edits: "Iterable[CellEdit] | None" = None) -> dict[str, int]:
        """The changes counted by kind, in :data:`KINDS` order (zeros left out)."""
        found: dict[str, int] = {}
        for edit in self.edits if edits is None else edits:
            found[edit.kind] = found.get(edit.kind, 0) + 1
        return {kind: found[kind] for kind in KINDS if kind in found}

    def ranges(self, kind: str | None = None, edits: "Iterable[CellEdit] | None" = None
               ) -> tuple[dict[str, list[str]], int, int]:
        """The cells of one kind (every kind: ``None``) as ranges by sheet, capped
        (:func:`~xlsx_agent.edit.report.capped`): the ranges, how many were cut, the
        cells."""
        by_sheet: dict[str, list[tuple[int, int]]] = {}
        for edit in self.edits if edits is None else edits:
            if kind is None or edit.kind == kind:
                by_sheet.setdefault(edit.sheet, []).append((edit.row, edit.column))
        return capped(by_sheet)

    def as_dict(self) -> dict[str, Any]:
        """Counts, each kind's cells as capped ranges, and what is outside the policy:
        addresses and kinds only (the formulas are on :attr:`edits`)."""
        out: dict[str, Any] = {"sheets": len(self.sheets), "counts": self.counts()}
        out.update(_by_kind(self, self.edits))
        if self.structure:
            shown = self.structure[:MAX_RANGES]
            out["structure"] = {"counts": self.structure_counts(),
                                "changes": [change.as_dict(formulas=False)
                                            for change in shown]}
            if len(self.structure) > len(shown):
                out["structure"]["cut"] = len(self.structure) - len(shown)
        if self.outside is not None:
            outside: dict[str, Any] = {"count": len(self.outside)}
            if self.outside:
                outside["counts"] = self.counts(self.outside)
                ranges, cut, _ = self.ranges(edits=self.outside)
                outside["cells"] = ranges
                if cut:
                    outside["cut"] = cut
            if self.sheets_outside:
                outside["sheets"] = self.sheets_outside
            if self.structure:
                outside["structure"] = len(self.structure)
            out["outside"] = outside
        return out


def _by_kind(diff: CellsDiff, edits: tuple[CellEdit, ...]) -> dict[str, Any]:
    cells: dict[str, Any] = {}
    cuts: dict[str, int] = {}
    for kind in diff.counts(edits):
        ranges, cut, _ = diff.ranges(kind, edits)
        cells[kind] = ranges
        if cut:
            cuts[kind] = cut
    out: dict[str, Any] = {"cells": cells} if cells else {}
    if cuts:
        out["cut"] = cuts
    return out


def compare_cells(a: "Workbook", b: "Workbook", *, formula_text: bool = False,
                  policy: Any = None) -> CellsDiff:
    """``a`` (the older) against ``b`` (the newer), cell by cell (see the module's notes);
    ``policy`` a write policy or its plain data (``a``'s own policies when ``None``)."""
    from .edit.policy import Allowed, as_policy

    policies = (as_policy(policy),) if policy is not None else tuple(a.policies)
    allowed = Allowed(a, policies) if policies else None
    earlier = {_ident(s): s for s in a.sheets if _holds_cells(s)}
    sheets: list[str] = []
    edits: list[CellEdit] = []
    structure: list[StructureChange] = []
    for sheet in b.sheets:
        other = earlier.get(_ident(sheet)) if _holds_cells(sheet) else None
        if other is None:
            continue
        name = quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}"
        sheets.append(name)
        edits.extend(_sheet(a, other, b, sheet, name, formula_text, structure))
    structure[:0] = _names(a, b, formula_text)
    outside = None
    sheets_outside = 0
    if allowed is not None:
        parts = {quote_sheet(s.name) if s.name else f"#{s.id}": earlier[_ident(s)].part
                 for s in b.sheets if _ident(s) in earlier and _holds_cells(s)}
        outside = tuple(e for e in edits if _outside(e, allowed, parts[e.sheet]))
        sheets_outside = _sheet_changes(a, b)
    return CellsDiff(sheets=tuple(sheets), edits=tuple(edits), outside=outside,
                     sheets_outside=sheets_outside, formula_text=formula_text,
                     structure=tuple(structure))


def _holds_cells(sheet: "Sheet") -> bool:
    return sheet.part is not None and sheet.type in _CELL_SHEETS


def _ident(sheet: "Sheet") -> tuple:
    return ("id", sheet.id) if sheet.id is not None else ("name", sheet.name.casefold())


def _sheet_changes(a: "Workbook", b: "Workbook") -> int:
    from .compare import _sheets

    return sum(1 for _ in _sheets(a.sheets, b.sheets))


def _outside(edit: CellEdit, allowed: Any, part: str | None) -> bool:
    if edit.kind == "cache_changed":
        return False
    if not allowed.allows(part, Cell(edit.row, edit.column)):
        return True
    if not allowed.formulas and edit.kind in FORMULA_WRITTEN:
        return True
    return not allowed.replace_formulas and edit.kind in FORMULA_REPLACED


# -- structure ---------------------------------------------------------------------------------


def _names(a: "Workbook", b: "Workbook", formula_text: bool) -> list[StructureChange]:
    """The defined names added, removed, or whose refers-to (normalised as a formula is) or
    hidden flag changed, by name and scope, case aside."""
    def by_key(workbook: "Workbook") -> dict:
        out: dict[tuple[str, str | None], Any] = {}
        for name in workbook.names:
            if name.bad_scope is not None:
                continue
            scope = name.scope.casefold() if name.scope else None
            out.setdefault((name.name.casefold(), scope), name)
        return out

    old, new = by_key(a), by_key(b)
    changes: list[StructureChange] = []
    for key in [*new, *[k for k in old if k not in new]]:
        one, two = old.get(key), new.get(key)
        if one is not None and two is not None:
            if normalise(one.formula) == normalise(two.formula) and one.hidden == two.hidden:
                continue
            kind = "name_changed"
        else:
            kind = "name_added" if one is None else "name_removed"
        shown = two if two is not None else one
        assert shown is not None   # a key of one of them
        sheet = quote_sheet(shown.scope) if shown.scope else None

        texts: list[tuple[str, ...]] = [(), ()]
        withheld = False
        if formula_text:
            for n, (workbook, found) in enumerate(((a, one), (b, two))):
                if found is not None and found.formula:
                    if _text_withheld(workbook, found.scope or "", found.formula, Cell(1, 1)):
                        withheld = True
                    else:
                        texts[n] = (found.formula,)
        changes.append(StructureChange(kind, sheet, shown.name, texts[0], texts[1], withheld))
    return changes


def _sheet_structure(a: "Workbook", sheet_a: "Sheet", held_a: dict, b: "Workbook",
                     sheet_b: "Sheet", held_b: dict, name: str,
                     formula_text: bool) -> list[StructureChange]:
    """One sheet's merges, data validations (standard and ``x14``, as one: the same rule
    stored in the other form is no change) and conditional formats compared, each by the
    cells it covers."""
    out: list[StructureChange] = []
    merges_a = {_ref(area) for area in held_a.get("merges", ())}
    merges_b = {_ref(area) for area in held_b.get("merges", ())}
    for ref in sorted(merges_b - merges_a, key=_ref_order):
        out.append(StructureChange("merge_added", name, ref))
    for ref in sorted(merges_a - merges_b, key=_ref_order):
        out.append(StructureChange("merge_removed", name, ref))
    for kind, read, compared, texts in (
            ("validation", "validations", _rule, _rule_texts),
            ("conditional_format", "formats", _format, _format_texts)):
        old = _grouped(held_a.get(read, ()))
        new = _grouped(held_b.get(read, ()))
        for sqref in [*new, *[k for k in old if k not in new]]:
            one, two = old.get(sqref), new.get(sqref)
            if one is not None and two is not None:
                if [compared(r) for r in one] == [compared(r) for r in two]:
                    continue
                change = f"{kind}_changed"
            else:
                change = f"{kind}_added" if one is None else f"{kind}_removed"
            shown = sqref if len(sqref) <= SQREF_CHARS else sqref[:SQREF_CHARS - 1] + "…"
            sides: list[tuple[str, ...]] = [(), ()]
            withheld = False
            if formula_text:
                anchor = _first_cell(sqref)
                for n, (workbook, sheet, rules) in enumerate(((a, sheet_a, one),
                                                              (b, sheet_b, two))):
                    found = tuple(t for r in rules or () for t in texts(r))
                    if any(_text_withheld(workbook, sheet.name, t, anchor) for t in found):
                        withheld = True
                    else:
                        sides[n] = found
            out.append(StructureChange(change, name, shown, sides[0], sides[1], withheld))
    return out


def _grouped(items: Iterable) -> dict[str, list]:
    """Rules by the cells they cover (their ``sqref``, its ranges sorted)."""
    out: dict[str, list] = {}
    for item in items:
        out.setdefault(_sqref(item.sqref), []).append(item)
    return out


def _sqref(text: str) -> str:
    parts = text.upper().split()
    areas = [area for area in (_area(part) for part in parts) if area is not None]
    if len(areas) == len(parts):
        return " ".join(_ref(area) for area in sorted(set(areas), key=_area_order))

    return " ".join(parts)


def _ref(area: Area) -> str:
    first = f"{_letters(area.first.column)}{area.first.row}"
    if area.first == area.last:
        return first
    return f"{first}:{_letters(area.last.column)}{area.last.row}"


def _area_order(area: Area) -> tuple:
    return (area.first.row, area.first.column, area.last.row, area.last.column)


def _ref_order(ref: str) -> tuple:
    area = _area(ref)
    return _area_order(area) if area is not None else (0, 0, 0, 0)


def _first_cell(sqref: str) -> Cell:
    area = _area(sqref.split()[0]) if sqref.split() else None
    return area.first if area is not None else Cell(1, 1)


#: The validation types whose operator counts (ECMA-376 Part 1, §18.3.1.32: ``between``
#: when the file gives none).
_COMPARING = ("whole", "decimal", "date", "time", "textLength")


def _rule(rule: Validation) -> tuple:
    """A data validation as compared: its type, operator (defaulted), ``allowBlank`` and
    formulas, normalised; its messages and the form it is stored in are not compared."""
    operator = (rule.operator or "between") if rule.type in _COMPARING else None
    return (rule.type, operator, rule.allow_blank,
            normalise(rule.formula1) if rule.formula1 else None,
            normalise(rule.formula2) if rule.formula2 else None)


def _rule_texts(rule: Validation) -> tuple[str, ...]:
    return tuple(t for t in (rule.formula1, rule.formula2) if t)


def _format(found: ConditionalFormat) -> tuple:
    """A conditional format as compared: its rules in priority order, each element with
    its formulas normalised (the format a rule applies is a style, not compared)."""
    return tuple(_normalised(rule) for rule in found.rules)


def _normalised(node: tuple) -> tuple:
    name, attrs, text, children = node
    if name in ("formula", "f") and text:
        text = normalise(text)
    return (name, attrs, text, tuple(_normalised(child) for child in children))


def _format_texts(found: ConditionalFormat) -> tuple[str, ...]:
    out: list[str] = []

    def walk(node: tuple) -> None:
        name, _, text, children = node
        if name in ("formula", "f") and text:
            out.append(text)
        for child in children:
            walk(child)

    for rule in found.rules:
        walk(rule)
    return tuple(out)


def _text_withheld(workbook: "Workbook", sheet: str, text: str, cell: Cell) -> bool:
    redaction = workbook.redaction
    if redaction is not None and redaction.holds(text):
        return True
    return workbook.regions.formula_withheld(text, sheet, cell)


# -- one sheet ---------------------------------------------------------------------------------

#: A cell as read here: its formula as stored (``None`` without one: see :func:`_cells`)
#: and its value (its cached value, for a formula; ``None`` for none).
_State = tuple


def _sheet(a: "Workbook", sheet_a: "Sheet", b: "Workbook", sheet_b: "Sheet", name: str,
           formula_text: bool, structure: "list[StructureChange]") -> Iterator[CellEdit]:
    held_a: dict = {}
    held_b: dict = {}
    old = _cells(a, sheet_a, held_a)
    new = _cells(b, sheet_b, held_b)
    structure.extend(_sheet_structure(a, sheet_a, held_a, b, sheet_b, held_b, name,
                                      formula_text))
    hidden = [*a.regions.withheld(sheet_a.name), *b.regions.withheld(sheet_b.name)]
    for row, column in sorted(old.keys() | new.keys()):
        if hidden and any(area.contains(Cell(row, column)) for area in hidden):
            continue
        one, two = old.get((row, column)), new.get((row, column))
        kind = _kind(one, two)
        if kind is None:
            continue
        if not formula_text:
            yield CellEdit(name, row, column, kind)
            continue
        texts = []
        withheld = False
        for workbook, sheet, state in ((a, sheet_a, one), (b, sheet_b, two)):
            text = _text(state[0]) if state is not None and state[0] is not None else None
            if text is not None and _withheld(workbook, sheet, text, row, column):
                text, withheld = None, True
            texts.append(text)
        yield CellEdit(name, row, column, kind, texts[0], texts[1], withheld)


def _kind(one: "_State | None", two: "_State | None") -> str | None:
    fa = one[0] if one is not None else None
    fb = two[0] if two is not None else None
    va = one[1] if one is not None else None
    vb = two[1] if two is not None else None
    if fa is not None and fb is not None:
        if fa != fb and _compared(fa) != _compared(fb):
            return "formula_changed"
        return "cache_changed" if va != vb else None
    if fa is not None:
        return "formula_to_literal" if vb is not None else "formula_removed"
    if fb is not None:
        return "literal_to_formula" if va is not None else "formula_added"
    return "value_changed" if va != vb else None


@lru_cache(maxsize=4096)
def _compared(formula: tuple) -> tuple:
    """A formula as stored (:func:`_cells`) as it is compared: its text normalised
    (:func:`~xlsx_agent.formula.normalise`), a shared member's as its expanded text (or,
    when its group's text cannot be moved, by the group's text and its place in it).  Only
    asked when two cells' stored forms differ: most formulas are compared as stored."""
    kind = formula[0]
    if kind == "normal":
        return ("normal", normalise(formula[1]))
    if kind == "member":
        _, text, rows, columns = formula
        try:
            return ("normal", normalise(shift(text, rows, columns)))
        except ShiftError:
            return ("member", normalise(text), rows, columns)
    if kind == "array":
        return ("array", normalise(formula[1]), formula[2])
    if kind == "dataTable":
        return ("dataTable", formula[1], formula[2], normalise(formula[3]))
    return formula


def _text(formula: tuple) -> str | None:
    """A formula's text as a cell holds it (a shared member's expanded); ``None`` for a
    block's other cell or text that cannot be moved."""
    kind = formula[0]
    if kind in ("normal", "array"):
        return formula[1]
    if kind == "dataTable":
        return formula[3] or None
    if kind == "member":
        try:
            return shift(formula[1], formula[2], formula[3])
        except ShiftError:
            return None
    return None


def _withheld(workbook: "Workbook", sheet: "Sheet", text: str, row: int, column: int) -> bool:
    redaction = workbook.redaction
    if redaction is not None and redaction.holds(text):
        return True
    return workbook.regions.formula_withheld(text, sheet.name, Cell(row, column))


def _strings(workbook: "Workbook") -> list[str]:
    """The shared strings as stored: a redaction would make two texts one."""
    if workbook.redaction is None:
        return workbook.strings.items
    part = workbook._related("sharedStrings")
    return SharedStrings(workbook.package.read(part) if part else None).items


def _cells(workbook: "Workbook", sheet: "Sheet",
           structure: "dict | None" = None) -> dict[tuple[int, int], _State]:
    """Every cell of ``sheet`` that holds a value or a formula: its formula as stored --
    ``("normal", text)``, ``("member", group text, rows, columns)`` for a shared group's
    member (its place from the group's first cell), ``("array", text, ref)``,
    ``("dataTable", ref, attributes, text)``, ``("block", ref)`` for an array's or data
    table's other cell, ``("shared", si)`` for a member whose group has no first cell --
    and its value.  ``structure``, when given, is filled with what the same pass met of the
    sheet's structure: ``merges``, ``validations``, ``formats`` (conditional)."""
    strings = _strings(workbook)
    reader = workbook._reader(sheet)
    out: dict[tuple[int, int], _State] = {}
    groups: dict[str, tuple[int, int, str]] = {}
    members: list[tuple[int, int, str]] = []
    blocks: list[tuple[Area, tuple[int, int], str | None]] = []

    def run() -> None:
        for row, column, _, node in reader._walk():
            kind = node.get("t") or "n"
            v_text = f_node = is_node = None
            for child in node:
                tag = child.tag
                if tag in _V:
                    v_text = child.text or ""
                elif tag in _F:
                    f_node = child
                elif tag in _IS:
                    is_node = child
            value = _value(kind, v_text, is_node, strings)
            if f_node is None:
                if value is not None:
                    out[(row, column)] = (None, value)
                continue
            text = f_node.text or ""
            form = f_node.get("t") or "normal"
            if form == "shared":
                group = f_node.get("si") or ""
                if text:
                    groups.setdefault(group, (row, column, text))
                    out[(row, column)] = (("normal", text), value)
                else:
                    members.append((row, column, group))
                    out[(row, column)] = (("shared", group), value)
            elif form in ("array", "dataTable"):
                ref = (f_node.get("ref") or "").upper()
                if form == "array":
                    out[(row, column)] = (("array", text, ref), value)
                else:
                    attrs = tuple(f_node.get(n) for n in ("r1", "r2", "dt2D", "dtr", "del1",
                                                          "del2"))
                    out[(row, column)] = (("dataTable", ref, attrs, text), value)
                block = _area(ref)
                if block is not None:
                    blocks.append((block, (row, column), ref))
            else:
                out[(row, column)] = (("normal", text), value)

    from .workbook import _parsed

    _parsed(sheet.part or f"sheet #{sheet.id}", run)
    if structure is not None:
        structure.update(merges=reader.merges, validations=reader.validations,
                         formats=reader.conditional_formats)
    for row, column, group in members:
        master = groups.get(group)
        if master is not None:
            first_row, first_column, text = master
            out[(row, column)] = (("member", text, row - first_row, column - first_column),
                                  out[(row, column)][1])
    for block, anchor, ref in blocks:
        for position, state in list(_inside(out, block)):
            if position != anchor and state[0] is None:
                out[position] = (("block", ref), state[1])
    return out


def _inside(cells: dict, block: Area) -> Iterator[tuple[tuple[int, int], _State]]:
    """The cells of ``cells`` inside ``block``: by position when the block is small, by a
    pass over the cells otherwise."""
    size = (block.last.row - block.first.row + 1) * (block.last.column - block.first.column + 1)
    if size <= len(cells):
        for row in range(block.first.row, block.last.row + 1):
            for column in range(block.first.column, block.last.column + 1):
                state = cells.get((row, column))
                if state is not None:
                    yield (row, column), state
        return
    for position, state in cells.items():
        if block.contains(Cell(*position)):
            yield position, state


def _value(kind: str, v_text: str | None, is_node: Any, strings: list[str]) -> Any:
    """A cell's stored value as compared: ``(kind, value)``, ``None`` for none."""
    if kind == "inlineStr":
        if is_node is not None:
            return ("s", item_text(is_node))
        return ("s", v_text) if v_text is not None else None
    if v_text is None:
        return None
    if kind == "s":
        try:
            index = int(v_text)
            if index >= 0:
                return ("s", strings[index])
        except (ValueError, IndexError):
            pass
        return ("bad", v_text)
    if kind == "str":
        return ("s", v_text)
    if kind in ("b", "e", "d"):
        return (kind, v_text.strip())
    if not v_text.strip():
        return None
    try:
        return ("n", float(v_text))
    except ValueError:
        return ("bad", v_text)


__all__ = ["CellEdit", "CellsDiff", "FORMULA_REPLACED", "FORMULA_WRITTEN", "KINDS",
           "STRUCTURE_KINDS", "StructureChange", "WRITES", "compare_cells"]
