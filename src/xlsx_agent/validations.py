"""Data-validation rules, inspected (0.2, X5): what each sheet's rules allow and where a
list takes its values from, without evaluating anything::

    rules = wb.validations()                  # every sheet's, in the workbook's order
    rules[0].origin, rules[0].type             # ("x14", "list")
    rules[0].source, rules[0].resolved         # ("name", True): through a defined name
    rules[0].range                             # "Lists!$A$2:$A$5" -- where the values are
    [r.cells for r in rules if r.resolved is False]   # the rules no static reading resolves

**What a rule is** (ECMA-376 Part 1, §18.3.1.32 ``dataValidation``): its type
(ST_DataValidationType: ``none``, ``whole``, ``decimal``, ``list``, ``date``, ``time``,
``textLength``, ``custom``), its operator (ST_DataValidationOperator; the attribute's
default, ``between``, when the file gives none, for the types that compare: ``whole``,
``decimal``, ``date``, ``time``, ``textLength``), ``allowBlank``, the cells it covers
(``sqref``) and its formulas.  A rule Excel writes in the worksheet's ``extLst`` -- the
``x14:dataValidation`` of Microsoft's [MS-XLSX] extensions, with ``xm:sqref`` and
``x14:formula1/xm:f``, which Excel uses for a list on another sheet -- is read alike and
says ``origin="x14"``; one in ``dataValidations`` says ``"standard"``.  The rules are given
in the file's order: the standard ones, then the ``x14`` ones (the order the parts lie in).

**A list's source** (``source``), read from its first formula's tokens
(:mod:`xlsx_agent.formula`):

* ``inline``: one string literal, ``"S,M,L"`` -- its items, split at the commas;
* ``range``: one reference, on the rule's sheet or another (``Lists!$A$2:$A$5``);
* ``name``: one defined name -- the one local to the rule's sheet first, then the
  workbook-wide one, as Excel looks a name up from a sheet; followed through a name that
  is another name (at most :data:`NAME_DEPTH`), resolved when it ends at a range, else
  unresolved with what it ends at as the ``reason``;
* ``table``: one structured reference to a table's column (``Rates[Category]``): that
  column's data cells;
* ``indirect``, ``offset``: a formula calling ``INDIRECT`` or ``OFFSET`` (a dependent list,
  ``INDIRECT($C10)``, picks its list by another cell's value; a dynamic range grows with
  its data): never resolved, since only an evaluator could say where it points, and
  reported as such, by address, with ``uses``, the references and names the formula holds
  (as written: a relative reference is relative to the rule's first cell);
* ``formula``: any other formula, unresolved alike; ``external``: another workbook's cells,
  never followed.

``resolved`` is ``True`` when the values are known statically (an inline list, or a range
read), ``False`` otherwise (``reason``: ``indirect``, ``offset``, ``formula``,
``external``, ``unknown_name``, ``unknown_sheet``, ``unknown_table``, ``ref_error``, or
``withheld`` for a range over cells a declared table hides from the model:
:mod:`xlsx_agent.declare`); ``None`` for a rule of another type.  ``count`` is how many
values a resolved list allows (blanks aside, at most :data:`~xlsx_agent.profile.
ALLOWED_CAP` read; ``cut`` past it).

**A value written, checked** (:func:`check`, what :meth:`Workbook.set_cells
<xlsx_agent.Workbook.set_cells>` reports as ``invalid`` and ``unchecked``: a warning, never
a refusal).  A cell under a rule is ``invalid`` when the value written does not satisfy a
rule that can be read without evaluating anything: a resolved list (the value among its
values, text matched case and surrounding spaces aside, numbers as numbers -- how Excel
matches a list's text is not measured, so only a value no reading could match is named);
``whole`` (a whole number), ``decimal`` (a number) and ``date`` (a date or a number,
compared as its serial: a date is its serial number, ECMA-376 Part 1, §18.17.4) and
``textLength`` (the text's length in characters), with each bound a number, or one cell
holding a number and no formula; text or a boolean under ``whole``, ``decimal`` or
``date`` satisfies none (ST_DataValidationType, §18.18.21: the rule checks for values of
its type).  It is ``unchecked`` when the rule cannot be read so
(an unresolved list, a bound given by a formula such as ``TODAY()``, a ``time`` or
``custom`` rule, a formula written); a blank written is not checked.  Only the first rule
covering a cell is asked, as a cell has one rule in Excel.

**Content** (ROADMAP.md, "Design rules", 7).  A rule's formulas and a list's values are
content: they are on the rule for a program (``formula1``, ``formula2``, ``values``) and in
:meth:`ValidationRule.as_dict` only when asked; the rest -- addresses, kinds, counts, names -- is
always there.  Nothing here is logged.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Iterable

from ooxml_edit.sheets import Area, Cell

from .address import RangeAddress, quote_sheet
from .formula import ShiftError, Token, tokens
from .read.names import range_of, sheet_names
from .read.stream import Validation
from .read.types import _letters

if TYPE_CHECKING:  # pragma: no cover
    from .workbook import Sheet, Workbook

#: ST_DataValidationType's values that compare a value with one or two bounds (and so have
#: an operator).
COMPARING = ("whole", "decimal", "date", "time", "textLength")
#: Every :attr:`ValidationRule.source` of a list.
SOURCES = ("inline", "range", "name", "table", "indirect", "offset", "formula", "external")
#: How many names a list's source is followed through (a name that is another name).
NAME_DEPTH = 8
#: The most ranges of a rule's cells :meth:`ValidationRule.as_dict` lists; the rest are counted.
CELLS_SHOWN = 20


@dataclass(frozen=True)
class ValidationRule:
    """One data-validation rule, inspected (see the module's notes).

    ``sheet`` the sheet as an address writes it; ``index`` the rule's place among the
    sheet's (from 0, standard rules first); ``origin`` ``"standard"`` or ``"x14"``;
    ``type``; ``operator`` (``between`` when the file gives none, for a comparing type;
    ``None`` for any other); ``allow_blank``; ``cells`` the ranges it covers, as written
    (``("C10:C200",)``).  A list's ``source``, ``resolved``, ``reason``, ``range`` (where
    its values are, an address), ``name`` (the defined name it went through), ``count``,
    ``cut`` and ``uses``; ``None`` or empty for other types.  ``formula1``, ``formula2`` and
    ``values`` are content: for a program."""

    sheet: str
    index: int
    origin: str
    type: str
    operator: str | None
    allow_blank: bool
    cells: tuple[str, ...]
    areas: tuple[Area, ...] = field(default=(), repr=False)
    source: str | None = None
    resolved: bool | None = None
    reason: str | None = None
    range: str | None = None
    name: str | None = None
    count: int | None = None
    cut: bool = False
    uses: tuple[str, ...] = ()
    formula1: str | None = field(default=None, repr=False)
    formula2: str | None = field(default=None, repr=False)
    values: tuple | None = field(default=None, repr=False)

    @property
    def address(self) -> str:
        """The rule's cells with the sheet: ``Stock!C10:C200`` (the first range, and ``…``
        after it when there are more)."""
        first = f"{self.sheet}!{self.cells[0]}" if self.cells else self.sheet
        return first + (" …" if len(self.cells) > 1 else "")

    def covers(self, cell: Cell) -> bool:
        """Whether the rule covers ``cell`` of its sheet."""
        return any(area.contains(cell) for area in self.areas)

    def as_dict(self, *, formulas: bool = False, values: bool = False) -> dict[str, Any]:
        """The rule as plain data: everything but content, and with ``formulas`` its
        formulas, with ``values`` a resolved list's values."""
        out: dict[str, Any] = {"sheet": self.sheet, "origin": self.origin, "type": self.type}
        if self.operator is not None:
            out["operator"] = self.operator
        if self.allow_blank:
            out["allow_blank"] = True
        out["cells"] = " ".join(self.cells[:CELLS_SHOWN])
        if len(self.cells) > CELLS_SHOWN:
            out["more_cells"] = len(self.cells) - CELLS_SHOWN
        for name in ("source", "resolved", "reason", "range", "name", "count"):
            value = getattr(self, name)
            if value is not None:
                out[name] = value
        if self.cut:
            out["cut"] = True
        if self.uses:
            out["uses"] = list(self.uses)
        if formulas:
            for name in ("formula1", "formula2"):
                value = getattr(self, name)
                if value is not None:
                    out[name] = value
        if values and self.values is not None:
            out["values"] = list(self.values)
        return out


def inspect(workbook: "Workbook", sheets: "Iterable[Sheet] | None" = None) -> list[ValidationRule]:
    """Every rule of ``sheets`` (every sheet holding cells, by default), inspected."""
    out: list[ValidationRule] = []
    for sheet in sheets if sheets is not None else workbook.sheets:
        if sheet.part is None or sheet.type not in ("worksheet", "macrosheet", "dialogsheet"):
            continue
        rules = workbook.scan(sheet).validations
        ordered = [r for r in rules if r.origin == "standard"] + \
            [r for r in rules if r.origin != "standard"]
        for index, rule in enumerate(ordered):
            out.append(_rule(workbook, sheet, index, rule))
    return out


def _rule(workbook: "Workbook", sheet: "Sheet", index: int,
          rule: Validation) -> ValidationRule:
    kind = rule.type
    operator = (rule.operator or "between") if kind in COMPARING else None
    cells = tuple(rule.sqref.split()) if rule.sqref else tuple(_ref(a) for a in rule.areas)
    found = _source(workbook, sheet.name, rule.formula1) if kind == "list" else {}
    return ValidationRule(
        sheet=quote_sheet(sheet.name) if sheet.name else f"#{sheet.id}", index=index,
        origin=rule.origin, type=kind, operator=operator, allow_blank=rule.allow_blank,
        cells=cells, areas=rule.areas, formula1=rule.formula1, formula2=rule.formula2,
        **found)


def _ref(area: Area) -> str:
    first = f"{_letters(area.first.column)}{area.first.row}"
    if area.first == area.last:
        return first
    return f"{first}:{_letters(area.last.column)}{area.last.row}"


# -- a list's source ---------------------------------------------------------------------------


def _source(workbook: "Workbook", sheet: str, formula: str | None) -> dict[str, Any]:
    text = (formula or "").strip()
    if text.startswith("="):
        text = text[1:].lstrip()
    if not text:
        return {"source": "formula", "resolved": False, "reason": "formula"}
    try:
        found = [t for t in tokens(text) if not (t.kind == "other" and not t.text.strip())]
    except ShiftError:
        return {"source": "formula", "resolved": False, "reason": "formula"}
    if len(found) == 1:
        token = found[0]
        if token.kind == "string":
            items = _literal(token.text)
            if items is not None:
                values = tuple(workbook._redacted(items).split(","))
                return {"source": "inline", "resolved": True, "count": len(values),
                        "values": values}
        if token.kind == "ref":
            if token.book:
                return {"source": "external", "resolved": False, "reason": "external"}
            return _range(workbook, sheet, token.text, "range")
        if token.kind == "name":
            return _name(workbook, sheet, token.text)
        if token.kind == "structured":
            return _table(workbook, token.text)
    return _dynamic(found)


def _literal(text: str) -> str | None:
    """A string literal's text (``""`` a quote inside); ``None`` when it is not one."""
    if len(text) < 2 or text[0] != '"' or text[-1] != '"':
        return None
    return text[1:-1].replace('""', '"')


def _dynamic(found: list[Token]) -> dict[str, Any]:
    names = {t.text.upper().removeprefix("_XLFN.") for t in found if t.kind == "function"}
    source = "indirect" if "INDIRECT" in names else "offset" if "OFFSET" in names else \
        "formula"
    if any(t.kind == "ref" and t.book for t in found):
        source = "external"
    uses = tuple(dict.fromkeys(t.text for t in found if t.kind in ("ref", "name",
                                                                     "structured")))
    return {"source": source, "resolved": False, "reason": source, "uses": uses}


def _range(workbook: "Workbook", sheet: str, text: str, source: str,
           name: str | None = None) -> dict[str, Any]:
    sheets = sheet_names(workbook._every_sheet)
    target, reason = range_of(text if "!" in text else f"{quote_sheet(sheet)}!{text}", sheets)
    if target is not None and workbook._sheet_withheld(target.sheet.name):
        target, reason = None, "withheld"
    if target is None:
        out = {"source": source, "resolved": False, "reason": reason or "formula"}
        if name is not None:
            out["name"] = name
        return out
    return _read(workbook, target, source, name)


def _read(workbook: "Workbook", target: RangeAddress, source: str,
          name: str | None) -> dict[str, Any]:
    from .profile import ALLOWED_CAP, _withheld

    out: dict[str, Any] = {"source": source, "range": target.text}
    if name is not None:
        out["name"] = name
    if _withheld(workbook, target):
        return {**out, "resolved": False, "reason": "withheld"}
    values: list = []
    cut = False
    for row in workbook.read(target):
        for value in row:
            if value.t == "blank":
                continue
            if len(values) >= ALLOWED_CAP:
                cut = True
                break
            values.append(value.v)
        if cut:
            break
    return {**out, "resolved": True, "count": len(values), "cut": cut,
            "values": tuple(values)}


def _name(workbook: "Workbook", sheet: str, text: str) -> dict[str, Any]:
    """A list's source given as a defined name, followed (see the module's notes)."""
    first = None
    for _ in range(NAME_DEPTH):
        defined = _lookup(workbook, sheet, text)
        if defined is None:
            return {"source": "name", "resolved": False, "reason": "unknown_name",
                    **({"name": first} if first else {})}
        first = first or defined.name
        if defined.range is not None:
            if workbook._sheet_withheld(defined.range.sheet.name):
                return {"source": "name", "resolved": False, "reason": "withheld",
                        "name": first}
            return _read(workbook, defined.range, "name", first)
        body = (defined.formula or "").strip().removeprefix("=").strip()
        try:
            found = [t for t in tokens(body) if not (t.kind == "other" and not t.text.strip())]
        except ShiftError:
            found = []
        if len(found) == 1 and found[0].kind == "name":
            text = found[0].text
            if defined.scope is not None:
                sheet = defined.scope
            continue
        if len(found) == 1 and found[0].kind == "structured":
            return {**_table(workbook, found[0].text), "source": "name", "name": first}
        reason = defined.reason or "formula"
        if reason == "formula":
            reason = _dynamic(found)["reason"] if found else "formula"
        return {"source": "name", "resolved": False, "reason": reason, "name": first}
    return {"source": "name", "resolved": False, "reason": "formula", "name": first}


def _lookup(workbook: "Workbook", sheet: str, text: str) -> Any:
    folded = text.casefold()
    local = [n for n in workbook.names if n.scope is not None
             and n.scope.casefold() == sheet.casefold() and n.name.casefold() == folded]
    wide = [n for n in workbook.names if n.scope is None and n.bad_scope is None
            and n.name.casefold() == folded]
    found = local or wide
    return found[0] if len(found) == 1 else None


def _table(workbook: "Workbook", text: str) -> dict[str, Any]:
    """A structured reference to one table column's data cells, ``Rates[Category]``."""
    table, _, rest = text.partition("[")
    column = rest[:-1] if rest.endswith("]") else None
    if not table or column is None or any(ch in column for ch in "[]#,:'"):
        return {"source": "table", "resolved": False, "reason": "formula"}
    region = next((r for r in workbook.tables if r.name and
                   r.name.casefold() == table.strip().casefold()), None)
    if region is None:
        return {"source": "table", "resolved": False, "reason": "unknown_table"}
    try:
        offset = region.column_index(column.strip())
    except Exception:  # NameNotFound: no such column
        return {"source": "table", "resolved": False, "reason": "unknown_table"}
    data = region.data
    if data is None:
        return {"source": "table", "resolved": True, "count": 0, "values": ()}
    left = region.area.first.column + offset
    target = RangeAddress(region.sheet_ref, Area(Cell(data.first.row, left),
                                                 Cell(data.last.row, left)))
    return _read(workbook, target, "table", None)


# -- a value written, checked ------------------------------------------------------------------

#: What each operator asks of a value ``x`` and the bounds ``a`` (and ``b``).
_OPERATORS = {
    "between": lambda x, a, b: a <= x <= b,
    "notBetween": lambda x, a, b: not a <= x <= b,
    "equal": lambda x, a, b: x == a,
    "notEqual": lambda x, a, b: x != a,
    "greaterThan": lambda x, a, b: x > a,
    "lessThan": lambda x, a, b: x < a,
    "greaterThanOrEqual": lambda x, a, b: x >= a,
    "lessThanOrEqual": lambda x, a, b: x <= a,
}
#: The operators that take a second bound.
_TWO_BOUNDS = ("between", "notBetween")


def check(workbook: "Workbook", rules: "list[ValidationRule]", cell: Cell, kind: str,
          value: Any, serial: "float | None" = None,
          memo: "dict | None" = None) -> str | None:
    """Whether a value written into ``cell`` satisfies the first of ``rules`` (one sheet's)
    that covers it (see the module's notes): ``"invalid"``, ``"unchecked"``, or ``None``
    (satisfied, no rule, or a blank).  ``kind`` and ``value`` as the cell writer plans
    them: ``text`` (a string or rich text), ``number``, ``bool``, ``date`` (a ``date``,
    with its ``serial`` in the workbook's date system), ``formula`` or ``blank``.  ``memo``
    keeps the bounds read, for a caller checking many cells."""
    rule = next((r for r in rules if r.covers(cell)), None)
    if rule is None or kind == "blank" or rule.type in ("none", "any"):
        return None
    if kind == "formula":
        return "unchecked"
    text = getattr(value, "text", value) if kind == "text" else None
    if rule.type == "list":
        if not rule.resolved or rule.values is None or rule.cut:
            return "unchecked"
        listed = _listed(kind, text if kind == "text" else value, rule.values)
        return None if listed else "invalid"
    if rule.type not in ("whole", "decimal", "date", "textLength"):
        return "unchecked"
    if rule.type == "textLength":
        if kind != "text":
            return "unchecked"
        measured: float = len(text)
    else:
        if kind in ("bool", "text"):
            return "invalid"
        measured = float(serial if kind == "date" else value)
        if rule.type == "whole" and not measured.is_integer():
            return "invalid"
    operator = rule.operator or "between"
    test = _OPERATORS.get(operator)
    if test is None:
        return "unchecked"
    memo = {} if memo is None else memo
    key = (rule.sheet, rule.index)
    if key not in memo:
        memo[key] = (_bound(workbook, rule, rule.formula1),
                     _bound(workbook, rule, rule.formula2) if operator in _TWO_BOUNDS else 0.0)
    low, high = memo[key]
    if low is None or high is None:
        return "unchecked"
    return None if test(measured, low, high) else "invalid"


def _listed(kind: str, value: Any, allowed: tuple) -> bool:
    for item in allowed:
        if kind == "text":
            if isinstance(item, str) and item.strip().casefold() == value.strip().casefold():
                return True
        elif kind == "bool":
            if isinstance(item, bool) and item == value:
                return True
            if isinstance(item, str) and item.strip().upper() == ("TRUE" if value else "FALSE"):
                return True
        elif kind == "date":
            if isinstance(item, str) and item.strip() == value.isoformat():
                return True
        elif not isinstance(item, bool):
            try:
                if float(item) == float(value):
                    return True
            except (TypeError, ValueError):
                continue
    return False


def _bound(workbook: "Workbook", rule: ValidationRule, formula: str | None) -> float | None:
    """A rule's bound as a number: a number written, or one cell holding a number and no
    formula; ``None`` for anything else."""
    text = (formula or "").strip().removeprefix("=").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        pass
    try:
        found = [t for t in tokens(text) if not (t.kind == "other" and not t.text.strip())]
    except ShiftError:
        return None
    if len(found) != 1 or found[0].kind != "ref" or found[0].book or ":" in found[0].body:
        return None
    sheets = sheet_names(workbook._every_sheet)
    reference = found[0].text if found[0].sheet else f"{rule.sheet}!{found[0].text}"
    target, _ = range_of(reference, sheets)
    if target is None or workbook._sheet_withheld(target.sheet.name):
        return None
    value = workbook.cell(target)
    if value.formula or isinstance(value.v, bool):
        return None
    if value.t in ("number", "percent"):
        return float(value.v)
    return float(value.serial) if value.serial is not None else None


__all__ = ["CELLS_SHOWN", "COMPARING", "NAME_DEPTH", "SOURCES", "ValidationRule", "check",
           "inspect"]
