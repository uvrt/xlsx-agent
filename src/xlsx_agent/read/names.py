"""Defined names: the workbook's ``definedNames``, each read as a range or said not to be one.

A defined name (ECMA-376 Part 1, §18.2.5) is a name for a cell, a range, a constant or a
formula, workbook-wide or local to one sheet (``localSheetId``: the sheet's position in the
workbook's ``sheets``, counted from 0).  A name over one rectangle on one sheet is usable as
a table (``t:Prices``, or ``t:Stock!Prices`` when it is local to ``Stock``); any other is
listed with the reason it is not a range, and never evaluated (ROADMAP.md, "Messy
workbooks": "defined names that are formulas, not ranges ... listed, not resolved")::

    Stock!$B$2:$B$9            a range
    'Q3 sales'!$A:$A           whole columns, clipped to the used range when read
    0.21, SUM(Stock!B2:B9)     "formula": a constant or a formula, never evaluated
    Stock!$A$1,Stock!$C$1      "formula": more than one area
    Sheet1:Sheet3!$A$1         "formula": a reference over several sheets
    [1]Stock!$A$1              "external": another workbook's cells, never followed
    #REF!, Stock!#REF!         "ref_error": what the name pointed at was deleted
    Gone!$A$1                  "unknown_sheet": a sheet the workbook does not have

A name's formula text is kept on :attr:`DefinedName.formula` for a program and never put
in a message: formulas hold literals (ROADMAP.md, "Design rules", 7).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Sequence

from ..address import AddressError, RangeAddress, SheetRef, TableAddress, is_address_name, \
    parse_address, quote_sheet
from .strings import decode_xstring

#: Why a defined name is not a range.
NOT_A_RANGE = ("formula", "external", "ref_error", "unknown_sheet")

#: The prefix ECMA-376 reserves for the built-in names (``_xlnm.Print_Area``...; §18.2.5).
BUILTIN_PREFIX = "_xlnm."

# What may not stand in a bare sheet name in a formula: operators, separators, quotes.
_NOT_IN_FORMULA_SHEET = frozenset(" ()+-*/^&=<>,;\"{}%:[]'!")


@dataclass(frozen=True)
class DefinedName:
    """One defined name.

    ``name`` as the workbook spells it; ``scope`` the sheet it is local to (its name), or
    ``None`` for a workbook-wide name; ``hidden`` the name is hidden in Excel's interface;
    ``builtin`` it is one of the names ECMA-376 reserves (``_xlnm.``).  ``range`` the cells
    it names, when it names one rectangle on one sheet; else ``None`` and ``reason`` says
    why (:data:`NOT_A_RANGE`).  ``formula`` is its text, for a program only (never in a
    message)."""

    name: str
    scope: str | None = None
    hidden: bool = False
    range: RangeAddress | None = None
    reason: str | None = None
    builtin: bool = False
    formula: str = field(default="", repr=False)
    #: The sheet's index as the file gives it (``localSheetId``), when it names no sheet.
    bad_scope: str | None = None

    @property
    def is_range(self) -> bool:
        return self.range is not None

    @property
    def address(self) -> TableAddress | None:
        """``t:Prices`` or ``t:Stock!Prices``; ``None`` for a name an address cannot hold
        (see :mod:`xlsx_agent.address`) or one whose scope names no sheet."""
        if not is_address_name(self.name) or self.bad_scope is not None:
            return None
        scope = SheetRef(name=self.scope) if self.scope else None
        return TableAddress(name=self.name, scope=scope)


def read_names(holder, sheets: Sequence) -> list[DefinedName]:
    """The ``definedName`` elements under ``holder`` (the workbook's ``definedNames``, or
    ``None``), in the file's order; ``sheets`` the workbook's sheets, in order."""
    if holder is None:
        return []
    names = [sheet.name for sheet in sheets]
    folded = sheet_names(sheets)
    out = []
    for node in holder:
        if not isinstance(node.tag, str) or node.tag.rpartition("}")[2] != "definedName":
            continue
        name = decode_xstring(node.get("name") or "")
        if not name:
            continue
        scope = bad_scope = None
        local = node.get("localSheetId")
        if local is not None:
            index = _int(local)
            if index is not None and 0 <= index < len(names) and names[index]:
                scope = names[index]
            else:
                bad_scope = local
        formula = node.text or ""
        area, reason = range_of(formula, folded)
        out.append(DefinedName(name=name, scope=scope, hidden=_true(node.get("hidden")),
                               range=area, reason=reason,
                               builtin=name.casefold().startswith(BUILTIN_PREFIX.casefold()),
                               formula=formula, bad_scope=bad_scope))
    return out


def sheet_names(sheets: Sequence) -> dict[str, str]:
    """Each sheet's name, case folded, to its name; a sheet whose name a redaction changed
    (:mod:`xlsx_agent.redact`) by its stored name too, which formulas hold."""
    out = {}
    for sheet in sheets:
        out[sheet.name.casefold()] = sheet.name
        stored = getattr(sheet, "stored", None)
        if stored:
            out.setdefault(stored.casefold(), sheet.name)
    return out


def range_of(formula: str, sheets: dict[str, str]) -> tuple[RangeAddress | None, str | None]:
    """A name's formula as ``(range, None)`` when it is one rectangle on one sheet, else
    ``(None, reason)``.  ``sheets`` maps each sheet's name, case folded, to its name."""
    text = formula.strip()
    if text.startswith("="):
        text = text[1:].lstrip()
    if "#REF!" in text.upper():
        return None, "ref_error"
    if text.startswith("[") or text.startswith("'["):
        return None, "external"
    sheet, rest = _sheet_and_rest(text)
    if sheet is None:
        return None, "formula"
    if sheet.startswith("[") or (sheet[:1] != "'" and "[" in sheet):
        return None, "external"
    found = sheets.get(sheet.casefold())
    if not rest or not all(ch.isalnum() or ch in "$:" for ch in rest):
        return None, "formula"
    try:
        address = parse_address(quote_sheet(found or sheet) + "!" + rest)
    except (AddressError, ValueError):
        return None, "formula"
    if found is None:
        return None, "unknown_sheet"
    return address, None


def _sheet_and_rest(text: str) -> tuple[str | None, str]:
    """A formula's sheet (unquoted) and what follows its ``!``; ``(None, "")`` when it does
    not start with a sheet and a ``!``."""
    if text.startswith("'"):
        parts, i = [], 1
        while True:
            j = text.find("'", i)
            if j < 0:
                return None, ""
            parts.append(text[i:j])
            if text.startswith("''", j):
                parts.append("'")
                i = j + 2
                continue
            i = j + 1
            break
        if not text.startswith("!", i):
            return None, ""
        name = "".join(parts)
        if not name or name.startswith("["):
            return (name, text[i + 1:]) if name.startswith("[") else (None, "")
        return name, text[i + 1:]
    bang = text.find("!")
    if bang <= 0:
        return None, ""
    name = text[:bang]
    if name.startswith("["):
        return name, text[bang + 1:]
    if any(ch in _NOT_IN_FORMULA_SHEET or ch.isspace() for ch in name):
        return None, ""
    return name, text[bang + 1:]


def _true(value: str | None) -> bool:
    return value is not None and value.strip() in ("1", "true")


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None
