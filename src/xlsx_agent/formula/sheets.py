"""Sheets in formula references (E2's rename, E5): found, renamed, and deleted as Excel does.

A reference names its sheet in a prefix: ``Stock!``, ``'Tax rates'!``, a run of sheets
``Stock:Lists!`` (a 3-D reference, ECMA-376 Part 1, §18.17.2.3), quoted as a whole when
either end needs quotes (``'Stock:Tax rates'!``, ``'Stock:a''b'!``), or a prefix before a
name local to a sheet (``Lists!Local``) or before ``#REF!``.  A prefix with a workbook
index (``[1]Lists!``) names another workbook's sheet and is never touched::

    rename("VLOOKUP(C10,Lists!$A$2:$B$5,2)", "Lists", "Tax rates")
        # "VLOOKUP(C10,'Tax rates'!$A$2:$B$5,2)"
    delete("SUM(Stock:Lists!B2)+Lists!A1", "Lists", ["Stock", "Lists", "Totals"])
        # "SUM(Stock!B2)+#REF!"

**What Excel does, measured** (Excel for Mac 16.106, 2026-11-08; made-up workbooks whose
formulas, defined names, validations, conditional formats, chart series and a pivot cache
refer to a sheet, the sheet renamed or deleted through Excel's object model and the file
saved):

* *Renamed*: every reference's prefix is written with the new name, quoted as an address
  quotes it (:func:`xlsx_agent.address.needs_quotes`), a 3-D run quoted as a whole when an
  end needs quotes.  Text in a string literal (``INDIRECT("Lists!B2")``) is left as it is
  -- it is no reference -- and so is a hyperlink's location (the caller's to report).
* *Deleted*: a reference to the sheet becomes ``#REF!``, its prefix and cells together
  (``VLOOKUP(C10,#REF!,2,FALSE)``, ``COUNTA(#REF!)``, ``OFFSET(#REF!,...)``); a 3-D run
  with the sheet at one end shrinks to the next sheet inward (``Stock:Lists!B2`` deleting
  ``Lists`` is ``Stock!B2``; ``Stock:Totals!`` deleting ``Stock`` is ``Lists:Totals!``),
  and one with the sheet inside it stays as it is.
* Sheet names are matched as Excel compares them (:func:`fold_name`).

A structured reference to a table (``Rates[Tax]``) becomes ``#REF!`` when the table goes
(:func:`drop_tables`; measured: ``SUM(Rates[Tax])`` is ``SUM(#REF!)`` once the sheet
holding ``Rates`` is deleted).  Text the lexer cannot read is left as it is and reported by
the callers, never guessed.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import Iterable, Sequence

from ..address import needs_quotes
from . import ShiftError, Token, tokens

#: Characters Excel leaves out when it compares sheet names (measured with the oracle,
#: Excel for Mac 16.106: ``a<SOFT HYPHEN>b`` and ``ab`` are one name, and so with a
#: byte-order mark or a zero-width space).
_IGNORED = dict.fromkeys(map(ord, "\u00ad\ufeff\u200b"))


def fold_name(name: str) -> str:
    """``name`` as Excel compares sheet names, as far as measured (Excel for Mac 16.106,
    2026-10-27/28, the oracle's ``same_name``): case aside, compatibility forms as their
    plain ones (NFKC: a superscript and a fullwidth digit, a no-break space and a space),
    a soft hyphen, a byte-order mark and a zero-width space left out."""
    return unicodedata.normalize("NFKC", name.translate(_IGNORED)).casefold()


@dataclass(frozen=True)
class Prefix:
    """A reference's sheet prefix read: its first and last sheet (``last`` ``None`` for
    one sheet), and whether it names another workbook."""

    first: str
    last: str | None = None
    book: bool = False


def read_prefix(text: str) -> Prefix | None:
    """A sheet prefix (``Stock!``, ``'Q3 sales'!``, ``Jan:Mar!``, ``'[1]Prices'!``) as a
    :class:`Prefix`; ``None`` for text that is none."""
    if not text.endswith("!"):
        return None
    body = text[:-1]
    book = False
    if body.startswith("["):
        book = True
    if body.startswith("'") and body.endswith("'") and len(body) >= 2:
        body = body[1:-1].replace("''", "'")
    if body.startswith("["):
        book = True
        close = body.find("]")
        body = body[close + 1:] if close >= 0 else body
    if not body:
        return None
    first, sep, last = body.partition(":")
    return Prefix(first, last if sep else None, book)


def spell_prefix(first: str, last: str | None = None) -> str:
    """The prefix Excel writes for a sheet, or a run of sheets (``'Stock:Tax rates'!``)."""
    names = [first] if last is None else [first, last]
    body = ":".join(names)
    if any(needs_quotes(name) for name in names):
        return "'" + body.replace("'", "''") + "'!"
    return body + "!"


def _prefixed(found: list[Token]) -> Iterable[tuple[int, Token, Prefix]]:
    """The tokens that carry a sheet prefix of this workbook: references, and a prefix
    alone (before a sheet's own name or ``#REF!``)."""
    for index, token in enumerate(found):
        if token.kind == "ref" and token.sheet:
            prefix = read_prefix(token.sheet)
        elif token.kind == "other" and token.text.endswith("!") and len(token.text) > 1:
            prefix = read_prefix(token.text)
        else:
            continue
        if prefix is not None and not prefix.book:
            yield index, token, prefix


def refers_to(text: str, name: str, order: Sequence[str] = ()) -> bool:
    """Whether a formula refers to the sheet ``name``: a reference on it, or a 3-D run
    with it at an end or (``order``, every sheet's name in the workbook's order, given)
    inside.  Unreadable text: ``False``."""
    try:
        found = tokens(text)
    except ShiftError:
        return False
    key = fold_name(name)
    folded = [fold_name(n) for n in order]
    for _, _, prefix in _prefixed(found):
        ends = [fold_name(prefix.first)] + ([fold_name(prefix.last)] if prefix.last else [])
        if key in ends:
            return True
        if prefix.last and key in folded and ends[0] in folded and ends[1] in folded:
            a, b = sorted((folded.index(ends[0]), folded.index(ends[1])))
            if a < folded.index(key) < b:
                return True
    return False


def rename(text: str, old: str, new: str) -> str | None:
    """``text`` with every reference to the sheet ``old`` naming ``new`` instead (see the
    module); ``None`` when nothing in it names ``old`` (or it cannot be read)."""
    try:
        found = tokens(text)
    except ShiftError:
        return None
    key = fold_name(old)
    out = [token.text for token in found]
    changed = False
    for index, token, prefix in _prefixed(found):
        hit = fold_name(prefix.first) == key
        first = new if hit else prefix.first
        last = prefix.last
        if last is not None and fold_name(last) == key:
            last, hit = new, True
        if not hit:
            continue
        spelled = spell_prefix(first, last)
        if spelled == (token.sheet if token.kind == "ref" else token.text):
            continue
        out[index] = spelled + (token.body if token.kind == "ref" else "")
        changed = True
    return "".join(out) if changed else None


def delete(text: str, name: str, order: Sequence[str]) -> str | None:
    """``text`` with the sheet ``name`` deleted from it as Excel deletes it (see the
    module): a reference on it ``#REF!``, a 3-D run ending at it shrunk inward; ``order``
    is every sheet's name in the workbook's order before the deletion.  ``None`` when
    nothing changes (or the text cannot be read)."""
    try:
        found = tokens(text)
    except ShiftError:
        return None
    key = fold_name(name)
    folded = [fold_name(n) for n in order]
    out = [token.text for token in found]
    changed = False
    for index, token, prefix in _prefixed(found):
        ends = [prefix.first] + ([prefix.last] if prefix.last else [])
        keys = [fold_name(end) for end in ends]
        if key not in keys:
            continue
        changed = True
        if prefix.last is not None and keys[0] != keys[1] and all(k in folded for k in keys):
            # A run: the deleted end moves one sheet inward, towards the other end.
            here, there = (0, 1) if keys[0] == key else (1, 0)
            step = 1 if folded.index(keys[there]) > folded.index(keys[here]) else -1
            inward = order[folded.index(keys[here]) + step]
            ends[here] = inward
            first, last = ends
            single = fold_name(first) == fold_name(last)
            spelled = spell_prefix(first) if single else spell_prefix(first, last)
            out[index] = spelled + (token.body if token.kind == "ref" else "")
            continue
        out[index] = "#REF!"
        if token.kind == "other":
            # A prefix alone: the name or #REF! after it goes with it.
            nxt = index + 1
            if nxt < len(found) and found[nxt].kind in ("name", "error"):
                out[nxt] = ""
    return "".join(out) if changed else None


def tables_named(text: str, tables: "set[str]") -> bool:
    """Whether a formula holds a structured reference to one of ``tables`` (names case
    folded)."""
    try:
        found = tokens(text)
    except ShiftError:
        return False
    return any(_table_of(token) in tables for token in found if token.kind == "structured")


def drop_tables(text: str, tables: "set[str]") -> str | None:
    """``text`` with each structured reference to one of ``tables`` (names case folded)
    as ``#REF!``, as Excel writes it when the table goes; ``None`` when none."""
    try:
        found = tokens(text)
    except ShiftError:
        return None
    out = []
    changed = False
    for token in found:
        if token.kind == "structured" and _table_of(token) in tables:
            out.append("#REF!")
            changed = True
        else:
            out.append(token.text)
    return "".join(out) if changed else None


def _table_of(token: Token) -> str | None:
    """The table a structured reference names (case folded), ``None`` for one without
    (``[@Price]``, inside its own table)."""
    head = token.text.split("[", 1)[0].strip()
    return head.casefold() if head else None


__all__ = ["Prefix", "delete", "drop_tables", "fold_name", "read_prefix", "refers_to",
           "rename", "spell_prefix", "tables_named"]
