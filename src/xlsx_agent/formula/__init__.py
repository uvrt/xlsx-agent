"""Formula references (0.2, a part of E2): the references in a formula's text, found and
moved, nothing evaluated.

A cell stores its formula as text in the A1 style, without the ``=`` (ECMA-376 Part 1,
§18.17, "Formulas").  This module reads that text into tokens, as far as references need:
string literals, quoted and unquoted sheet names, another workbook's index (``[1]``),
cell and range references with their ``$`` anchors, whole columns (``A:C``) and whole rows
(``3:5``), structured references (``Stock[Price]``, ``Stock[[#This Row],[Price]]``),
error literals, numbers, function names and defined names::

    tokens("SUM(Stock!$D$10:D14)*Markup")   # ..., Token("ref", "Stock!$D$10:D14"), ...
    shift("D10*E10", 1, 0)                  # "D11*E11": a shared formula's next member
    shift("$D$10+D10", 2, 1)                # "$D$10+E12": anchored parts stay
    is_external("[1]Prices!$B$2*2")         # True: another workbook's cells

**What moves** (:func:`shift`): the relative parts of cell, range, column and row
references -- what Excel moves when a formula is filled or a shared formula's member
computed from its master -- on any sheet, of this workbook or another.  Left as they are:
anchored parts (``$``), structured references (their rows are the formula's own row
anyway), defined names, string literals (an ``INDIRECT("A1")`` text is not a reference)
and errors.  A reference that would move off the grid, and text the lexer cannot read --
an unclosed string, quote or bracket -- are refused (:class:`ShiftError`), never guessed:
the caller's edit is then refused before anything changes.

**What is the same formula** (:func:`normalise`, X5's formula-aware compare): two texts
that differ only where the grammar gives them no meaning are one formula.  Case: defined
names "are not case-sensitive" (ECMA-376 Part 1, §18.17, "Names"), nor are function names,
error literals, column letters and sheet names as this lexer reads them (``column_index``
reads letters case aside, :data:`ERRORS` are matched case aside, sheets are matched case
aside as Excel matches them), so those tokens are compared in capitals; string literals,
structured references and whatever the lexer leaves as it is keep their case.  Whitespace:
outside a string literal, a space between two operands is the intersection operator
(``A1:B5 B2:C3``, the space operator of §18.17's grammar) and is kept, as one space; any
other run of spaces, tabs and line breaks -- around an operator, a parenthesis or a comma
-- only separates tokens and is dropped.  A leading ``=`` is not part of what a cell stores
(§18.17).  Text the lexer cannot read is compared exactly.

This is the subset E3's shared-formula split needs; E2 adds renaming sheets in references
and the rest of the grammar on this lexer.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterator, TypeGuard

from ooxml_edit.sheets import MAX_COLUMN, MAX_ROW, column_index, column_letters

#: Every :attr:`Token.kind`.
KINDS = ("ref", "name", "function", "string", "error", "number", "structured", "other")

#: The error literals of ECMA-376 Part 1, §18.17.2.8 and the ones later versions write.
ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
          "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!",
          "#BUSY!", "#UNKNOWN!", "#PYTHON!")

_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})")
_COLUMN = re.compile(r"(\$?)([A-Za-z]{1,3})")
_ROW = re.compile(r"(\$?)([0-9]{1,7})")
_NUMBER = re.compile(r"[0-9]+(?:\.[0-9]*)?(?:[eE][+-]?[0-9]+)?|\.[0-9]+(?:[eE][+-]?[0-9]+)?")
# A run of what a name, a function, a sheet or a reference is spelled with: every character
# from U+00A1 on, as Excel writes bare sheet names of them (x², a§b, aµb; measured, 0.1.6:
# xlsx_agent.address.needs_quotes); the ones it quotes never stand bare in its formulas.
_WORD = re.compile(r"[A-Za-z0-9_.\\?$\u00a1-\U0010ffff]+")
_BOOK = re.compile(r"\[[0-9]+\]")
_DIGITS = frozenset("0123456789")


class ShiftError(ValueError):
    """A formula's references could not be moved safely: a reference off the grid, or text
    the lexer cannot read.  The message names the problem, never the formula's text."""


@dataclass(frozen=True)
class Token:
    """One piece of a formula's text: ``kind`` (:data:`KINDS`), ``text`` as written, and
    where it starts.  A ``ref`` has its ``sheet`` prefix as written (``Stock!``,
    ``'Q3 sales'!``, ``[1]Prices!``, ``""`` for none) and its ``body`` (``$D$10:D14``),
    and ``book`` says it names another workbook."""

    kind: str
    text: str
    start: int
    sheet: str = ""
    book: bool = False

    @property
    def body(self) -> str:
        return self.text[len(self.sheet):]


def tokens(text: str) -> list[Token]:
    """``text`` (a formula as a cell stores it, with or without its ``=``) as tokens (see
    the module's notes); :class:`ShiftError` for an unclosed string, quote or bracket.
    Joining every token's text gives ``text`` back."""
    return list(_lex(text))


def _lex(text: str) -> Iterator[Token]:
    i, n = 0, len(text)
    while i < n:
        char = text[i]
        if char == '"':
            end = _closing(text, i, '"')
            yield Token("string", text[i:end], i)
            i = end
        elif char == "#":
            error = next((e for e in ERRORS if text.startswith(e, i)
                          or text[i:i + len(e)].upper() == e), None)
            if error is None:            # a spill reference's # (A1#), or text we leave
                yield Token("other", char, i)
                i += 1
            else:
                yield Token("error", text[i:i + len(error)], i)
                i += len(error)
        elif char == "'":
            end = _closing(text, i, "'")
            if end < n and text[end] == "!":
                yield from _reference(text, i, end + 1)
                i = _reference_end(text, end + 1)
            else:                        # not a sheet: leave it as it is
                yield Token("other", text[i:end], i)
                i = end
        elif char == "[":
            book = _BOOK.match(text, i)
            if book is not None and _starts_token(text, i):
                after = book.end()
                word = _WORD.match(text, after)
                if after < n and text[after] == "!":           # [1]!Name
                    yield Token("other", text[i:after + 1], i)
                    i = after + 1
                    continue
                if after < n and text[after] == "'":
                    end = _closing(text, after, "'")
                    if end < n and text[end] == "!":
                        yield from _reference(text, i, end + 1, book=True)
                        i = _reference_end(text, end + 1)
                        continue
                if word is not None and word.end() < n and text[word.end()] == "!":
                    yield from _reference(text, i, word.end() + 1, book=True)
                    i = _reference_end(text, word.end() + 1)
                    continue
            end = _brackets(text, i)
            yield Token("structured", text[i:end], i)
            i = end
        elif char in _DIGITS or (char == "." and i + 1 < n and text[i + 1] in _DIGITS):
            row = _rows(text, i)
            if row is not None:
                yield Token("ref", text[i:row], i)
                i = row
                continue
            number = _NUMBER.match(text, i)
            assert number is not None   # a digit, or . and a digit, starts a number
            yield Token("number", number.group(0), i)
            i = number.end()
        elif char == "$" or char.isalpha() or char in "_\\" or ord(char) >= 0xA1:
            word = _WORD.match(text, i)
            assert word is not None   # a letter, $, _ or \ starts a word
            end = word.end()

            # A sheet range (Jan:Mar!A1) or a sheet (Stock!A1)
            if end < n and text[end] == "!":
                yield from _reference(text, i, end + 1)
                i = _reference_end(text, end + 1)
                continue
            if end < n and text[end] == ":":
                other = _WORD.match(text, end + 1)
                if other is not None and other.end() < len(text) \
                        and text[other.end()] == "!":
                    yield from _reference(text, i, other.end() + 1)
                    i = _reference_end(text, other.end() + 1)
                    continue
            span = _span(text, i)
            if span is not None:
                yield Token("ref", text[i:span], i)
                i = span
                continue
            rest = text[end:].lstrip(" ")
            if rest.startswith("("):
                yield Token("function", word.group(0), i)
            elif rest.startswith("["):
                close = _brackets(text, end + len(text[end:]) - len(rest))
                yield Token("structured", text[i:close], i)
                i = close
                continue
            else:
                yield Token("name", word.group(0), i)
            i = end
        else:
            yield Token("other", char, i)
            i += 1


def _starts_token(text: str, i: int) -> bool:
    """Whether position ``i`` starts a token: not inside a name (a structured reference's
    ``[`` follows its table's name)."""
    if i == 0:
        return True
    before = text[i - 1]
    return not (before.isalnum() or before in "_.\\]?")


def _closing(text: str, i: int, quote: str) -> int:
    """The position after the quote that closes the one at ``i`` (a doubled quote is one
    inside)."""
    j = i + 1
    while True:
        j = text.find(quote, j)
        if j < 0:
            raise ShiftError("a string or sheet name is not closed")
        if j + 1 < len(text) and text[j + 1] == quote:
            j += 2
            continue
        return j + 1


def _brackets(text: str, i: int) -> int:
    """The position after the ``]`` that closes the ``[`` at ``i``; inside, ``'`` escapes
    the next character, as structured references write ``'[``."""
    depth, j, n = 0, i, len(text)
    while j < n:
        char = text[j]
        if char == "'":
            j += 2
            continue
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth == 0:
                return j + 1
        j += 1
    raise ShiftError("a bracket is not closed")


def _cell(match: "re.Match[str] | None") -> "TypeGuard[re.Match[str]]":
    return match is not None and column_index(match.group(2)) <= MAX_COLUMN \
        and 1 <= int(match.group(4)) <= MAX_ROW


def _boundary(text: str, j: int) -> bool:
    """Whether a reference may end at ``j``: no letter, digit, ``_``, ``.``, ``(`` or ``[``
    follows (a name or a function would go on)."""
    if j >= len(text):
        return True
    return not (text[j].isalnum() or text[j] in "_.\\?([$")


def _span(text: str, i: int) -> int | None:
    """The end of a cell, range or whole-column reference at ``i``, or ``None``."""
    cell = _CELL.match(text, i)
    if _cell(cell):
        end = cell.end()
        if end < len(text) and text[end] == ":":
            other = _CELL.match(text, end + 1)
            if _cell(other) and _boundary(text, other.end()):
                return other.end()
        if _boundary(text, end):
            return end
        return None
    column = _COLUMN.match(text, i)
    if column is not None and column_index(column.group(2)) <= MAX_COLUMN \
            and column.end() < len(text) and text[column.end()] == ":":
        other = _COLUMN.match(text, column.end() + 1)
        if other is not None and column_index(other.group(2)) <= MAX_COLUMN \
                and _boundary(text, other.end()) and not _CELL.match(text, column.end() + 1):
            return other.end()
    return None


def _rows(text: str, i: int) -> int | None:
    """The end of a whole-row reference (``3:5``, ``$3:$5``) at ``i``, or ``None``."""
    first = _ROW.match(text, i)
    if first is None or first.end() >= len(text) or text[first.end()] != ":":
        return None
    other = _ROW.match(text, first.end() + 1)
    if other is None or not _boundary(text, other.end()):
        return None
    if not (1 <= int(first.group(2)) <= MAX_ROW and 1 <= int(other.group(2)) <= MAX_ROW):
        return None
    return other.end()


def _reference(text: str, start: int, body: int, *, book: bool = False) -> Iterator[Token]:
    """The reference whose sheet prefix runs from ``start`` to ``body``."""
    end = _reference_end(text, body)
    if end == body:                    # Stock!#REF! or a prefix before something else
        yield Token("other", text[start:body], start)
        return
    yield Token("ref", text[start:end], start, sheet=text[start:body],
                book=book or text[start:start + 1] == "[" or text[start + 1:start + 2] == "[")


def _reference_end(text: str, body: int) -> int:
    span = _span(text, body)
    if span is not None:
        return span
    rows = _rows(text, body) if body < len(text) and (text[body].isdigit()
                                                       or text[body] == "$") else None
    return rows if rows is not None else body


# -- what the tokens say -----------------------------------------------------------------------


def shift(text: str, rows: int, columns: int) -> str:
    """``text`` with the relative parts of its references moved by ``rows`` and ``columns``
    (see the module's notes); :class:`ShiftError` when one would leave the grid or the text
    cannot be read."""
    if not rows and not columns:
        return text
    out = []
    for token in tokens(text):
        out.append(token.sheet + _moved(token.body, rows, columns)
                   if token.kind == "ref" else token.text)
    return "".join(out)


def _moved(body: str, rows: int, columns: int) -> str:
    parts = body.split(":")
    return ":".join(_moved_part(part, rows, columns) for part in parts)


def _moved_part(part: str, rows: int, columns: int) -> str:
    cell = _CELL.fullmatch(part)
    if cell is not None:
        return (_column(cell.group(1), cell.group(2), columns)
                + _row(cell.group(3), cell.group(4), rows))
    column = _COLUMN.fullmatch(part)
    if column is not None:
        return _column(column.group(1), column.group(2), columns)
    row = _ROW.fullmatch(part)
    if row is not None:
        return _row(row.group(1), row.group(2), rows)
    raise ShiftError("a reference the lexer cannot move")


def _column(anchor: str, letters: str, columns: int) -> str:
    if anchor or not columns:
        return anchor + letters
    index = column_index(letters) + columns
    if not 1 <= index <= MAX_COLUMN:
        raise ShiftError("a reference would move off the sheet's columns")
    return column_letters(index)


def _row(anchor: str, digits: str, rows: int) -> str:
    if anchor or not rows:
        return anchor + digits
    number = int(digits) + rows
    if not 1 <= number <= MAX_ROW:
        raise ShiftError("a reference would move off the sheet's rows")
    return str(number)


def is_external(text: str, names: "frozenset[str] | set[str]" = frozenset()) -> bool:
    """Whether a formula refers to another workbook: a reference with a workbook index
    (``[1]Prices!B2``, ``'[1]Q3 sales'!B2``, ``[1]!Rate``), or one of ``names`` (defined
    names, case folded, that refer to another workbook).  Text the lexer cannot read is not
    called external."""
    if "[" not in text and not names:
        return False
    try:
        found = tokens(text)
    except ShiftError:
        return False
    for token in found:
        if token.kind == "ref" and token.book:
            return True
        if token.kind == "other" and _BOOK.match(token.text):
            return True
        if names and token.kind == "name" and token.text.casefold() in names:
            return True
    return False


# -- the same formula ---------------------------------------------------------------------------

#: The whitespace a formula's text may hold between tokens.
_SPACES = frozenset(" \t\r\n\xa0")
#: Tokens compared in capitals (see the module's notes).
_CASELESS = ("ref", "name", "function", "error", "number")
#: Tokens an operand ends or starts with: a space between two is the intersection operator.
_OPERAND = ("ref", "name", "structured", "number", "string", "error")


def normalise(text: str) -> str:
    """``text`` (a formula, with or without its ``=``) as it is compared (see the module's
    notes): the case and whitespace the grammar gives no meaning left out.  Text the lexer
    cannot read is given back as it is, less its ``=``."""
    body = text[1:] if text.startswith("=") else text
    try:
        found = tokens(body)
    except ShiftError:
        return body
    out: list[str] = []
    last: Token | None = None
    space = False
    for token in found:
        if token.kind == "other" and token.text in _SPACES:
            space = True
            continue
        if space and last is not None and _ends_operand(last) and _starts_operand(token):
            out.append(" ")
        space = False
        out.append(token.text.upper() if token.kind in _CASELESS else token.text)
        last = token
    return "".join(out)


def _ends_operand(token: Token) -> bool:
    return token.kind in _OPERAND or (token.kind == "other" and token.text in (")", "}"))


def _starts_operand(token: Token) -> bool:
    return token.kind in _OPERAND or token.kind == "function" \
        or (token.kind == "other" and token.text in ("(", "{"))


__all__ = ["ERRORS", "KINDS", "ShiftError", "Token", "is_external", "normalise", "shift",
           "tokens"]
