"""Cleaning a declared column's text: a leading label stripped, white space collapsed, a
cell split into a list -- the caller's rules, applied once, where the table is read.

A declared field (:mod:`xlsx_agent.declare`) takes ``clean``::

    {"header": "Owners", "id": "owners",
     "clean": {"strip": r"^Lead:[^;]*;\\s*", "collapse_space": True, "split": ";"}}

so that a cell holding ``"Lead: Robin;  Kim ;Ash"`` reads as ``("Kim", "Ash")``.  The rules
are the application's configuration, never a model's: no tool takes them, and they cost
no definition tokens.

**The rules**, applied in this order, to text values only (a number, a date, a bool, an
error or a blank is never changed; a formula's cached text is text):

* ``strip``: a regular expression (Python's :mod:`re`) matched at the start of the text
  (:meth:`re.Pattern.match`: a leading ``^`` changes nothing); what it matches is removed.
  It sees the first :data:`LABEL_WINDOW` characters of a cell, so a label is found within
  them.  A pattern that matches nothing changes nothing.
* ``split``: a separator, literal text (``";"``, ``"\\n"``, ``" / "``), not a pattern:
  the text is cut at each, every piece trimmed of white space at its ends, and the empty
  pieces dropped.  The value is then a **list** (type ``list``, a tuple in Python, a list
  in a tool's JSON), even of one item -- so the column has one type.
* ``collapse_space``: every run of white space (as :meth:`str.split` finds it: spaces, tabs,
  line breaks, no-break spaces...) becomes one space, and none is left at either end; on
  each item of a split.

A text cleaned to nothing (a cell that held only a label) is ``blank``.  A cleaned text
that reads as a number or a date in the workbook's locale is flagged as the reader flags
one (``number_as_text``, ``date_as_text``), so a derive's ``coerce`` reads it.

**Never silent.**  A value a rule changed -- a label removed, white space changed, a split
into more than one item or items trimmed -- says so (:attr:`Value.cleaned
<xlsx_agent.read.types.Value.cleaned>`, ``"cleaned": true`` in a tool's cell), and keeps
the text as stored (:attr:`Value.original <xlsx_agent.read.types.Value.original>`, the
Python API only).  A one-item list of text no rule changed is not counted.  They are
counted, with their first addresses, as ``coerced`` cells are: ``cleaned`` in an
``xlsx_read`` or ``xlsx_query`` page (its cells), a ``cleaned`` flag in ``xlsx_profile`` and
``describe`` (per column), ``cleaned`` per column in ``xlsx_derive``'s result
(:attr:`Derived.cleaned <xlsx_agent.derive.Derived.cleaned>`), and a
:class:`CleanedWarning` in the result of a deck's or document's tool that took the table
as its ``from``.  A cell read by its address on the sheet (``'Team 1'!C2``) is the cell
as stored.

**Lists downstream.**  ``xlsx_query`` and a derive's ``where`` test a list with ``eq``,
``in`` and ``contains``, a row passing when any item does, and sort it item by item.
``xlsx_derive`` counts a list column's cells (``count``) and refuses every other function,
``group_by`` and ``pivot`` on it (``wrong_type``): a cell of several values belongs to no
one group, and an item count would be a new function in the tool's definition.  A
derived dataset (``d:``) so never holds a list.  A deck's or document's table or chart
from ``t:`` gets a list as its items joined by a line break (one cell, a line per item:
:data:`~xlsx_agent.dataset.LIST_JOIN`).  ``xlsx_compare`` compares lists as values (a
change in any item is a change of the cell).  ``xlsx_validate_data`` checks a list as
``text``, item by item.

**Bounded** (the patterns are the caller's, and a pattern can backtrack for ever).  A
pattern is compiled when declared (:func:`parse_clean`), at most :data:`MAX_PATTERN`
characters, and refused (``bad_pattern``) when its backtracking could grow faster than the
square of the text it sees: a repeat inside a repeat (``(a+)+``, ``(a?b)*``), an
alternation inside a repeat (``(a|ab)*``), a backreference or a conditional group, more
than :data:`MAX_LOOSE_REPEATS` repeat that can give characters back to what follows it, or
alternations whose choices multiply past :data:`MAX_CHOICES`.  A repeat of one character
class followed by a character it cannot match (``[^;]*;``, ``\\s*Lead``) cannot give any
back, nor can a repeat that ends the pattern; ``.*;`` and ``\\s*[^;]*`` can.  With the
:data:`LABEL_WINDOW` a pattern sees, a cell then costs at most about the square of 256
steps, however long its text: the worst form allowed (``^\\s*\\s*x``, on a cell of 32,767
spaces) took 0.2 ms a cell on Python 3.12 and 0.1 ms on 3.15 (measured 2026-10-13, an
Intel Core i7-6567U); a declared label (``^Lead:[^;]*;\\s*``) takes about a microsecond.
The checks read the pattern as Python's own parser does (:mod:`re._parser`, Python 3.11
on); a pattern is refused, not guessed at, where that parser is missing.

**Which values.**  The rules apply to the table's data cells as read through it -- by
``t:`` address, ``describe``'s samples, ``xlsx_read``, ``xlsx_query``, ``xlsx_profile``,
``xlsx_derive``, ``xlsx_compare``, ``xlsx_validate_data`` and a deck's or document's
``from`` -- never to the header, and never to the key column, which is refused
(``clean_key``): an address by key names the key as the cell holds it, and two keys that
cleaned alike would name one row.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Mapping

from .read.types import looks_like_date, looks_like_number

if TYPE_CHECKING:  # pragma: no cover
    from .read.types import Value

#: The keys a field's ``clean`` takes.
CLEAN_KEYS = ("strip", "split", "collapse_space")

#: The longest ``strip`` pattern (characters).
MAX_PATTERN = 200

#: The longest ``split`` separator (characters).
MAX_SEPARATOR = 10

#: How many characters at the start of a cell a ``strip`` pattern sees.
LABEL_WINDOW = 256

#: How many repeats in a pattern may give characters back to what follows them.
MAX_LOOSE_REPEATS = 1

#: How many ways the alternations of a pattern may combine, at most (each ``a|b|c``
#: multiplies by its choices).
MAX_CHOICES = 16


class CleanedWarning(UserWarning):
    """Values a declared column's rules changed went into a deck's or document's table or
    chart (``from: "t:checks"``): how many, and the first cells, never their text.  The
    tool layer adds it to the tool's result as a warning."""


class RuleError(ValueError):
    """A ``clean`` rule refused: ``code`` (``bad_declaration``, ``bad_pattern``) and
    ``key``, the rule at fault (``strip``); the message names the rule, never a cell."""

    def __init__(self, code: str, key: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.key = key
        self.message = message


@dataclass(frozen=True)
class Clean:
    """A column's rules: ``strip`` (a compiled pattern, or ``None``), ``split`` (the
    separator, or ``None``) and ``collapse_space``.  ``spec`` is the rules as declared."""

    strip: "re.Pattern[str] | None" = None
    split: str | None = None
    collapse_space: bool = False

    @property
    def spec(self) -> dict:
        out: dict[str, Any] = {}
        if self.strip is not None:
            out["strip"] = self.strip.pattern
        if self.split is not None:
            out["split"] = self.split
        if self.collapse_space:
            out["collapse_space"] = True
        return out

    def apply(self, text: str) -> "tuple[str | tuple[str, ...] | None, bool]":
        """``text`` cleaned -- text, a tuple of items (``split``), or ``None`` when nothing
        is left -- and whether a rule changed it (a one-item list of the text as it was is
        no change)."""
        out = text
        if self.strip is not None:
            found = self.strip.match(out, 0, LABEL_WINDOW)
            if found is not None and found.end():
                out = out[found.end():]
        if self.split is not None:
            items = []
            for piece in out.split(self.split):
                piece = " ".join(piece.split()) if self.collapse_space else piece.strip()
                if piece:
                    items.append(piece)
            changed = items != [text]
            return (tuple(items) if items else None), changed
        if self.collapse_space:
            out = " ".join(out.split())
        return (out if out else None), out != text


def parse_clean(item: Any) -> Clean:
    """A field's ``clean`` (a mapping of :data:`CLEAN_KEYS`) checked and compiled;
    :class:`RuleError` otherwise."""
    if not isinstance(item, Mapping) or not item:
        raise RuleError("bad_declaration", "", "clean is a mapping of strip, split and "
                        "collapse_space")
    for key in item:
        if key not in CLEAN_KEYS:
            raise RuleError("bad_declaration", str(key), "unknown rule")
    strip = None
    if "strip" in item:
        strip = compile_pattern(item["strip"])
    split = item.get("split")
    if split is not None and not (isinstance(split, str) and split
                                  and len(split) <= MAX_SEPARATOR):
        raise RuleError("bad_declaration", "split", f"a separator is text of 1 to "
                        f"{MAX_SEPARATOR} characters, taken as written (not a pattern)")
    collapse = item.get("collapse_space", False)
    if not isinstance(collapse, bool):
        raise RuleError("bad_declaration", "collapse_space", "true or false")
    if strip is None and split is None and not collapse:
        raise RuleError("bad_declaration", "", "clean names no rule: strip, split or "
                        "collapse_space")
    return Clean(strip=strip, split=split, collapse_space=collapse)


def compile_pattern(text: Any, *, key: str = "strip",
                    loose: int = MAX_LOOSE_REPEATS) -> "re.Pattern[str]":
    """``text`` compiled, when it is a pattern whose cost is bounded (see the module's
    notes; ``loose`` the most repeats that may give characters back, ``key`` the rule
    named when refused); :class:`RuleError` ``bad_pattern`` otherwise."""
    if not isinstance(text, str) or not text or len(text) > MAX_PATTERN:
        raise RuleError("bad_pattern", key, f"a pattern is text of 1 to {MAX_PATTERN} "
                        "characters")
    try:
        pattern = re.compile(text)
    except re.error as error:
        raise RuleError("bad_pattern", key, f"not a regular expression: {error.msg} "
                        f"at position {error.pos}") from None
    _check(text, pattern, key, loose)
    return pattern


def clean_value(value: "Value", rule: Clean, workbook: Any) -> None:
    """Apply ``rule`` to ``value`` in place, when it is text: its cleaned text (flagged
    again as the reader flags text, in ``workbook``'s locale), its items (``list``), or
    ``blank``; ``cleaned`` and ``original`` set when a rule changed it."""
    if value.t != "text" or not isinstance(value.v, str):
        return
    text = value.v
    out, changed = rule.apply(text)
    if out is None:
        value.v, value.t, value.issue = None, "blank", None
    elif isinstance(out, tuple):
        value.v, value.t, value.issue = out, "list", None
    elif changed:
        value.v = out
        if not value.formula:
            value.issue = looks(out, workbook)
    if changed:
        value.cleaned, value.original = True, text


def looks(text: str, workbook: Any) -> str | None:
    """``number_as_text`` or ``date_as_text`` when ``text`` reads as one in ``workbook``'s
    locale and reader settings, as the reader flags a cell's text."""
    locale = workbook.locale
    if looks_like_number(text, locale, grouping=workbook.grouping):
        return "number_as_text"
    if looks_like_date(text, locale, local_dates=workbook.local_dates):
        return "date_as_text"
    return None


# -- the cost of a pattern -------------------------------------------------------------------


_ONE_CHAR = ("LITERAL", "NOT_LITERAL", "IN", "ANY", "CATEGORY")
_REPEATS = ("MAX_REPEAT", "MIN_REPEAT", "POSSESSIVE_REPEAT")
_REFUSED = ("GROUPREF", "GROUPREF_IGNORE", "GROUPREF_LOC_IGNORE", "GROUPREF_UNI_IGNORE",
            "GROUPREF_EXISTS")

#: The characters two one-character matchers are compared on: every code point up to the
#: CJK symbols, and a few spaces and marks past them.
_SAMPLE = [chr(c) for c in range(0x3001)] + ["﻿", "，", "：", "；"]

_CATEGORIES = {"CATEGORY_DIGIT": re.compile(r"\d"), "CATEGORY_SPACE": re.compile(r"\s"),
               "CATEGORY_WORD": re.compile(r"\w")}


def _check(text: str, pattern: "re.Pattern[str]", key: str = "strip",
           loose: int = MAX_LOOSE_REPEATS) -> None:
    try:
        from re import _parser as parser
    except ImportError:  # pragma: no cover - re._parser is Python 3.11's on
        raise RuleError("bad_pattern", key, "this Python cannot check a pattern's "
                        "cost; patterns are refused") from None
    parsed = parser.parse(text, pattern.flags)
    folds = bool(pattern.flags & re.IGNORECASE) or re.search(r"\(\?[a-zA-Z-]*i", text) is not None
    state = {"loose": 0, "choices": 1, "key": key}
    _walk(list(parsed), state, folds, inside=False, last=True)
    if state["loose"] > loose:
        raise RuleError("bad_pattern", key, f"{state['loose']} repeats can give "
                        "characters back to what follows them (as .* before ; can), and "
                        f"at most {loose} may: write [^;]*; for .*;")
    if state["choices"] > MAX_CHOICES:
        raise RuleError("bad_pattern", key, f"its alternations combine in more than "
                        f"{MAX_CHOICES} ways")


def _walk(items: list, state: dict, folds: bool, *, inside: bool, last: bool) -> None:
    """Check a sequence of parsed items: ``inside`` a repeat, ``last`` when nothing of the
    pattern follows the sequence."""
    for n, (op, av) in enumerate(items):
        name = str(op)
        if name in _REFUSED:
            raise RuleError("bad_pattern", state["key"], "a backreference or conditional group "
                            "makes a pattern's cost unbounded")
        if name in _REPEATS:
            low, high, body = av
            body = list(body)
            if inside:
                raise RuleError("bad_pattern", state["key"], "a repeat inside a repeat (as "
                                "(a+)+) can take for ever")
            _walk(body, state, folds, inside=high > 1, last=False)
            if high >= 1 and name != "POSSESSIVE_REPEAT" \
                    and not _tight(body, items[n + 1:], last and n == len(items) - 1, folds):
                state["loose"] += 1
        elif name == "BRANCH":
            if inside:
                raise RuleError("bad_pattern", state["key"], "an alternation inside a repeat "
                                "(as (a|ab)*) can take for ever")
            choices = av[1]
            state["choices"] *= len(choices)
            if state["choices"] > MAX_CHOICES:
                raise RuleError("bad_pattern", state["key"], f"its alternations combine in more "
                                f"than {MAX_CHOICES} ways")
            for choice in choices:
                _walk(list(choice), state, folds, inside=inside, last=False)
        elif name == "SUBPATTERN":
            _walk(list(av[3]), state, folds, inside=inside,
                  last=last and n == len(items) - 1)
        elif name == "ATOMIC_GROUP":
            _walk(list(av), state, folds, inside=inside, last=False)
        elif name in ("ASSERT", "ASSERT_NOT"):
            _walk(list(av[1]), state, folds, inside=inside, last=False)


def _single(body: list) -> bool:
    return len(body) == 1 and str(body[0][0]) in _ONE_CHAR


def _tight(body: list, rest: list, at_end: bool, folds: bool) -> bool:
    """Whether a repeat gives nothing back: it ends the pattern, or it repeats one
    character class and what follows begins with a character that class never matches."""
    if at_end:
        return True
    if not _single(body):
        return False
    following = _first(rest)
    return following is not None and _disjoint(body[0], following, folds)


def _first(rest: list) -> "tuple | None":
    """The one-character matcher the rest of a sequence must begin with, or ``None``."""
    if not rest:
        return None
    op, av = rest[0]
    name = str(op)
    if name in _ONE_CHAR:
        return rest[0]
    if name in ("SUBPATTERN",):
        return _first(list(av[3]))
    if name == "ATOMIC_GROUP":
        return _first(list(av))
    if name in _REPEATS and av[0] >= 1:
        return _first(list(av[2]))
    return None


def _disjoint(one: tuple, two: tuple, folds: bool) -> bool:
    a, b = _matcher(one, folds), _matcher(two, folds)
    return not any(a(ch) and b(ch) for ch in _SAMPLE)


def _matcher(node: tuple, folds: bool):
    """A test of one character for a parsed one-character matcher; with ``folds``, a
    character matches when any of its cases does (so two matchers are disjoint only when
    they are in every case)."""
    test = _plain(node)
    if not folds:
        return test
    return lambda ch: any(test(c) for c in {ch, ch.lower(), ch.upper(), ch.casefold()}
                          if len(c) == 1)


def _plain(node: tuple):
    op, av = node
    name = str(op)
    if name == "LITERAL":
        return lambda ch: ord(ch) == av
    if name == "NOT_LITERAL":
        return lambda ch: ord(ch) != av
    if name == "ANY":
        return lambda ch: True
    if name == "CATEGORY":
        return _category(av)
    items = list(av)
    negate = bool(items) and str(items[0][0]) == "NEGATE"
    tests = [_item(item) for item in (items[1:] if negate else items)]
    if negate:
        return lambda ch: not any(t(ch) for t in tests)
    return lambda ch: any(t(ch) for t in tests)


def _item(item: tuple):
    op, av = item
    name = str(op)
    if name == "LITERAL":
        return lambda ch: ord(ch) == av
    if name == "RANGE":
        low, high = av
        return lambda ch: low <= ord(ch) <= high
    if name == "CATEGORY":
        return _category(av)
    return lambda ch: True          # anything else: taken to match every character


def _category(av: Any):
    name = str(av)
    base = name.replace("NOT_", "")
    found = _CATEGORIES.get(base)
    if found is None:
        return lambda ch: True
    if "NOT_" in name:
        return lambda ch: found.match(ch) is None
    return lambda ch: found.match(ch) is not None


__all__ = ["CLEAN_KEYS", "Clean", "CleanedWarning", "LABEL_WINDOW", "MAX_CHOICES",
           "MAX_LOOSE_REPEATS", "MAX_PATTERN", "MAX_SEPARATOR", "RuleError", "clean_value",
           "compile_pattern", "looks", "parse_clean"]
