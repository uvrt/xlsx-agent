"""Redacting terms: words the model must never read in a workbook, replaced wherever its
text is read -- the caller's rule, declared with the tables.

A declaration (:mod:`xlsx_agent.declare`) takes, beside its tables and hidden sheets::

    {"redact": ["Acme Corp", "re:Project\\\\s+\\\\w+"], "with": "[redacted]"}

so that a cell holding ``"Call ACME  corp. today"`` reads as ``"Call [redacted]. today"``.
The rules are the application's configuration, never a model's: no tool takes them, and
they cost no definition tokens.

**Terms.**  Each item of ``redact`` is a literal term -- found case aside (Python's
``IGNORECASE``), any run of white space in the text standing for a space in the term
(``"Acme Corp"`` finds ``"ACME\\ncorp"``) -- or, after ``re:``, a regular expression
(Python's :mod:`re`, case counting unless it says ``(?i)``).  ``with`` is the text each
match becomes (``"[redacted]"`` when not given).  ``whole_words`` (false by default)
finds a literal term only where no letter, digit or ``_`` touches either end, so
``"Ann"`` leaves ``"Annual"`` alone; by default a term is found inside words too -- a
term missed is a leak, an extra match only an untidy cell.  Several ``redact``
declarations may be given; their rules apply in order, each to what the one before left.

**Bounded.**  A regular expression is checked as a ``clean`` rule's ``strip`` pattern is
(:func:`~xlsx_agent.clean.compile_pattern`: at most 200 characters, no nested repeats,
alternations in repeats, backreferences or conditionals) and more strictly: **no** repeat
may give characters back to what follows it (``[^;]*;`` and ``\\s+\\w+`` do not; ``.*;``
does), and it must not match empty text.  A redaction searches the whole text, not a
window at its start, so a text is searched :data:`WINDOW` characters at a time, each
window overlapping the next by :data:`WINDOW`: a match is found when it is at most that
long (a longer one is found when no window cuts it), and a match a window's end cut is
checked against the whole text.  Each window then costs at most about the square of its
length: the worst pattern tried, ``[a-z]+[0-9]`` on a cell of 32,767 letters, took 0.26 s
on Python 3.12 and 0.24 s on 3.15 (Intel Core i7-6567U, 2026-10-15), once per distinct
text; two literal terms about 2 microseconds a cell.

**Where.**  The workbook's text as every result reads it (:meth:`Workbook.redact
<xlsx_agent.workbook.Workbook.redact>`, set before anything else is read): shared and
inline strings and a formula's cached text, so every value read, searched, profiled,
derived, compared, checked or placed in a deck or document; comments (text and author);
hyperlinks (address and place); a drop-down's literal list; a number format's code
(``formats``); an Excel table's column names; sheet names.  A formula whose text holds a match keeps its value, not its text
(``formula_withheld``, as one over a hidden column).  A defined name or an Excel table
whose *name* holds a match is not listed (the Excel table's cells stay, found as a block
of the sheet; a name is an identifier, and a stand-in for it would be a guess).  A sheet
whose name holds a match is shown, and addressed, by its name redacted (made unique with
`` (2)``, `` (3)``... when two would read alike); the caller's own declarations still
name it as stored.  Nothing tells the model a text was redacted but the replacement
itself: no count, no flag (a count of a name's matches is information about it).  The
program reads :attr:`Redaction.redacted` (how many distinct texts were changed so far)
and :attr:`Sheet.stored_name <xlsx_agent.workbook.Sheet.stored_name>`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Mapping, Sequence

from .clean import RuleError, compile_pattern

#: The keys a redaction's declaration takes.
REDACT_KEYS = ("redact", "with", "whole_words")

#: What a match becomes when ``with`` is not given.
DEFAULT_WITH = "[redacted]"

#: The prefix of a term that is a regular expression.
REGEX = "re:"

#: The most terms one declaration takes, the longest literal term and the longest ``with``.
MAX_TERMS = 200
MAX_TERM = 200
MAX_WITH = 50

#: How long a match a regular expression is sure to find; a text longer than twice this
#: is searched a window of twice this at a time.
WINDOW = 256

#: How many texts a redaction remembers the outcome of (then forgets them all).
_MEMO = 65536


class RedactError(ValueError):
    """A redaction refused: ``code`` (``bad_declaration``, ``bad_pattern``) and ``key``,
    ``redact[2]`` or ``with``.  The message never quotes a term."""

    def __init__(self, code: str, key: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.key = key
        self.message = message


@dataclass(frozen=True)
class Rule:
    """One declaration's terms, compiled: ``literal`` one pattern for its literal terms
    (``None``: none), ``patterns`` its regular expressions, ``replacement`` its ``with``,
    ``spec`` the declaration as given (for comparing two)."""

    literal: "re.Pattern[str] | None"
    patterns: "tuple[re.Pattern[str], ...]"
    replacement: str
    spec: tuple = field(repr=False)


def parse_rule(item: Mapping) -> Rule:
    """A declaration ``{"redact": [...], "with": ..., "whole_words": ...}`` checked and
    compiled; :class:`RedactError` otherwise."""
    for key in item:
        if key not in REDACT_KEYS:
            raise RedactError("bad_declaration", str(key), "a redaction takes redact, with "
                              "and whole_words")
    terms = item.get("redact")
    if isinstance(terms, str):
        terms = [terms]
    if isinstance(terms, (bytes, Mapping)) or not isinstance(terms, Sequence) or not terms \
            or len(terms) > MAX_TERMS:
        raise RedactError("bad_declaration", "redact", f"1 to {MAX_TERMS} terms: text, or "
                          "re: and a regular expression")
    replacement = item.get("with", DEFAULT_WITH)
    if not isinstance(replacement, str) or not replacement.strip() \
            or len(replacement) > MAX_WITH or any(ord(ch) < 32 for ch in replacement):
        raise RedactError("bad_declaration", "with", f"the text a match becomes: 1 to "
                          f"{MAX_WITH} characters on one line, not only spaces")
    whole = item.get("whole_words", False)
    if not isinstance(whole, bool):
        raise RedactError("bad_declaration", "whole_words", "true or false")
    literals: list[str] = []
    patterns: list[re.Pattern[str]] = []
    for n, term in enumerate(terms):
        key = f"redact[{n}]" if len(terms) > 1 or not isinstance(item.get("redact"), str) \
            else "redact"
        if not isinstance(term, str) or not term.strip():
            raise RedactError("bad_declaration", key, "a term is text, or re: and a regular "
                              "expression")
        if term.startswith(REGEX):
            try:
                pattern = compile_pattern(term[len(REGEX):], key=key, loose=0)
            except RuleError as error:
                raise RedactError(error.code, key, error.message) from None
            if pattern.search("") is not None:
                raise RedactError("bad_pattern", key, "the pattern matches empty text; "
                                  "it must match one character at least")
            patterns.append(pattern)
            continue
        if len(term) > MAX_TERM:
            raise RedactError("bad_declaration", key, f"a term is at most {MAX_TERM} "
                              "characters")
        literals.append(term)
    if replacement and any(_literal_hit(t, replacement, whole) for t in literals):
        raise RedactError("bad_declaration", "with", "the text a match becomes holds a term "
                          "it replaces")
    literal = None
    if literals:
        # Longest first, so that a term holding another is found whole.
        body = "|".join(r"\s+".join(re.escape(word) for word in term.split())
                        for term in sorted(set(literals), key=len, reverse=True))
        if whole:
            body = rf"(?<!\w)(?:{body})(?!\w)"
        literal = re.compile(body, re.IGNORECASE)
    return Rule(literal, tuple(patterns), replacement,
                spec=(tuple(terms), replacement, whole))


def _literal_hit(term: str, text: str, whole: bool) -> bool:
    body = r"\s+".join(re.escape(word) for word in term.split())
    if whole:
        body = rf"(?<!\w)(?:{body})(?!\w)"
    return re.search(body, text, re.IGNORECASE) is not None


class Redaction:
    """A workbook's redaction rules, applied to its text (:meth:`text`).  ``redacted``:
    how many distinct texts it has changed so far, for the program (never a result)."""

    def __init__(self, rules: "Sequence[Rule]") -> None:
        self.rules = tuple(rules)
        self._memo: dict[str, str] = {}
        self._changed: set[int] = set()

    @property
    def spec(self) -> tuple:
        return tuple(rule.spec for rule in self.rules)

    @property
    def redacted(self) -> int:
        """How many distinct texts were changed so far."""
        return len(self._changed)

    def text(self, text: str) -> str:
        """``text`` with every match of every rule replaced."""
        if not text:
            return text
        found = self._memo.get(text)
        if found is not None:
            return found
        out = text
        for rule in self.rules:
            if rule.literal is not None:
                out = rule.literal.sub(_constant(rule.replacement), out)
            for pattern in rule.patterns:
                out = _sub(pattern, rule.replacement, out)
        if out != text:
            self._changed.add(hash(text))
        if len(self._memo) >= _MEMO:
            self._memo.clear()
        self._memo[text] = out
        return out

    def holds(self, text: str | None) -> bool:
        """Whether ``text`` holds a match (a formula's text, a name)."""
        return text is not None and bool(text) and self.text(text) != text



def _constant(text: str):
    return lambda match: text


def _sub(pattern: "re.Pattern[str]", replacement: str, text: str) -> str:
    """``text`` with ``pattern``'s matches replaced, searched a window at a time (see the
    module's notes): every match of at most :data:`WINDOW` characters is found."""
    n = len(text)
    if n <= 2 * WINDOW:
        return pattern.sub(_constant(replacement), text)
    out: list[str] = []
    last = pos = 0
    while pos < n:
        end = min(n, pos + 2 * WINDOW)
        match = pattern.search(text, pos, end)
        if match is not None and match.end() == end < n:
            # Cut by the window's end (a longer match, or a \b or $ the cut made): the
            # whole text decides.
            whole = pattern.match(text, match.start())
            if whole is None or whole.end() == whole.start():
                pos = match.start() + 1
                continue
            match = whole
        if match is None:
            if end == n:
                break
            pos = end - WINDOW
            continue
        if match.end() == match.start():
            pos = match.start() + 1
            continue
        out.append(text[last:match.start()])
        out.append(replacement)
        last = pos = match.end()
    out.append(text[last:])
    return "".join(out)


def parse_rules(items: "Sequence[Any]") -> Redaction:
    return Redaction([parse_rule(item) for item in items])


__all__ = ["DEFAULT_WITH", "MAX_TERM", "MAX_TERMS", "MAX_WITH", "REDACT_KEYS", "REGEX",
           "RedactError", "Redaction", "Rule", "WINDOW", "parse_rule"]
