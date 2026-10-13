"""What a cell's value is: its type, and a number read as a date or time.

Every value read comes back typed, with its address (ROADMAP.md, "Values and types")::

    Value(a="Stock!F2", v=12.5, t="number")
    Value(a="Stock!C2", v="2026-01-05", t="date", serial=46027.0)
    Value(a="Stock!F7", v=None, t="blank", formula=True, issue="no_cached_value")

Types: ``number`` (as stored), ``percent`` (the stored fraction), ``date``, ``time`` and
``datetime`` (ISO 8601 text, in the workbook's date system), ``text``, ``bool``, ``error``
(the code) and ``blank``.

**Dates** (ECMA-376 Part 1, §18.17.4.1, "Date Representation").  A serial's whole part is
a day and its fraction the time of day.  In the 1900 date system serial 1 is 1900-01-01,
and the system counts a 29 February 1900 (serial 60), as the standard's 1900 backward
compatibility date system says; a serial from 61 on is one more than its distance from
1899-12-31.  Serial 60 is read as ``"1900-02-29"`` (ISO text only: no calendar holds it).
In the 1904 date system serial 0 is 1904-01-01.  The type follows the format: a ``date``
format gives the day (the time of day, if any, stays in :attr:`Value.serial`), a ``time``
format the time of day, a ``datetime`` format both; an elapsed-time format (``[h]:mm:ss``)
gives an ISO 8601 duration (``"PT36H0M0S"``).  Times are rounded to the millisecond.  A
serial outside the system (negative, a day 0 in the 1900 system where the format shows a
day, past 9999-12-31) stays a ``number`` with ``issue="date_out_of_range"``.

**Locale** (ROADMAP.md, decision 4): a caller setting.  It decides how text that looks like
a number is recognised (``issue="number_as_text"``) and the day order of text that looks
like a date (``issue="date_as_text"``; ISO 8601 dates are recognised in any locale), and
names the UI language the language-dependent built-in formats are read in
(:mod:`xlsx_agent.read.styles`).
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

#: The language tags :attr:`Locale.language` takes: those whose built-in number formats
#: ECMA-376 Part 1, §18.8.30 lists.
FORMAT_LANGUAGES = ("zh-TW", "zh-CN", "ja-JP", "ko-KR", "th-TH")

#: Every value type.
TYPES = ("number", "percent", "date", "time", "datetime", "text", "bool", "error", "blank")

#: The most characters a cell holds in Excel; text this long was probably cut by the
#: program that wrote it (ROADMAP.md, "Messy workbooks": ``at_limit``).
TEXT_LIMIT = 32767

_DAY_ORDERS = ("dmy", "mdy", "ymd")

_EPOCH_1900 = date(1899, 12, 31)
_EPOCH_1904 = date(1904, 1, 1)
# 9999-12-31, the last day either system holds.
_LAST_1900 = (date(9999, 12, 31) - _EPOCH_1900).days + 1
_LAST_1904 = (date(9999, 12, 31) - _EPOCH_1904).days
_MS_PER_DAY = 86_400_000


@dataclass(frozen=True)
class Locale:
    """How text is read as numbers and dates, set by the caller, never by a model.

    ``decimal`` and ``thousands`` are the separators of a number written as text (by
    default ``.`` and ``,``); ``day_order`` (``"dmy"``, ``"mdy"`` or ``"ymd"``) is the order
    of a date written as text with ``/``, ``-`` or ``.``, ``None`` to recognise ISO 8601
    dates only; ``language`` (one of :data:`FORMAT_LANGUAGES`, or ``None``) the UI language
    the language-dependent built-in number formats are read in."""

    decimal: str = "."
    thousands: str | None = ","
    day_order: str | None = None
    language: str | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.decimal, str) or len(self.decimal) != 1 or self.decimal.isalnum():
            raise ValueError("decimal is one character, not a letter or digit")
        if self.thousands is not None and (not isinstance(self.thousands, str)
                                           or len(self.thousands) != 1
                                           or self.thousands.isalnum()
                                           or self.thousands == self.decimal):
            raise ValueError("thousands is one character, not a letter or digit, and not the decimal")
        if self.day_order is not None and self.day_order not in _DAY_ORDERS:
            raise ValueError(f"day_order is one of {_DAY_ORDERS} or None")
        if self.language is not None:
            for tag in FORMAT_LANGUAGES:
                if self.language.casefold() == tag.casefold():
                    object.__setattr__(self, "language", tag)
                    break
            else:
                raise ValueError(f"language is one of {FORMAT_LANGUAGES} or None")

    def number_pattern(self) -> re.Pattern:
        cached = _PATTERNS.get(self)
        if cached is None:
            d = re.escape(self.decimal)
            whole = r"\d+"
            if self.thousands is not None:
                whole = rf"(?:\d{{1,3}}(?:{re.escape(self.thousands)}\d{{3}})+|\d+)"
            cached = _PATTERNS[self] = re.compile(
                rf"[+-]?(?:{whole}(?:{d}\d*)?|{d}\d+)(?:[eE][+-]?\d+)?%?")
        return cached


_PATTERNS: dict[Locale, re.Pattern] = {}


@dataclass(slots=True)
class Value:
    """One cell's value, typed (see the module's notes).  Read-only by convention.

    ``sheet`` is the sheet as an address writes it, ``row`` and ``column`` its position
    (from 1); :attr:`a` is its address.  ``v`` the value; ``t`` its type (:data:`TYPES`).
    ``formula``: the cell holds a formula and ``v`` is its cached value.  ``issue``: one
    code for what is odd about the value (``no_cached_value``, ``number_as_text``,
    ``date_as_text``, ``date_out_of_range``, ``ambiguous_format``, ``unknown_format``,
    ``bad_value``, ``missing_string``), ``None`` when nothing is.  ``hidden``: on a hidden
    sheet, row or column.  ``at_limit``: text of :data:`TEXT_LIMIT` characters or more.
    ``serial``: the stored number of a date or time.  ``format``: the number format's code,
    when asked for.  ``formula_text`` (a normal or array formula, or a shared group's
    first cell), ``shared_from`` (a shared group's other cells: the address of the cell
    that holds the text), ``array`` (the range an array formula or a spill fills),
    ``volatile`` (the formula calls ``NOW``, ``TODAY``, ``RAND``, ``OFFSET`` or
    ``INDIRECT``: its value is as of the last save)."""

    sheet: str
    row: int
    column: int
    v: object = None
    t: str = "blank"
    formula: bool = False
    issue: str | None = None
    hidden: bool = False
    at_limit: bool = False
    serial: float | None = None
    format: str | None = None
    formula_text: str | None = field(default=None, repr=False)
    shared_from: str | None = None
    array: str | None = None
    volatile: bool = False

    @property
    def a(self) -> str:
        """The cell's address: ``Stock!F2``."""
        return f"{self.sheet}!{_letters(self.column)}{self.row}"

    @property
    def ref(self) -> str:
        """The cell's reference without its sheet: ``F2``."""
        return f"{_letters(self.column)}{self.row}"

    def as_dict(self) -> dict:
        """``{"a": ..., "v": ..., "t": ...}`` and whichever of the rest are set."""
        out = {"a": self.a, "v": self.v, "t": self.t}
        if self.formula:
            out["formula"] = True
        for name in ("issue", "serial", "format", "shared_from", "array"):
            value = getattr(self, name)
            if value is not None:
                out[name] = value
        for name in ("hidden", "at_limit", "volatile"):
            if getattr(self, name):
                out[name] = True
        return out


@dataclass(frozen=True)
class Problem:
    """Something odd met in the workbook, by code and address, never by content.

    ``address`` is the first cell it was met at (a sheet's or the workbook's problem has
    the sheet, or ``None``); ``count`` how many cells it was met at."""

    code: str
    address: str | None = None
    count: int = 1


_LETTERS: dict[int, str] = {}


def _letters(column: int) -> str:
    letters = _LETTERS.get(column)
    if letters is None:
        number, out = column, ""
        while number:
            number, rest = divmod(number - 1, 26)
            out = chr(65 + rest) + out
        letters = _LETTERS[column] = out
    return letters


# ------------------------------------------------------------------------------------------
# Numbers as dates and times
# ------------------------------------------------------------------------------------------


def serial_day(day: int, date1904: bool) -> str | None:
    """Day ``day`` of the date system as ISO 8601 text; ``None`` outside the system."""
    if date1904:
        if not 0 <= day <= _LAST_1904:
            return None
        return (_EPOCH_1904 + timedelta(days=day)).isoformat()
    if not 1 <= day <= _LAST_1900:
        return None
    if day == 60:
        return "1900-02-29"
    return (_EPOCH_1900 + timedelta(days=day if day < 60 else day - 1)).isoformat()


def day_serial(year: int, month: int, day: int, date1904: bool) -> int | None:
    """The serial of a day in the date system, :func:`serial_day`'s inverse (1900-02-29 is
    the 1900 system's serial 60); ``None`` for a day no calendar has, or one outside the
    system."""
    if (year, month, day) == (1900, 2, 29) and not date1904:
        return 60
    try:
        when = date(year, month, day)
    except ValueError:
        return None
    if date1904:
        number = (when - _EPOCH_1904).days
        return number if 0 <= number <= _LAST_1904 else None
    number = (when - _EPOCH_1900).days
    number += 1 if number >= 60 else 0
    return number if 1 <= number <= _LAST_1900 else None


def _split(serial: float) -> tuple[int, int]:
    """A serial as (day, milliseconds into the day), rounded to the millisecond."""
    total = round(serial * _MS_PER_DAY)
    return divmod(total, _MS_PER_DAY)


def _clock(ms: int) -> str:
    seconds, milli = divmod(ms, 1000)
    hours, rest = divmod(seconds, 3600)
    minutes, seconds = divmod(rest, 60)
    text = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    return f"{text}.{milli:03d}" if milli else text


def serial_text(serial: float, kind: str, date1904: bool, elapsed: bool = False) -> str | None:
    """A serial as ISO 8601 text for a ``date``, ``time`` or ``datetime`` format (see the
    module's notes); ``None`` when it is outside the date system."""
    if not math.isfinite(serial) or serial < 0:
        return None
    day, ms = _split(serial)
    if kind == "time":
        if elapsed:
            total = day * _MS_PER_DAY + ms
            seconds, milli = divmod(total, 1000)
            hours, rest = divmod(seconds, 3600)
            minutes, seconds = divmod(rest, 60)
            tail = f"{seconds}.{milli:03d}".rstrip("0").rstrip(".") if milli else f"{seconds}"
            return f"PT{hours}H{minutes}M{tail}S"
        if day > (_LAST_1904 if date1904 else _LAST_1900):
            return None
        return _clock(ms)
    when = serial_day(day, date1904)
    if when is None:
        return None
    return when if kind == "date" else f"{when}T{_clock(ms)}"


# ------------------------------------------------------------------------------------------
# Text that looks like a number or a date
# ------------------------------------------------------------------------------------------

_ISO_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?")
_LOCAL_DATE = re.compile(r"(\d{1,4})([/.-])(\d{1,2})\2(\d{1,4})")


def looks_like_number(text: str, locale: Locale) -> bool:
    """Whether ``text`` is a number written out with the locale's separators (a sign, an
    exponent and a trailing ``%`` allowed; space around it ignored)."""
    stripped = text.strip()
    return bool(stripped) and locale.number_pattern().fullmatch(stripped) is not None


def looks_like_date(text: str, locale: Locale) -> bool:
    """Whether ``text`` is a real calendar date: ISO 8601 (``2026-01-05``, with an
    optional time) in any locale, or with ``/``, ``-`` or ``.`` in the locale's
    :attr:`~Locale.day_order` when it has one."""
    stripped = text.strip()
    match = _ISO_DATE.fullmatch(stripped)
    if match:
        return _real(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    if locale.day_order is None:
        return False
    match = _LOCAL_DATE.fullmatch(stripped)
    if not match:
        return False
    parts = dict(zip(locale.day_order, (match.group(1), match.group(3), match.group(4))))
    if len(parts["y"]) not in (2, 4) or len(parts["d"]) > 2 or len(parts["m"]) > 2:
        return False
    year = int(parts["y"])
    if len(parts["y"]) == 2:
        year += 2000  # only whether such a day exists is asked; the century is not read
    return _real(year, int(parts["m"]), int(parts["d"]))


def _real(year: int, month: int, day: int) -> bool:
    try:
        date(year, month, day)
    except ValueError:
        return False
    return True
