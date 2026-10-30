"""What a cell's value is: its type, and a number read as a date or time.

Every value read comes back typed, with its address (ROADMAP.md, "Values and types")::

    Value(a="Stock!F2", v=12.5, t="number")
    Value(a="Stock!C2", v="2026-01-05", t="date", serial=46027.0)
    Value(a="Stock!F7", v=None, t="blank", formula=True, issue="no_cached_value")

Types: ``number`` (as stored), ``percent`` (the stored fraction), ``date``, ``time`` and
``datetime`` (ISO 8601 text, in the workbook's date system), ``text``, ``bool``, ``error``
(the code) and ``blank``; and ``list``, a tuple of texts, which no cell stores: a declared
column's text split by its caller's rule (:mod:`xlsx_agent.clean`).

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

**Locale** (ROADMAP.md, decision 4): a caller setting, the shared
:class:`ooxml_common.locales.Locale` (or any object with its fields; en-US,
:data:`~ooxml_common.locales.EN_US`, when none is given).  Its ``decimal_separator`` and
``thousands_separator`` decide how text that looks like a number is recognised
(``issue="number_as_text"``), its ``date_order`` the order of day and month in text that
looks like a date (``issue="date_as_text"``) when the reader is asked to read such dates,
and its ``language`` the UI language the language-dependent built-in formats are read in
(:func:`format_language`, :mod:`xlsx_agent.read.styles`).

Two reader settings are not locale facts and stay the reader's own
(:meth:`xlsx_agent.Workbook.open`): ``grouping`` -- whether a number's whole part may be
grouped with the thousands separator (on, as before) -- and ``local_dates`` -- whether
text dates other than ISO 8601 are recognised (off: ISO 8601 dates only, in any locale).
What Excel does, measured (Excel 16.106 Mac, Dutch UI, regions nl-BE and en-US,
2026-10-13; ``VALUE`` and ``--A`` on text): the region's separators decide, grouping
included -- ``1.234,5``, ``1234,5`` parsed and ``1,234.5`` an error under nl-BE, the other
way round under en-US; ``3/4/2026`` and ``3-4-2026`` are 3 April under nl-BE (``dmy``) and
4 March under en-US (``mdy``), ``13/4/2026`` an error under en-US and ``4/13/2026`` under
nl-BE; ``2026-04-03`` is a date in both.  So ``local_dates=True`` with the caller's
``date_order`` is Excel's reading; it stays off by default so that a text such as
``05/01/2026`` is flagged only when the caller says which day it is.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal
from fractions import Fraction
from typing import Any, cast

from ooxml_common.locales import DATE_ORDERS, EN_US

from ..address import _letters, format_cell

#: The UI languages whose built-in number formats ECMA-376 Part 1, §18.8.30 lists: the
#: tables a locale's ``language`` maps to (:func:`format_language`).
FORMAT_LANGUAGES = ("zh-TW", "zh-CN", "ja-JP", "ko-KR", "th-TH")

#: Every value type (``list`` last: a cell's text split by a declared rule,
#: :mod:`xlsx_agent.clean`).
TYPES = ("number", "percent", "date", "time", "datetime", "text", "bool", "error", "blank",
         "list")

#: A formula's cached value: as last saved with no known invalidation (never checked to be
#: current), invalidated since, or absent (ooxml-edit's ``CACHE_STATUSES``).
CACHE_STATUSES = ("cached_unverified", "invalidated", "absent")

#: The most characters a cell holds in Excel; text this long was probably cut by the
#: program that wrote it (ROADMAP.md, "Messy workbooks": ``at_limit``).
TEXT_LIMIT = 32767

_EPOCH_1900 = date(1899, 12, 31)
_EPOCH_1904 = date(1904, 1, 1)
# 9999-12-31, the last day either system holds.
_LAST_1900 = (date(9999, 12, 31) - _EPOCH_1900).days + 1
_LAST_1904 = (date(9999, 12, 31) - _EPOCH_1904).days
_MS_PER_DAY = 86_400_000


#: The attributes a locale has (:class:`ooxml_common.locales.Locale`'s fields), each a
#: string: all :func:`reading_locale` checks of an object it is given.
LOCALE_FIELDS = ("decimal_separator", "thousands_separator", "date_order", "language",
                 "list_separator")


def reading_locale(locale: Any) -> Any:
    """``locale`` as a workbook keeps it: ``None`` is :data:`ooxml_common.locales.EN_US`;
    anything else is kept as given -- normally an :class:`ooxml_common.locales.Locale`, or
    any object with its five string fields (an application's own type, as
    ``ooxml_edit.tools`` takes one).  Raises :class:`TypeError` for anything else: a
    language tag alone is not a locale."""
    if locale is None:
        return EN_US
    missing = [name for name in LOCALE_FIELDS
               if not isinstance(getattr(locale, name, None), str)]
    if missing:
        raise TypeError(f"locale takes an ooxml_common.locales.Locale (or an object with "
                        f"string {', '.join(LOCALE_FIELDS)}), not {type(locale).__name__}: "
                        f"no {', '.join(missing)}")
    return locale


def format_language(language: str | None) -> str | None:
    """The :data:`FORMAT_LANGUAGES` table a locale's ``language`` reads the
    language-dependent built-in number formats with: that tag when the language is one of
    them (compared in any case, ``_`` read as ``-``), else ``None`` -- the ids whose kind
    differs between languages are then ``"ambiguous"`` (:mod:`xlsx_agent.read.styles`).

    Measured (Excel 16.106 Mac, Dutch UI, regions nl-BE and en-US, 2026-10-13): ids 27 to
    36 and 50 to 58 show as id 14 (a date) or id 21 (a time) in both regions, so a
    language outside the five tables gets no table of its own here.  An en-US Excel UI is
    not measured."""
    if not isinstance(language, str):
        return None
    wanted = language.replace("_", "-").casefold()
    for tag in FORMAT_LANGUAGES:
        if wanted == tag.casefold():
            return tag
    return None


def number_pattern(decimal: str, thousands: str | None) -> re.Pattern:
    """A number written out with ``decimal`` and, unless ``None``, ``thousands`` grouping
    the whole part in threes (a sign, an exponent and a trailing ``%`` allowed)."""
    key = (decimal, thousands)
    cached = _PATTERNS.get(key)
    if cached is None:
        d = re.escape(decimal)
        whole = r"\d+"
        if thousands is not None:
            whole = rf"(?:\d{{1,3}}(?:{re.escape(thousands)}\d{{3}})+|\d+)"
        cached = _PATTERNS[key] = re.compile(
            rf"[+-]?(?:{whole}(?:{d}\d*)?|{d}\d+)(?:[eE][+-]?\d+)?%?")
    return cached


_PATTERNS: dict[tuple, re.Pattern] = {}


@dataclass(slots=True)
class Value:
    """One cell's value, typed (see the module's notes).  Read-only by convention.

    ``sheet`` is the sheet as an address writes it, ``row`` and ``column`` its position
    (from 1); :attr:`a` is its address.  ``v`` the value; ``t`` its type (:data:`TYPES`).
    ``formula``: the cell holds a formula and ``v`` is its cached value.  ``issue``: one
    code for what is odd about the value (``no_cached_value``, ``number_as_text``,
    ``date_as_text``, ``date_out_of_range``, ``ambiguous_format``, ``unknown_format``,
    ``bad_value``, ``missing_string``; ``unmeasured_format``: a number a declared column
    reads as shown, whose shown text is not measured, :mod:`xlsx_agent.clean`), ``None``
    when nothing is.  ``hidden``: on a hidden
    sheet, row or column.  ``at_limit``: text of :data:`TEXT_LIMIT` characters or more.
    ``serial``: the stored number of a date or time.  ``format``: the number format's code,
    when asked for.  ``formula_text`` (a normal or array formula, or a shared group's
    first cell), ``shared_from`` (a shared group's other cells: the address of the cell
    that holds the text), ``array`` (the range an array formula or a spill fills),
    ``volatile`` (the formula calls ``NOW``, ``TODAY``, ``RAND``, ``OFFSET`` or
    ``INDIRECT``: its value is as of the last save).  ``cache``: a formula's cached value's
    status (:data:`CACHE_STATUSES`; ``None`` for a cell without a formula):
    ``cached_unverified`` as last saved, with no known invalidation -- never checked to be
    current; ``invalidated`` when the workbook has been written since it was opened, or
    arrived asking for a recalculation or saved with manual calculation; ``absent`` when
    there is none (``issue="no_cached_value"``).  ``external``: the formula refers to another
    workbook (a workbook index, ``[1]Prices!B2``, or a defined name that does), whose
    values may have changed without any change here.  ``formula_withheld``: the formula
    reaches a cell hidden from the model by a declared table, so its text is not given
    (``formula_text`` is ``None``; ``v`` is its cached value, as for any formula:
    :mod:`xlsx_agent.declare`).  ``origin``: the cell's own sheet (as
    an address writes it), row and column, when ``sheet``, ``row`` and ``column`` are its
    place in a declared table read as one (a table over several sheets, or with columns
    hidden from the model: :mod:`xlsx_agent.declare`); :attr:`a` is then the cell's own
    address.  ``cleaned``: a declared column's rule changed the text (:mod:`xlsx_agent.clean`;
    ``v`` is the text cleaned, or a tuple of items for ``t="list"``), and ``original`` is
    the text as stored.  ``markdown``: a text value's Markdown, when the read asked for it
    (:mod:`xlsx_agent.markdown`: bold, italic and strikethrough runs, list lines, every
    other character escaped); a read with ``markdown`` gives it as ``v``."""

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
    formula_withheld: bool = False
    shared_from: str | None = None
    array: str | None = None
    volatile: bool = False
    cache: str | None = None
    external: bool = False
    origin: "tuple[str, int, int] | None" = None
    cleaned: bool = False
    original: str | None = field(default=None, repr=False)
    markdown: str | None = field(default=None, repr=False)

    @property
    def a(self) -> str:
        """The cell's address: ``Stock!F2`` (its own, for a value of a declared table), in
        the canonical form (:mod:`xlsx_agent.address`: what
        :func:`~xlsx_agent.address.canonical_address` gives of it)."""
        if self.origin is not None:
            sheet, row, column = self.origin
            return format_cell(sheet, row, column)
        return format_cell(self.sheet, self.row, self.column)

    @property
    def ref(self) -> str:
        """The cell's reference without its sheet: ``F2``."""
        if self.origin is not None:
            return f"{_letters(self.origin[2])}{self.origin[1]}"
        return f"{_letters(self.column)}{self.row}"

    @property
    def place(self) -> "tuple[str, int, int]":
        """The cell's own sheet (as an address writes it), row and column."""
        if self.origin is not None:
            return self.origin
        return (self.sheet, self.row, self.column)

    def as_dict(self) -> dict:
        """``{"a": ..., "v": ..., "t": ...}`` and whichever of the rest are set (a list's
        ``v`` as a list)."""
        out = {"a": self.a, "v": list(cast("list[str]", self.v)) if self.t == "list" else self.v,
               "t": self.t}
        if self.formula:
            out["formula"] = True
        for name in ("cache", "issue", "serial", "format", "shared_from", "array"):
            value = getattr(self, name)
            if value is not None:
                out[name] = value
        for name in ("hidden", "at_limit", "volatile", "external", "cleaned"):
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


def looks_like_number(text: str, locale: Any = EN_US, *, grouping: bool = True) -> bool:
    """Whether ``text`` is a number written out with the locale's separators (a sign, an
    exponent and a trailing ``%`` allowed; space around it ignored).  ``grouping``: the
    whole part may be grouped in threes with the thousands separator."""
    stripped = text.strip()
    if not stripped:
        return False
    pattern = number_pattern(locale.decimal_separator,
                             locale.thousands_separator if grouping else None)
    return pattern.fullmatch(stripped) is not None


def text_number(text: str, locale: Any = EN_US, *, grouping: bool = True) -> float | None:
    """The number ``text`` writes out, when :func:`looks_like_number` says it is one
    (``None`` otherwise): the thousands separators dropped, the decimal separator read as
    a point, a trailing ``%`` dividing by 100 (``25%`` is 0.25, the stored fraction), and
    the decimal read exactly and rounded once to the nearest double.  What a caller who
    opts in to ``coerce=["number_as_text"]`` gets (:mod:`xlsx_agent.derive`)."""
    if not looks_like_number(text, locale, grouping=grouping):
        return None
    stripped = text.strip()
    hundredths = stripped.endswith("%")
    if hundredths:
        stripped = stripped[:-1]
    if grouping and locale.thousands_separator:
        stripped = stripped.replace(locale.thousands_separator, "")
    stripped = stripped.replace(locale.decimal_separator, ".")
    exact = Fraction(Decimal(stripped))
    if hundredths:
        exact /= 100
    return float(exact)


def looks_like_date(text: str, locale: Any = EN_US, *, local_dates: bool = False) -> bool:
    """Whether ``text`` is a real calendar date: ISO 8601 (``2026-01-05``, with an
    optional time) in any locale, or, with ``local_dates``, with ``/``, ``-`` or ``.`` in
    the locale's ``date_order``."""
    stripped = text.strip()
    match = _ISO_DATE.fullmatch(stripped)
    if match:
        return _real(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    order = locale.date_order
    if not local_dates or order not in DATE_ORDERS:
        return False
    match = _LOCAL_DATE.fullmatch(stripped)
    if not match:
        return False
    parts = dict(zip(order, (match.group(1), match.group(3), match.group(4))))
    if len(parts["y"]) not in (2, 4) or len(parts["d"]) > 2 or len(parts["m"]) > 2:
        return False
    year = int(parts["y"])
    if len(parts["y"]) == 2:
        year += 2000  # only whether such a day exists is asked; the century is not read
    return _real(year, int(parts["m"]), int(parts["d"]))


def text_date(text: str, locale: Any = EN_US, *, local_dates: bool = False) -> str | None:
    """The date ``text`` writes out, as ISO 8601 text (``2026-01-05``, a time kept as
    ``2026-01-05T10:30``), when :func:`looks_like_date` says it is one; ``None`` otherwise,
    and for a two-digit year, whose century Excel decides by a rule not measured here.
    What a caller who opts in to ``coerce=["date_as_text"]`` gets (:mod:`xlsx_agent.derive`)
    and what ``xlsx_profile`` gives the range of."""
    if not looks_like_date(text, locale, local_dates=local_dates):
        return None
    stripped = text.strip()
    if _ISO_DATE.fullmatch(stripped):
        return stripped.replace(" ", "T", 1)
    match = _LOCAL_DATE.fullmatch(stripped)
    assert match is not None   # looks_like_date matched it
    parts = dict(
zip(locale.date_order, (match.group(1), match.group(3), match.group(4))))
    if len(parts["y"]) != 4:
        return None
    return f"{parts['y']}-{int(parts['m']):02d}-{int(parts['d']):02d}"


def _real(year: int, month: int, day: int) -> bool:
    try:
        date(year, month, day)
    except ValueError:
        return False
    return True
