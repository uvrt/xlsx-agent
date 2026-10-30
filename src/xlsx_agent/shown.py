"""A number as Excel shows it with its format code, where that is measured; nothing where
it is not.

A declared column's ``clean`` with ``"on": "shown"`` (:mod:`xlsx_agent.clean`) reads a
number cell as the text Excel shows for it -- ``5.2`` in ``General`` as ``5.2`` (``5,2``
with a decimal comma), ``0.25`` in ``0%`` as ``25%``, ``1234.5`` in ``#,##0.00`` as
``1,234.50`` -- so that its rules (``extract``, ``strip``, ``split``) apply to numbers as to
text.  :func:`shown_text` makes that text, or gives ``None`` when what Excel shows for it is
not measured: the cell is then refused (blank, ``issue="unmeasured_format"``), never
guessed.

**What is measured** is :data:`ooxml_common.numfmt.MEASURED`: 108 format codes, each with
the text Excel for Mac 16.106 showed the number 46027.5 in (2026-10-25, region Dutch
(Belgium)), 43 of them showing it as a number.  The forms those 43 hold are rendered here,
and the tests hold each of the 36 rendered to Excel's text:

* digit placeholders ``0`` (a digit, a zero when there is none) and ``#`` (a digit when
  there is one), a decimal point (the locale's ``decimal_separator``; ``#,##0.00`` showed
  ``46.027,50``), a ``,`` between placeholders grouping the whole part in threes (the
  locale's ``thousands_separator``), and ``,`` after the last placeholder dividing by a
  thousand each (``#,##0,,"M"`` showed ``0M``);
* ``%`` (times a hundred: ``0%`` showed ``4602750%``), ``E+`` with one ``0`` before the
  point (``0.00E+00`` showed ``4,60E+04``);
* text shown as it is: quoted (``"Total: "0``), escaped (``0\\d``), and the character after
  ``_`` as a space (``0_d`` showed ``46028``, a space after it); text between the whole
  part's placeholders, the digits filled in from the right (``0\\;0`` showed ``4602;8``);
  a currency's symbol (``[$€-407] #,##0.00`` showed ``€ 46.027,50``; a language tag alone,
  ``[$-409]``, shows nothing and keeps the region's separators); a colour (``[Red]``),
  shown as nothing here;
* ``General`` (``46027,5``, and ``General "x"`` ``46027,5 x``), and a text format (``@``,
  ``"x"@``) on a number, which showed it as ``General`` does;
* a section of text alone (``"Dates"`` showed ``Dates``; ``;;;`` nothing).

The separators are the workbook's locale (:attr:`Workbook.locale
<xlsx_agent.workbook.Workbook.locale>`), as the reader's: Excel takes them from the
region, not from the code (``[$-409]#,##0.00`` showed ``46.027,50`` in the Dutch region).

**What is not measured, and refused** (``None``):

* the codes of the table Excel showed otherwise than these forms give: a fraction
  (``# ?/?``, ``## ??/16``; any ``?`` or ``/``), engineering notation (``##0.0E+0``), a
  fill (``0*d``), a condition (``[<100]0;0.00``, ``[Blue][<0]0;0``) and a numeral system
  (``[DBNum1][$-804]General``) -- seven codes -- and any code with a bracket, letter or
  token these forms do not hold;
* a negative number (every measured number was positive, so neither a second section's
  text nor a sign is measured), and zero under a code with a zero section;
* a number on a half at the last digit shown, where rounding the stored binary number,
  rounding its shortest decimal half up and rounding it half to even give different
  digits (46027.5 in ``0`` showed ``46028``, which all three give; ``0.125`` in ``0.00``
  is refused);
* more than 15 significant digits shown, an exponent ``E-``, or a ``#`` alone before the
  point of a number under one;
* in ``General``, a number whose shortest text is longer than the measured one (seven
  characters): ``General`` fits a number to its column's width, which is not measured.
"""

from __future__ import annotations

import math
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal, localcontext
from functools import lru_cache
from typing import Any

from ooxml_common import numfmt

#: The longest text ``General`` is rendered to: the measured ``46027,5``.
GENERAL_WIDTH = 7

#: The most significant digits a number is shown with.
MAX_DIGITS = 15

#: The token kinds a section rendered here may hold (:func:`ooxml_common.numfmt.tokens`).
_KINDS = frozenset({"digit", "literal", "quoted", "escaped", "pad", "color", "locale",
                    "general", "percent", "exponent"})

#: Literal characters that are no plain text in a number: a point or comma outside the
#: number, a fraction's slash.
_REFUSED_LITERALS = frozenset("./,")


def shown_text(number: "int | float", code: str | None, locale: Any) -> str | None:
    """``number`` as Excel shows it in format ``code``, with ``locale``'s separators (see
    the module's notes); ``None`` when that is not measured."""
    if isinstance(number, bool) or not isinstance(number, (int, float)) or code is None:
        return None
    if isinstance(number, float) and not math.isfinite(number):
        return None
    if number < 0:
        return None
    count, kind, tokens = _read(code)
    if number == 0 and count > 2 or kind is None or kind in numfmt.DATE_KINDS:
        return None
    if kind == "text":
        return _general(number, locale)
    return _section(number, list(tokens), locale)


@lru_cache(maxsize=256)
def _read(code: str) -> "tuple[int, str | None, tuple[numfmt.Token, ...]]":
    """A code's sections counted, its kind (``None``: unreadable) and its first section's
    tokens: once per code, not per cell."""
    parts = numfmt.sections(code)
    try:
        kind: str | None = numfmt.classify(code).kind
    except (ValueError, IndexError):
        kind = None
    return len(parts), kind, tuple(numfmt.tokens(parts[0]))


def _general(number: "int | float", locale: Any) -> str | None:
    """``number`` as ``General`` shows it: its shortest text, within the measured width."""
    if isinstance(number, int) or float(number).is_integer() and abs(number) < 1e15:
        text = str(int(number))
    else:
        text = repr(float(number))
        if "e" in text or "E" in text:
            return None
        text = text.replace(".", locale.decimal_separator)
    return text if len(text) <= GENERAL_WIDTH else None


def _section(number: "int | float", tokens: "list[numfmt.Token]", locale: Any) -> str | None:
    """``number`` through one section's tokens; ``None`` where a form is not measured."""
    if any(t.kind not in _KINDS for t in tokens):
        return None
    if any(t.kind == "digit" and t.text == "?" for t in tokens):
        return None
    if any(t.kind == "exponent" and t.text[1:] != "+" for t in tokens):
        return None
    digits = [n for n, t in enumerate(tokens) if t.kind == "digit"]
    general = [n for n, t in enumerate(tokens) if t.kind == "general"]
    percents = sum(1 for t in tokens if t.kind == "percent")
    if percents > 1 or len(general) > 1 or (digits and general):
        return None
    if general:
        if percents:
            return None
        shown = _general(number, locale)
        if shown is None:
            return None
        return _around(tokens, general[0], general[0], shown)
    if not digits:
        if percents:
            return None
        return _around(tokens, len(tokens), len(tokens), "")
    first, last = digits[0], digits[-1]
    # The scaling commas: right after the last placeholder.
    scale = 0
    while last + 1 + scale < len(tokens) and tokens[last + 1 + scale] == _COMMA:
        scale += 1
    number_text = _number(number, tokens[first:last + 1], scale, percents, locale)
    if number_text is None:
        return None
    return _around(tokens, first, last + scale, number_text)


_COMMA = numfmt.Token("literal", ",")
_POINT = numfmt.Token("literal", ".")


def _around(tokens: "list[numfmt.Token]", first: int, last: int, middle: str) -> str | None:
    """The text before token ``first`` and after token ``last``, with ``middle`` between."""
    before = _text(tokens[:first])
    after = _text(tokens[last + 1:])
    if before is None or after is None:
        return None
    return before + middle + after


def _text(tokens: "list[numfmt.Token]") -> str | None:
    """Tokens of text as Excel shows them; ``None`` for one that is not text."""
    out = []
    for token in tokens:
        kind = token.kind
        if kind == "literal":
            if token.text in _REFUSED_LITERALS:
                return None
            out.append(token.text)
        elif kind in ("quoted", "escaped"):
            out.append(token.text)
        elif kind == "pad":
            out.append(" ")
        elif kind == "locale":
            out.append(token.text.split("-", 1)[0])
        elif kind in ("color", "percent"):
            out.append("%" if kind == "percent" else "")
        else:
            return None
    return "".join(out)


def _number(number: "int | float", tokens: "list[numfmt.Token]", scale: int, percents: int,
            locale: Any) -> str | None:
    """The number's own part: ``tokens`` from its first placeholder to its last."""
    exponent = next((n for n, t in enumerate(tokens) if t.kind == "exponent"), None)
    mantissa = tokens if exponent is None else tokens[:exponent]
    power = tokens[exponent + 1:] if exponent is not None else []
    if exponent is not None and (scale or not power
                                 or any(t != numfmt.Token("digit", "0") for t in power)):
        return None
    point = next((n for n, t in enumerate(mantissa) if t == _POINT), None)
    whole = mantissa if point is None else mantissa[:point]
    fraction = [] if point is None else mantissa[point + 1:]
    if point is not None and (not fraction or any(t.kind != "digit" for t in fraction)):
        return None
    marks = [t.text for t in fraction]
    if marks != sorted(marks, key=lambda mark: mark != "0"):   # 0s first, then #s
        return None
    places, required = len(marks), marks.count("0")
    slots = [n for n, t in enumerate(whole) if t.kind == "digit"]
    if not slots:
        return None
    grouped = any(t == _COMMA for t in whole[slots[0]:slots[-1]])
    if any(t == _COMMA for t in whole[slots[-1]:]):   # a comma before the point: no form
        return None
    between = [t for t in whole[slots[0]:slots[-1]] if t.kind != "digit" and t != _COMMA]
    if between and (grouped or any(whole[n].text != "0" for n in slots)):
        return None
    if any(t.kind not in ("digit",) and t != _COMMA and t.kind not in
           ("literal", "quoted", "escaped", "pad") for t in whole):
        return None
    if any(t.kind == "literal" and t.text in _REFUSED_LITERALS and t != _COMMA for t in whole):
        return None
    zeros = sum(1 for n in slots if whole[n].text == "0")

    shift = 2 * percents - 3 * scale
    if exponent is not None:
        if len(slots) != 1 or whole[slots[0]].text != "0":
            return None
        rounded = _rounded_exponent(number, shift, places)
        if rounded is None:
            return None
        digits_text, power_value = rounded
        head, _, tail = digits_text.partition(".")
        sign = "-" if power_value < 0 else "+"
        power_text = str(abs(power_value)).rjust(len(power), "0")
        mantissa_text = _joined(head, tail, required, locale)
        if mantissa_text is None:
            return None
        return mantissa_text + tokens[exponent].text[0] + sign + power_text
    rounded_text = _rounded(number, shift, places)
    if rounded_text is None:
        return None
    head, _, tail = rounded_text.partition(".")
    if head == "0" and not zeros:
        return None   # ``#`` alone before the point of a number under one: not measured
    head = head.rjust(zeros, "0")
    if between:
        filled = _interleaved(head, whole[slots[0]:slots[-1] + 1])
        if filled is None:
            return None
        head = filled
    elif grouped:
        head = _grouped(head, locale.thousands_separator)
    return _joined(head, tail, required, locale)


def _joined(head: str, tail: str, required: int, locale: Any) -> str | None:
    """The whole part, and the decimals with the optional zeros at their end dropped."""
    if not tail:
        return head
    tail = tail.rstrip("0").ljust(required, "0")
    if not tail:
        return None   # a point with no digit after it (``0.##`` of 5): not measured
    return head + locale.decimal_separator + tail


def _grouped(head: str, separator: str) -> str:
    groups: list[str] = []
    while len(head) > 3:
        groups.insert(0, head[-3:])
        head = head[:-3]
    groups.insert(0, head)
    return separator.join(groups)


def _interleaved(head: str, whole: "list[numfmt.Token]") -> str | None:
    """The whole part's digits filled into placeholders with text between them, from the
    right; the first placeholder takes the digits left over."""
    slots = [n for n, t in enumerate(whole) if t.kind == "digit"]
    out = [""] * len(whole)
    for n, token in enumerate(whole):
        if token.kind != "digit":
            text = _text([token])
            if text is None:
                return None
            out[n] = text
    rest = head
    for n in reversed(slots[1:]):
        out[n], rest = rest[-1:], rest[:-1]
    out[slots[0]] = rest
    return "".join(out)


def _readings(number: "int | float", shift: int) -> "tuple[Decimal, Decimal]":
    """The number times ten to ``shift``: the stored binary value exactly, and its
    shortest decimal text."""
    exact = Decimal(number) if isinstance(number, int) else Decimal(float(number))
    short = Decimal(number) if isinstance(number, int) else Decimal(repr(float(number)))
    return exact.scaleb(shift), short.scaleb(shift)


def _agreed(values: "list[Decimal]", places: int) -> Decimal | None:
    """``values`` (the readings) rounded to ``places`` decimals, when every rounding gives
    the same digits."""
    step = Decimal(1).scaleb(-places)
    exact, short = values
    found = {exact.quantize(step, ROUND_HALF_EVEN), short.quantize(step, ROUND_HALF_UP),
             short.quantize(step, ROUND_HALF_EVEN)}
    return found.pop() if len(found) == 1 else None


def _rounded(number: "int | float", shift: int, places: int) -> str | None:
    with localcontext() as context:
        context.prec = 1200
        found = _agreed(list(_readings(number, shift)), places)
        if found is None or len(found.normalize().as_tuple().digits) > MAX_DIGITS \
                and found != 0:
            return None
        return f"{found:f}"


def _rounded_exponent(number: "int | float", shift: int,
                      places: int) -> "tuple[str, int] | None":
    with localcontext() as context:
        context.prec = 1200
        exact, short = _readings(number, shift)
        if short == 0:
            return f"{Decimal(0).quantize(Decimal(1).scaleb(-places)):f}", 0
        power = short.adjusted()
        found = _agreed([exact.scaleb(-power), short.scaleb(-power)], places)
        if found is None:
            return None
        if found >= 10:
            power += 1
            found = _agreed([exact.scaleb(-power), short.scaleb(-power)], places)
            if found is None:
                return None
        return f"{found:f}", power


__all__ = ["GENERAL_WIDTH", "MAX_DIGITS", "shown_text"]
