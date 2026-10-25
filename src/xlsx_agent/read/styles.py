"""Number formats: which of a workbook's cells hold a percent, a date or a time.

A cell's ``s`` is an index into ``styles.xml``'s ``cellXfs``; that ``xf``'s ``numFmtId`` is
either one of the workbook's own formats (a ``numFmt`` in ``numFmts``, with its
``formatCode``) or a built-in id whose code is implied, not written.  The built-in ids are
ECMA-376 Part 1, §18.8.30 (``numFmt``), and only that list is used here:
:data:`BUILTIN_FORMATS` is its table for all languages, :data:`LANGUAGE_FORMATS` its tables
for ``zh-TW``, ``zh-CN``, ``ja-JP``, ``ko-KR`` and ``th-TH``.  The standard says that "some
of these Ids can be interpreted differently, depending on the UI language of the
implementing application", and that ids it does not list (5 to 8, say) follow the
``formatCode`` the file gives.  So:

* an id in the all-languages table is that table's code, whatever the language;
* an id in the language tables (27 to 36 and 50 to 58 for the four East Asian languages,
  59 to 81 for Thai) is read with the language the caller's locale names, when it is one
  of those five (:func:`~xlsx_agent.read.types.format_language`).  With none of them, an
  id whose kind (date or time) is the same in every language that lists it takes that
  kind; one whose kind differs between languages (34, 35, 52, 53, 55, 56) is
  ``"ambiguous"``, and a number in it is left a number, never guessed a date or a time.
  Measured (Excel 16.106 Mac, Dutch UI, 2026-10-13): ids 27 to 36 and 50 to 58 show as
  id 14 (a date) or id 21 (a time) under both the nl-BE and the en-US region;
* any other id without a code in the file is ``"unknown"``: its numbers stay numbers.

An Excel probe in E1 confirms the list on Excel itself (ROADMAP.md, "Values and types"),
and a difference wins over the standard.

A format code is classified by its tokens (:func:`classify_format`), as §18.8.31
(``numFmts``) describes them, with :mod:`ooxml_common.numfmt`, the one reading of a code
the libraries share, measured in Excel: date and time tokens outside quotes, escapes and
brackets (``[h]``, ``[mm]``, ``[ss]`` being elapsed time), ``m`` read as minutes right
after an hour or right before a second, ``%`` outside the same; ``General``, ``@``.  Only
the first section decides: it is the one a positive number is shown with (measured, Excel
for Mac 16.106: ``0;[Red]d`` shows a positive serial as a number).  No
display string is made from a code (ROADMAP.md: none before 0.3).
"""

from __future__ import annotations

from dataclasses import dataclass

from ooxml_common import numfmt as _numfmt
from ooxml_edit.xml import Element

#: ECMA-376 Part 1, §18.8.30: the built-in number formats, "All Languages".
BUILTIN_FORMATS: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

_ZH_TW = {
    27: '[$-404]e/m/d', 28: '[$-404]e"年"m"月"d"日"', 29: '[$-404]e"年"m"月"d"日"',
    30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"', 32: 'hh"時"mm"分"', 33: 'hh"時"mm"分"ss"秒"',
    34: '上午/下午hh"時"mm"分"', 35: '上午/下午hh"時"mm"分"ss"秒"', 36: '[$-404]e/m/d',
    50: '[$-404]e/m/d', 51: '[$-404]e"年"m"月"d"日"', 52: '上午/下午hh"時"mm"分"',
    53: '上午/下午hh"時"mm"分"ss"秒"', 54: '[$-404]e"年"m"月"d"日"', 55: '上午/下午hh"時"mm"分"',
    56: '上午/下午hh"時"mm"分"ss"秒"', 57: '[$-404]e/m/d', 58: '[$-404]e"年"m"月"d"日"',
}
_ZH_CN = {
    27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: 'm-d-yy',
    31: 'yyyy"年"m"月"d"日"', 32: 'h"时"mm"分"', 33: 'h"时"mm"分"ss"秒"',
    34: '上午/下午h"时"mm"分"', 35: '上午/下午h"时"mm"分"ss"秒"', 36: 'yyyy"年"m"月"',
    50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
    54: 'm"月"d"日"', 55: '上午/下午h"时"mm"分"', 56: '上午/下午h"时"mm"分"ss"秒"',
    57: 'yyyy"年"m"月"', 58: 'm"月"d"日"',
}
_JA_JP = {
    27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"',
    30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"', 32: 'h"時"mm"分"', 33: 'h"時"mm"分"ss"秒"',
    34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: '[$-411]ge.m.d', 50: '[$-411]ge.m.d',
    51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
    54: '[$-411]ggge"年"m"月"d"日"', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"',
    57: '[$-411]ge.m.d', 58: '[$-411]ggge"年"m"月"d"日"',
}
_KO_KR = {
    27: 'yyyy"年" mm"月" dd"日"', 28: 'mm-dd', 29: 'mm-dd', 30: 'mm-dd-yy',
    31: 'yyyy"년" mm"월" dd"일"', 32: 'h"시" mm"분"', 33: 'h"시" mm"분" ss"초"',
    34: 'yyyy-mm-dd', 35: 'yyyy-mm-dd', 36: 'yyyy"年" mm"月" dd"日"',
    50: 'yyyy"年" mm"月" dd"日"', 51: 'mm-dd', 52: 'yyyy-mm-dd', 53: 'yyyy-mm-dd',
    54: 'mm-dd', 55: 'yyyy-mm-dd', 56: 'yyyy-mm-dd', 57: 'yyyy"年" mm"月" dd"日"',
    58: 'mm-dd',
}
_TH_TH = {
    59: 't0', 60: 't0.00', 61: 't#,##0', 62: 't#,##0.00', 67: 't0%', 68: 't0.00%',
    69: 't# ?/?', 70: 't# ??/??', 71: 'ว/ด/ปปปป', 72: 'ว-ดดด-ปป', 73: 'ว-ดดด',
    74: 'ดดด-ปป', 75: 'ช:นน', 76: 'ช:นน:ทท', 77: 'ว/ด/ปปปป ช:นน', 78: 'นน:ทท',
    79: '[ช]:นน:ทท', 80: 'นน:ทท.0', 81: 'd/m/bb',
}

#: ECMA-376 Part 1, §18.8.30: the built-in formats whose code depends on the UI language,
#: by language tag.
LANGUAGE_FORMATS: dict[str, dict[int, str]] = {
    "zh-TW": _ZH_TW, "zh-CN": _ZH_CN, "ja-JP": _JA_JP, "ko-KR": _KO_KR, "th-TH": _TH_TH,
}

#: Every kind a number format is classified as.
KINDS = ("general", "number", "percent", "date", "time", "datetime", "text", "unknown",
         "ambiguous")

#: The kinds whose numbers are read as dates and times.
DATE_KINDS = ("date", "time", "datetime")

@dataclass(frozen=True)
class NumberFormat:
    """A number format as a cell's style gives it.

    ``id`` is the ``numFmtId``; ``code`` the format code (the file's, or the built-in one's
    from §18.8.30; ``None`` when the id is neither); ``kind`` one of :data:`KINDS`;
    ``elapsed`` whether a time is elapsed time (``[h]:mm:ss``); ``lcid`` the locale id of a
    ``[$-409]``-style tag in the code, if any; ``builtin`` whether the code is implied by
    the id rather than written in the file."""

    id: int
    code: str | None
    kind: str
    elapsed: bool = False
    lcid: int | None = None
    builtin: bool = False

    @property
    def is_date(self) -> bool:
        return self.kind in DATE_KINDS


GENERAL = NumberFormat(0, "General", "general", builtin=True)


def classify_format(code: str) -> tuple[str, bool, int | None]:
    """``(kind, elapsed, lcid)`` for a format code: ``kind`` is ``"general"``,
    ``"number"``, ``"percent"``, ``"date"``, ``"time"``, ``"datetime"`` or ``"text"``.

    The code's first section decides (the one a positive number is shown with), read by
    :func:`ooxml_common.numfmt.classify`, the reading the libraries share and hold to
    Excel's answers (``numfmt.MEASURED``): ``y``, ``d``, ``e``, ``g``, ``b`` (and the
    Thai ``ว ด ป``) are dates; ``h``, ``s``, ``AM/PM``, ``A/P`` (and ``ช น ท``) are
    times; ``m`` is minutes right after an hour or right before a second, else a month;
    ``[h]``, ``[mm]``, ``[ss]`` are elapsed time; ``[$-409]`` gives ``lcid``."""
    found = _numfmt.classify(code)
    return found.kind, found.elapsed, found.lcid


def builtin_format(number_format_id: int, language: str | None = None) -> NumberFormat:
    """The built-in format with id ``number_format_id`` (§18.8.30), read in ``language``
    (one of :data:`LANGUAGE_FORMATS`' tags, or ``None``) as the module's notes say."""
    if number_format_id in BUILTIN_FORMATS:
        code = BUILTIN_FORMATS[number_format_id]
        kind, elapsed, lcid = classify_format(code)
        return NumberFormat(number_format_id, code, kind, elapsed, lcid, builtin=True)
    tables = LANGUAGE_FORMATS
    if language is not None:
        table = tables.get(language)
        code = table.get(number_format_id) if table else None
        if code is None:
            return NumberFormat(number_format_id, None, "unknown", builtin=True)
        kind, elapsed, lcid = classify_format(code)
        return NumberFormat(number_format_id, code, kind, elapsed, lcid, builtin=True)
    found = [(classify_format(table[number_format_id]), table[number_format_id])
             for table in tables.values() if number_format_id in table]
    if not found:
        return NumberFormat(number_format_id, None, "unknown", builtin=True)
    kinds = {kind for (kind, _, _), _ in found}
    if len(kinds) > 1:
        return NumberFormat(number_format_id, None, "ambiguous", builtin=True)
    (kind, elapsed, _), code = found[0]
    # The code differs by language even where the kind does not; none is chosen.
    codes = {code for _, code in found}
    return NumberFormat(number_format_id, code if len(codes) == 1 else None, kind,
                        elapsed, None, builtin=True)


class Styles:
    """``styles.xml``'s number formats: each ``cellXfs`` entry's format, found once.

    ``root`` is the parsed style part (``None`` when the workbook has none: every cell is
    then ``General``); ``language`` reads the language-dependent built-in ids."""

    def __init__(self, root: Element | None, language: str | None = None) -> None:
        self.language = language
        self._codes: dict[int, str] = {}
        self._xf_ids: list[int | None] = []
        self._cache: dict[int, NumberFormat] = {}
        if root is None:
            return
        for node in root:
            local = node.tag.rpartition("}")[2] if isinstance(node.tag, str) else ""
            if local == "numFmts":
                for fmt in node:
                    number = _int(fmt.get("numFmtId"))
                    code = fmt.get("formatCode")
                    if number is not None and code is not None:
                        self._codes[number] = code
            elif local == "cellXfs":
                for xf in node:
                    if isinstance(xf.tag, str) and xf.tag.rpartition("}")[2] == "xf":
                        self._xf_ids.append(_int(xf.get("numFmtId")) or 0)

    def format_id(self, style: int) -> int:
        """The ``numFmtId`` of ``cellXfs[style]``; 0 (General) past the end of the list."""
        if 0 <= style < len(self._xf_ids):
            return self._xf_ids[style] or 0
        return 0

    def format(self, style: int) -> NumberFormat:
        """The number format of a cell whose ``s`` is ``style``."""
        cached = self._cache.get(style)
        if cached is None:
            cached = self._cache[style] = self.format_by_id(self.format_id(style))
        return cached

    def format_by_id(self, number_format_id: int) -> NumberFormat:
        """The format with ``numFmtId`` ``number_format_id``: the file's code when it gives
        one, else the built-in (§18.8.30)."""
        code = self._codes.get(number_format_id)
        if code is not None:
            kind, elapsed, lcid = classify_format(code)
            return NumberFormat(number_format_id, code, kind, elapsed, lcid)
        return builtin_format(number_format_id, self.language)


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None
