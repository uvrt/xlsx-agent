"""The xlsx handlers of the shared tools (``ooxml_edit.tools.shared``): one definition each
for every format, dispatched by the document's kind.

``open_document``, ``close_document`` and ``read_blob`` are the session's own
(``shared.SESSION_TOOLS``): they work for a workbook as for any document, through the
xlsx :data:`~xlsx_agent.tools.FORMAT` (its ``detect``, ``open``, ``problems`` and
``summary``).  This module adds the workbook's ``describe``: the catalogue
(:mod:`xlsx_agent.catalogue`), and its ``find_text``.  ``save_document``, ``undo`` and
``batch`` (E1) and ``new_document`` (E9) come with their milestones.

**find_text** is :meth:`Workbook.find_text <xlsx_agent.workbook.Workbook.find_text>`
(:mod:`xlsx_agent.find`): ``text`` as written (case counts), or a Python regex
(``regex``), in text cells, table and region names, table headers, comments and defined
names; ``range`` a sheet (``Stock!``), range or table.  Numbers, dates and other values
that are not text, and formula text (the shared tool has no option for it), are not
searched.  Each match is ``address``, ``kind``, ``context`` (:data:`~xlsx_agent.find.CONTEXT`
characters either side, the match cut at :data:`~xlsx_agent.find.MATCH_SHOWN`), and
``cell``, ``range``, ``name`` or ``hidden`` where they apply; ``sheets`` names each
sheet of the page both ways; ``kinds`` counts every match by kind.  ``Limits.max_list_items``
matches a page within ``Limits.max_result_chars``; ``total`` counts them all and
``next_cursor``, passed back with the same text, regex and range, gives the next page.
``slides`` and ``stories`` are a deck's and a document's: refused, naming ``range``.

**Content.**  Matches are a result.  The summary counts them; errors name the field and
the problem, never the text searched for or found.
"""

from __future__ import annotations

import json
import re
from typing import Any

from ooxml_edit.tools import Result, ToolError, shared

from ..find import KINDS
from .docformat import KIND, handler
from .read import _MARGIN, _cursor, _forms, _parse_cursor, _size

#: The session tools a workbook uses so far; ``undo`` and ``batch`` wait for edits (E1).
SESSION_NAMES = ("open_document", "close_document", "read_blob")


# -- S16 describe ------------------------------------------------------------------------------


@shared.handler("describe", kind=KIND)
@handler
def describe(call: Any, doc: str) -> Result:
    """The shared ``describe``'s xlsx handler: the workbook's catalogue, within its token
    target (:data:`~xlsx_agent.catalogue.DEFAULT_TOKENS`)."""
    entry = call.entry
    data = {"name": entry.name, **call.document.catalogue()}
    return Result(summary=f"Described {entry.name}", data=data)


# -- S6 find_text ------------------------------------------------------------------------------


@shared.handler("find_text", kind=KIND)
@handler
def find_text(call: Any, doc: str, text: str, regex: bool = False, slides: list | None = None,
              range: str | None = None, stories: str | None = None,  # noqa: A002 (shared)
              cursor: str | None = None) -> Result:
    """The shared ``find_text``'s xlsx handler (see the module's notes)."""
    if slides is not None:
        raise ToolError("invalid_arguments", "slides is for decks; a workbook takes range "
                        "(a sheet, range or table)", field="slides")
    if stories is not None:
        raise ToolError("invalid_arguments", "stories is for decks and documents; a workbook "
                        "takes range (a sheet, range or table)", field="stories")
    pattern = _pattern(text, regex)
    workbook = call.document
    key = json.dumps([text, bool(regex), range], ensure_ascii=False)
    start = 0
    if cursor:
        start, column = _parse_cursor(cursor, key, "find_text")
        if column:
            raise ToolError("invalid_arguments", "cursor is not one this tool gave",
                            field="cursor")
    limit = call.limits.max_list_items
    budget = call.limits.max_result_chars - _MARGIN - _size(_ROOM)
    matches: list[dict] = []
    sheets: dict[str, dict] = {}
    kinds = dict.fromkeys(KINDS, 0)
    used = 0
    stop: int | None = None
    total = 0
    for found in workbook.find_text(pattern, within=range):
        index = total
        total += 1
        kinds[found.kind] += 1
        if index < start or stop is not None:
            continue
        if len(matches) == limit:
            stop = index
            continue
        item = found.as_dict()
        sheet = _sheet_of(workbook, found)
        new_sheet = sheet is not None and sheet.name not in sheets
        size = _size(item) + 1 + (_size(_forms(workbook, sheet)) + 1 if new_sheet else 0)
        if used + size > budget and matches:
            stop = index
            continue
        if new_sheet:
            sheets[sheet.name] = _forms(workbook, sheet)
        matches.append(item)
        used += size
    if start and start >= total:
        raise ToolError("invalid_arguments", "cursor is past the end of the matches",
                        field="cursor")
    data: dict[str, Any] = {"sheets": list(sheets.values()), "matches": matches,
                            "kinds": {k: n for k, n in kinds.items() if n}}
    if range is not None:
        data["range"] = range
    next_cursor = _cursor(key, stop, 0) if stop is not None else None
    return Result(summary=_summary(start, len(matches), total, next_cursor is not None),
                  data=data, total=total, next_cursor=next_cursor)


#: ``kinds`` and ``range`` at their longest, kept room for in a page.
_ROOM = {"sheets": [], "matches": [], "kinds": dict.fromkeys(KINDS, 1048576 * 16384),
         "range": "'" + "x" * 31 + "'!A1:XFD1048576"}


def _pattern(text: str, regex: bool) -> re.Pattern:
    if not text:
        raise ToolError("invalid_arguments", "text is empty", field="text")
    if not regex:
        return re.compile(re.escape(text))
    try:
        return re.compile(text)
    except re.error as error:
        # where only: some of re.error's messages quote a piece of the pattern
        raise ToolError("invalid_arguments", "text is not a Python regular expression (at "
                        f"character {(error.pos or 0) + 1})", field="text") from None


def _sheet_of(workbook: Any, found: Any) -> Any:
    """The sheet a match is on, for the page's ``sheets``; ``None`` for a name that is on
    none."""
    where = found.cell or found.range or (found.address if found.kind in ("cell", "comment")
                                          else None)
    if where is None:
        return None
    try:
        return workbook.sheet_of(where)
    except (KeyError, ValueError):
        return None


def _summary(start: int, shown: int, total: int, more: bool) -> str:
    if not total:
        return "0 matches"
    if not shown:
        return f"{total} match(es)"
    text = f"{total} match(es); {start + 1}-{start + shown} given"
    return text + "; more with next_cursor" if more else text


TOOLS = [*[tool for tool in shared.SESSION_TOOLS if tool.name in SESSION_NAMES], describe,
         find_text]
