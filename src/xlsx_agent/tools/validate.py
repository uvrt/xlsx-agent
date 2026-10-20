"""``xlsx_validate_data``: a workbook checked against a declared schema -- required sheets,
tables or declared regions, keys, and per column type, required, unique, ranges, steps,
allowed values, the workbook's own drop-down lists and patterns (ROADMAP.md, "Tools").  In
the loaded core while the 6,000 budget holds (``tests/test_tools_schema.py`` keeps it).

What it does is :mod:`xlsx_agent.schema`'s (:meth:`Workbook.validate_data
<xlsx_agent.Workbook.validate_data>`); ``schema`` is that module's schema as JSON writes it.
The common strict subset has no unions, so ``min``, ``max`` and each ``enum`` value are
text as a filter's value is written (``5``, ``0.25``, ``2026-01-05``: the column's type
reads them, as ``xlsx_query``'s ``where`` does), ``key`` is always a list, and a declared
region spec is given beside ``table`` (``range``, ``name``, ``header_rows``, ``headers``)
rather than in its place.  A schema that is not one is refused (``invalid_arguments``,
``details.code`` ``bad_schema`` and ``details.path`` where), with the valid options.

**The result.**  ``ok`` (no problem at all); ``counts``: the problem cells by code, always,
over all of them; ``tables``: each table checked, its address, range, data rows and the
columns checked; ``cache``, when any cell checked holds a formula: those cells by cache
status (``cached_unverified``, ``invalidated``, ``absent``) and ``external`` -- an
``invalidated`` value checked is the one last saved; ``problems``: the problems in order, each with its code, ``at`` (a cell,
a range of cells next to each other in one column with the same problem, a table, or a
sheet as ``Stock!``), ``cells`` when more than one, ``table``, ``column``, what was
expected, and ``message``, the problem in a line for a person fixing the file.

**Bounded and paged** as ``xlsx_compare`` is: at most ``Limits.max_list_items`` problems a
page within ``Limits.max_result_chars``; ``cut`` says what was left (``problems``,
``chars``, ``tables``, and ``cells_not_listed``: problem cells past
:data:`~xlsx_agent.schema.MAX_PROBLEMS`, counted only), ``next_cursor`` gives the next page
of the same check.  ``total`` is the number of problems listed.  The check is made once
and kept for its next pages (the workbook does not change while open for reading).

**Content.**  No problem, summary or error holds a cell's content: problems carry
addresses, names, codes, types and counts (ROADMAP.md, "Design rules", 7).
"""

from __future__ import annotations

import json
from typing import Any

from ooxml_edit.tools import (CORE, Result, array, boolean, integer, number, obj, string,
                              tool)

from ..schema import CASES, COLUMN_TYPES, ValidationReport, validate
from .docformat import KIND, handler
from .query import KEPT
from .read import _MARGIN, CURSOR, DOC, _cursor, _read_cursor, _size

_COLUMN = obj({"name": string(),
               "type": string(enum=COLUMN_TYPES, optional=True),
               "required": boolean("No blanks.", optional=True),
               "unique": boolean(optional=True),
               "min": string("As filter values; needs type.", optional=True),
               "max": string(optional=True),
               "step": number("Values are min (or 0) plus whole steps: 0.5 for halves.",
                              optional=True),
               "enum": array(string(), "Allowed values.", optional=True),
               "case": string("Text in enum or list. Default exact.", enum=CASES,
                              optional=True),
               "validation_list": boolean("Allowed: the column's own drop-down list.",
                                          optional=True),
               "pattern": string("Regex text must fully match.", optional=True)},
              "A column to check, by its header.")

_TABLE = obj({"table": string("t:Products; or declare one: range, name, header_rows, "
                              "headers.", optional=True),
              "range": string(optional=True),
              "name": string(optional=True),
              "header_rows": integer(minimum=0, optional=True),
              "headers": array(string(), optional=True),
              "key": array(string(), "Key columns: unique together, never blank.",
                           optional=True),
              "columns": array(_COLUMN, optional=True)},
             "A table or region to check.")


@tool("xlsx_validate_data",
      "Check a workbook against a schema: required sheets and tables, key, and per column "
      "type, required, unique, min/max, step, enum, the column's drop-down list, pattern. "
      "Problems by code and address, never cell text; counts always; paged.",
      {"doc": DOC,
       "schema": obj({"sheets": array(string(), "Required sheets.", optional=True),
                      "tables": array(_TABLE, optional=True)},
                     "What the workbook must hold."),
       "cursor": CURSOR},
      group=CORE, kind=KIND)
@handler
def xlsx_validate_data(call: Any, doc: str, schema: dict,
                       cursor: str | None = None) -> Result:
    """``xlsx_validate_data``'s workbook handler (see the module's notes)."""
    workbook = call.document
    # A cursor belongs to one check: the document as it is, and the schema.
    ident = json.dumps(["validate", doc, call.entry.version, schema], ensure_ascii=False,
                       sort_keys=True, default=str)
    report = _report(workbook, ident, schema)
    total = len(report.problems)
    start, _ = _read_cursor(cursor, ident, total, 1, doc)
    chars = call.limits.max_result_chars
    limit = call.limits.max_list_items
    data: dict[str, Any] = {"doc": doc, "ok": report.ok, "counts": dict(report.counts)}
    data["tables"] = [t.as_dict() for t in report.tables[:limit]]
    if report.cache:
        data["cache"] = dict(report.cache)
    data["problems"] = []
    budget = chars - _MARGIN - _size({**data, "cut": _CUT_ROOM}) \
        - len(_summary(doc, report, 1, total, total, True))
    used = 0
    index = start
    full = False
    while index < total and len(data["problems"]) < limit:
        shaped = report.problems[index].as_dict()
        size = _size(shaped) + 1
        if used + size > budget and data["problems"]:
            full = True
            break
        data["problems"].append(shaped)
        used += size
        index += 1

    out_cut: dict[str, Any] = {}
    next_cursor = None
    if index < total:
        out_cut["problems"] = total - index
        if full:
            out_cut["chars"] = True
        next_cursor = _cursor(ident, index, 0)
    if len(report.tables) > len(data["tables"]):
        out_cut["tables"] = len(report.tables) - len(data["tables"])
    if report.cut:
        out_cut["cells_not_listed"] = report.cut
    if out_cut:
        data["cut"] = out_cut
    summary = _summary(doc, report, start + 1, index, total, next_cursor is not None)
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor)


def _report(workbook: Any, ident: str, schema: Any) -> ValidationReport:
    """The check, kept with the workbook's last queries for its next pages."""
    kept = workbook._queries
    found = kept.get(ident)
    if found is None:
        found = validate(workbook, schema)
        kept[ident] = found
        while len(kept) > KEPT:
            kept.popitem(last=False)
    else:
        kept.move_to_end(ident)
    return found


#: A ``cut`` at its longest, to keep room for it in a page.
_CUT_ROOM = {"problems": 10 ** 9, "chars": True, "tables": 10 ** 5, "cells_not_listed": 10 ** 9}


def _summary(doc: str, report: ValidationReport, first: int, last: int, total: int,
             more: bool) -> str:
    text = f"{doc}: {len(report.tables)} tables checked"
    if report.ok:
        return text + "; no problems"
    counts = ", ".join(f"{n} {code.replace('_', ' ')}" for code, n in report.counts.items())
    text += f"; problem cells: {counts}; problems {first}-{last} of {total}"
    return text + "; more with next_cursor" if more else text


TOOLS = [xlsx_validate_data]
