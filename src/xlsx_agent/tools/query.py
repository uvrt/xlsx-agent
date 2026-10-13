"""The query tool, ``xlsx_query``: a table's or region's rows that meet some filters, sorted,
with the columns asked for (ROADMAP.md, "Tools").  In the loaded core, beside ``xlsx_read``.

What it does is :mod:`xlsx_agent.query`'s: the filters (``eq``, ``in``, ``gte``, ``lte``,
``contains``, ``words_gt``, ``blank``, ``not_blank``, combined with *and*), each checked
against its column's type before any row is given -- a test the type does not take, or a value not of
that type, is refused naming the type, with the tests it takes as ``valid_options`` -- and
the sort.  A filter's ``value`` is text as JSON writes the value (``5``, ``2.5``, ``true``,
``2026-01-05``): the common strict subset has no unions, and the column's type decides how
it is read.  ``eq`` and ``in`` match text exactly (case counts); ``contains`` ignores case.

**Bounded and paged** as ``xlsx_read`` is (:mod:`.read`: its page, its cursor, its cuts):
``limit`` rows a page (by default ``Limits.max_list_items``, at most
:data:`~.read.READ_ROWS`) within ``Limits.max_result_chars``, a row too wide for what is
left split across pages, long text cut with its length.  ``total`` is the count of all
matching rows; ``cut.rows`` the matching rows not yet given; ``next_cursor``, passed back
with the same query, the next page (a cursor of another query is refused).  ``types``
gives the type each filtered or sorted column was taken as; ``not_compared``, per filtered
column, the cells of another type a comparing test never matches, counted with their first
addresses.

**Content.**  Rows are a result; the summary names the region and counts rows, and errors
name columns, tests and types: never a filter's value, which may be a cell's.
"""

from __future__ import annotations

import json
from typing import Any

from ooxml_edit.tools import (CORE, Result, ToolError, array, boolean, integer, obj, string,
                              tool)

from ..address import TableAddress, parse_address
from ..query import OPS, matches, read_rows, select
from .docformat import KIND, handler
from .read import (_CUT_ROOM, _FIRST, _MARGIN, CURSOR, DOC, READ_ROWS, TEXT_SHOWN, _Annotations,
                   _cursor, _forms, _page, _read_cursor, _sheet, _size)

#: How many queries' matches a workbook keeps for their next pages, the latest.
KEPT = 4


#: The filters, as ``xlsx_query`` and ``xlsx_derive`` take them.
WHERE = array(obj({"column": string(),
                   "op": string(enum=OPS),
                   "value": string(optional=True),
                   "values": array(string(), "For in.", optional=True)}),
              "Filters, all met. value as JSON writes it: 12.5, true, dates ISO "
              "(2026-01-05), percent as the fraction. gte/lte: numbers, dates, "
              "times. contains: text, case ignored; eq/in match exactly.",
              optional=True)

#: The sort keys, as ``xlsx_query`` and ``xlsx_derive`` take them.
SORT = array(obj({"column": string(), "desc": boolean(optional=True)}),
             "Sort keys, first decides first. Other types, then blanks, last.",
             optional=True)


@tool("xlsx_query",
      "Rows of a table or region that meet every filter, sorted, each cell typed with its "
      "address; total is the count of all matches. A filter on a column of another type "
      "is refused, naming the type. Paged as xlsx_read.",
      {"doc": DOC,
       "table": string("A table or region (t:Products, t:Stock!A1)."),
       "columns": array(string(), "Columns to give, by name. Default all.", optional=True),
       "where": WHERE,
       "sort": SORT,
       "limit": integer(f"Rows per page. Default 50, at most {READ_ROWS}.", minimum=1,
                        maximum=READ_ROWS, optional=True),
       "cursor": CURSOR},
      group=CORE, kind=KIND)
@handler
def xlsx_query(call: Any, doc: str, table: str, columns: list[str] | None = None,
               where: list[dict] | None = None, sort: list[dict] | None = None,
               limit: int | None = None, cursor: str | None = None) -> Result:
    """``xlsx_query``'s workbook handler (see the module's notes)."""
    workbook = call.document
    address = parse_address(table)
    if not isinstance(address, TableAddress) or not address.is_whole:
        raise ToolError("invalid_arguments", "xlsx_query takes a table or region (t:Products); "
                        "name its columns in columns", field="table")
    region = workbook.region(address)
    label = region.address.text
    names = select(region, columns)
    where = [{k: v for k, v in item.items() if v is not None} for item in where or ()]
    sort = [{k: v for k, v in item.items() if v is not None} for item in sort or ()]
    # A cursor belongs to one query: its region (as it is now), columns, filters and sort.
    key = json.dumps([label, region.range.text, region.columns, names, where, sort],
                     ensure_ascii=False, sort_keys=True)
    found = _matches(workbook, region, key, where, sort, names)
    total = len(found.rows)
    start_row, start_column = _read_cursor(cursor, key, total, len(names), label)
    rows_limit = min(limit or call.limits.max_list_items, READ_ROWS)
    chars = call.limits.max_result_chars

    sheet = _sheet(workbook, workbook.sheet_of(region.range))
    data: dict[str, Any] = {"target": label, "sheet": _forms(workbook, sheet),
                            "range": region.range.text}
    names_cut = 0
    if _size(list(names)) <= chars // 4:
        data["columns"] = list(names)
    else:  # a region too wide for its names on every page: describe and xlsx_profile
        names_cut = len(names)
    if found.types:
        data["types"] = found.types
    if found.not_compared:
        data["not_compared"] = {n: c.as_dict() for n, c in found.not_compared.items()}
    data["rows"] = []
    if not total:
        return Result(summary=_summary(label, 1, 0, 0, region.rows, False), data=data, total=0)

    page = found.rows[start_row:start_row + rows_limit]
    budget = chars - _MARGIN - _size({**data, "cut": _CUT_ROOM}) \
        - len(_summary(label, 1, total, total, region.rows, True))
    annotate = _Annotations(workbook, sheet, comments=False, hyperlinks=False)
    shown, cut_text, stop, full = _page(read_rows(workbook, region, names, page, kept=found.values), start_row,
                                        start_column, budget, annotate, False, TEXT_SHOWN)
    if stop is None and start_row + len(page) < total:
        stop = (start_row + len(page), 0)

    data["rows"] = shown
    cut: dict[str, Any] = {}
    next_cursor = None
    if stop is not None:
        cut["rows"] = total - stop[0]
        if stop[1]:
            cut["cells"] = len(names) - stop[1]
        if full:
            cut["chars"] = True
        next_cursor = _cursor(key, *stop)
    if cut_text:
        cut["text"] = {"count": len(cut_text), "first": cut_text[:_FIRST]}
    if names_cut:
        cut["column_names"] = names_cut
    if cut:
        data["cut"] = cut
    summary = _summary(label, start_row + 1, start_row + len(shown), total, region.rows,
                       next_cursor is not None)
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor)


def _matches(workbook: Any, region: Any, key: str, where: list, sort: list,
             names: tuple[str, ...]) -> Any:
    """The query's matches, kept for its next pages (the workbook does not change while it
    is open for reading): the last :data:`KEPT` queries."""
    kept = workbook._queries
    found = kept.get(key)
    if found is None:
        found = matches(workbook, region, where=where, sort=sort, columns=names)
        kept[key] = found
        while len(kept) > KEPT:
            kept.popitem(last=False)
    else:
        kept.move_to_end(key)
    return found


def _summary(target: str, first: int, last: int, total: int, rows: int, more: bool) -> str:
    if total == 0 or last < first:
        return f"{target}: no rows of {rows} match"
    text = f"{target}: matching rows {first}-{last} of {total} (of {rows})"
    return text + "; more with next_cursor" if more else text


TOOLS = [xlsx_query]
