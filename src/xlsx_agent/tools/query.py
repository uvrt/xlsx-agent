"""The query tool, ``xlsx_query``: a table's or region's rows that meet some filters, sorted,
with the columns asked for (ROADMAP.md, "Tools").  In the loaded core, beside ``xlsx_read``.

What it does is :mod:`xlsx_agent.query`'s: the filters (``eq``, ``in``, ``gte``, ``lte``,
``contains``, ``contains_any``, ``words_gt``, ``blank``, ``not_blank``, combined with *and*;
a filter's ``loose`` compares text ignoring case and white space), each checked
against its column's type before any row is given -- a test the type does not take, or a value not of
that type, is refused naming the type, with the tests it takes as ``valid_options`` -- and
the sort.  A filter's ``value`` is text as JSON writes the value (``5``, ``2.5``, ``true``,
``2026-01-05``): the common strict subset has no unions, and the column's type decides how
it is read.  ``eq`` and ``in`` match text exactly (case counts); ``contains`` ignores case.

**Bounded and paged** as ``xlsx_read`` is (:mod:`.read`: its page, its cursor, its cuts):
``limit`` rows a page (by default ``Limits.max_list_items``, at most
:data:`~.read.READ_ROWS`) within ``Limits.max_result_chars``, a row too wide for what is
left split across pages, text longer than ``text_chars`` (by default
:data:`~.read.TEXT_SHOWN`) cut with its length, and ``cut.text`` giving the call that has
it in full (:func:`~.read.text_cut`).  ``total`` is the count of all
matching rows; ``cut.rows`` the matching rows not yet given; ``next_cursor``, passed back
with the same query, the next page (a cursor of another query is refused).  ``types``
gives the type each filtered or sorted column was taken as; ``not_compared``, per filtered
column, the cells of another type a comparing test never matches, counted with their first
addresses; ``loosened``, per column a ``loose`` filter loosens, the matching rows that match
only loosely (round 40, R43).  A query that matches no row, whose text filters would match
some with ``loose``, says how many and which filters (a warning: the next call is that one).

**A dataset** (``d:by_region``; round 10, O5) is queried as a table is, no ``doc``
needed: its values as it reports them, each with its ``d:`` address
(``d:by_region/#2/total``), and ``dataset`` says where it came from and whether it is
stale, as ``xlsx_read`` of it does.

**Content.**  Rows are a result; the summary names the region and counts rows, and errors
name columns, tests and types: never a filter's value, which may be a cell's.
"""

from __future__ import annotations

import json
from typing import Any

from ooxml_edit.tools import (CORE, DATASET_PREFIX, Result, ToolError, array, boolean, integer,
                              obj, string, tool)

from ..address import TableAddress, parse_address
from ..clean import UNMATCHED, unmatched_warning
from ..query import LOOSENED, OPS, Matches, dataset_matches, matches, read_rows, select
from .docformat import handler
from .read import (_MARGIN, CALL_CURSOR, READ_ROWS, TEXT_ALIASES, TEXT_CHARS, TEXT_SHOWN,
                   _Annotations, _cursor, _forms, _page, _read_cursor, _sheet, _size, again,
                   cleaned_count, cut_room, text_cut)

#: How many queries' matches a workbook keeps for their next pages, the latest.
KEPT = 4


#: The filters, as ``xlsx_query`` and ``xlsx_derive`` take them.
WHERE = array(obj({"column": string(),
                   "op": string(enum=OPS),
                   "value": string(optional=True),
                   "values": array(string(), "For in, contains_any.", optional=True),
                   "loose": boolean("Text ignoring case and all white space.",
                                    optional=True)}),
              "Filters, all met. value as JSON writes it: 12.5, true, dates ISO "
              "(2026-01-05), percent as the fraction. gte/lte: numbers, dates, "
              "times. contains: text, case ignored; eq/in match exactly; "
              "contains_any: any of values.",
              optional=True)

#: The sort keys, as ``xlsx_query`` and ``xlsx_derive`` take them.
SORT = array(obj({"column": string(), "desc": boolean(optional=True)}),
             "Sort keys, first decides first. Other types, then blanks, last.",
             optional=True)


@tool("xlsx_query",
      "Rows of a table, region or dataset that meet every filter, sorted, each cell typed "
      "with its address; total is the count of all matches. Paged as xlsx_read.",
      {"doc": string("Document id; not needed for a dataset (d:).", optional=True),
       "table": string("A table or region (t:Products, t:Stock!A1) or dataset (d:by_region)."),
       "columns": array(string(), "Columns to give, by name. Default all.", optional=True),
       "where": WHERE,
       "sort": SORT,
       "limit": integer(f"Rows per page. Default 50, at most {READ_ROWS}.", minimum=1,
                        maximum=READ_ROWS, optional=True),
       "text_chars": TEXT_CHARS,
       "cursor": CALL_CURSOR},
      group=CORE, datasets=("table",), aliases=TEXT_ALIASES)
@handler
def xlsx_query(call: Any, table: str, doc: str | None = None, columns: list[str] | None = None,
               where: list[dict] | None = None, sort: list[dict] | None = None,
               limit: int | None = None, cursor: str | None = None,
               text_chars: int | None = None) -> Result:
    """``xlsx_query``'s handler, for a workbook's table or region, or a dataset (see the
    module's notes)."""
    from .derive import workbook_of

    where = [{k: v for k, v in item.items() if v is not None} for item in where or ()]
    sort = [{k: v for k, v in item.items() if v is not None} for item in sort or ()]
    if isinstance(table, str) and table.startswith(DATASET_PREFIX):
        return _query_dataset(call, table, columns, where, sort, limit, cursor, text_chars)
    workbook = workbook_of(call, doc)
    address = parse_address(table)
    if not isinstance(address, TableAddress) or not address.is_whole:
        raise ToolError("invalid_arguments", "xlsx_query takes a table or region (t:Products); "
                        "name its columns in columns", field="table")
    region = workbook.region(address)
    label = region.address.text
    names = select(region, columns)
    # A cursor belongs to one query: its region (as it is now), columns, filters and sort.
    key = json.dumps([label, region.where, region.columns, names, where, sort],
                     ensure_ascii=False, sort_keys=True)
    found = _matches(workbook, region, key, where, sort, names)
    total = len(found.rows)
    start_row, start_column = _read_cursor(cursor, key, total, len(names), label)
    rows_limit = min(limit or call.limits.max_list_items, READ_ROWS)
    chars = call.limits.max_result_chars

    data: dict[str, Any] = {"target": label}
    if region.span is None:     # a table over several sheets: each value names its own
        data["sheet"] = _forms(workbook, _sheet(workbook, workbook.sheet_of(region.range)))
    data["range"] = region.where
    names_cut = 0
    if _size(list(names)) <= chars // 4:
        data["columns"] = list(names)
    else:  # a region too wide for its names on every page: describe and xlsx_profile
        names_cut = len(names)
    if found.types:
        data["types"] = found.types
    if found.not_compared:
        data["not_compared"] = {n: c.as_dict() for n, c in found.not_compared.items()}
    if found.loosened:
        data["loosened"] = {n: c.as_dict() for n, c in found.loosened.items()}
    data["rows"] = []
    if not total:
        loosely = _loosely(where, lambda w: matches(workbook, region, where=w))
        return Result(summary=_summary(label, 1, 0, 0, region.rows, False), data=data, total=0,
                      warnings=list(found.notes) + loosely)

    page = found.rows[start_row:start_row + rows_limit]
    budget = chars - _MARGIN - _size({**data, "cut": cut_room(again(call))}) \
        - len(_summary(label, 1, total, total, region.rows, True))
    annotate = _Annotations(workbook, [], comments=False, hyperlinks=False)
    cleaned: list[str] = []
    unmatched: list[str] = []
    shown, cut_text, stop, full = _page(read_rows(workbook, region, names, page, kept=found.values), start_row,
                                        start_column, budget, annotate, False,
                                        text_chars or TEXT_SHOWN, cleaned, unmatched)
    if stop is None and start_row + len(page) < total:
        stop = (start_row + len(page), 0)

    data["rows"] = shown
    if cleaned:
        data["cleaned"] = cleaned_count(cleaned)
    warnings = list(found.notes)
    if unmatched:
        data[UNMATCHED] = cleaned_count(unmatched)
        warnings.append(unmatched_warning(len(unmatched), unmatched[0]))
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
        cut["text"] = text_cut(cut_text, shown, again(call), chars, warnings)
    if names_cut:
        cut["column_names"] = names_cut
    if cut:
        data["cut"] = cut
    summary = _summary(label, start_row + 1, start_row + len(shown), total, region.rows,
                       next_cursor is not None)
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor,
                  warnings=warnings)


def _query_dataset(call: Any, target: str, columns: list[str] | None, where: list, sort: list,
                   limit: int | None, cursor: str | None, text_chars: int | None) -> Result:
    """``xlsx_query`` of a dataset (round 10, O5): its rows as a table's, filtered and
    sorted as a table's are (:func:`~xlsx_agent.query.dataset_matches`), each value with
    its ``d:`` address, paged as ``xlsx_read`` pages a dataset; ``dataset`` says where it
    came from (stale, closed), as there."""
    from .derive import _NONE, _TEXT, _about, _entry

    entry, address, dataset = _entry(call, target, "table")
    if address.kind != "dataset":
        raise ToolError("invalid_arguments", f"xlsx_query takes a whole dataset "
                        f"({dataset.address.text}); name its columns in columns",
                        field="table", valid_options=[dataset.address.text])
    label = dataset.address.text
    names = select(dataset, columns)
    offsets = tuple(dataset.column_index(name) for name in names)
    # A cursor belongs to one query of one definition of the dataset (an id defined again
    # is another dataset).
    key = json.dumps([label, entry.defined_at, entry.version, dataset.rows, names, where, sort],
                     ensure_ascii=False, sort_keys=True, default=str)
    found = dataset_matches(dataset, where=where, sort=sort)
    total = len(found.rows)
    start_row, start_column = _read_cursor(cursor, key, total, len(names), label)
    rows_limit = min(limit or call.limits.max_list_items, READ_ROWS)
    chars = call.limits.max_result_chars

    data: dict[str, Any] = {"target": label, "dataset": _about(entry)}
    names_cut = 0
    if _size(list(names)) <= chars // 4:
        data["columns"] = list(names)
    else:
        names_cut = len(names)
    if found.types:
        data["types"] = found.types
    if found.not_compared:
        data["not_compared"] = {n: c.as_dict() for n, c in found.not_compared.items()}
    if found.loosened:
        data["loosened"] = {n: c.as_dict() for n, c in found.loosened.items()}
    data["rows"] = []
    if not total:
        loosely = _loosely(where, lambda w: dataset_matches(dataset, where=w))
        return Result(summary=_summary(label, 1, 0, 0, dataset.rows, False), data=data,
                      total=0, warnings=list(found.notes) + loosely)

    page = found.rows[start_row:start_row + rows_limit]
    budget = chars - _MARGIN - _size({**data, "cut": cut_room(again(call))}) \
        - len(_summary(label, 1, total, total, dataset.rows, True))
    values = ([dataset.value(position, o) for o in offsets] for position in page)
    shown, cut_text, stop, full = _page(values, start_row, start_column, budget, _NONE, False,
                                        text_chars or _TEXT)
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
    warnings = list(found.notes)
    if cut_text:
        cut["text"] = text_cut(cut_text, shown, again(call), chars, warnings)
    if names_cut:
        cut["column_names"] = names_cut
    if cut:
        data["cut"] = cut
    summary = _summary(label, start_row + 1, start_row + len(shown), total, dataset.rows,
                       next_cursor is not None)
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor,
                  warnings=warnings)


def _loosely(where: list, match: Any) -> list[str]:
    """For a query that matches no row: how many rows its text filters would match with
    ``loose`` -- the filters that compare text (eq, in, contains, contains_any) and are not
    loose already, made so -- as a warning naming those filters, so the next call can be
    that one (rule 2); ``[]`` when there are none, or loose finds no more."""
    loosen = [i for i, item in enumerate(where)
              if item.get("op") in LOOSENED and not item.get("loose")]
    if not loosen:
        return []
    found: Matches = match([{**item, "loose": True} if i in loosen else item
                            for i, item in enumerate(where)])
    changed = [column for column, count in found.loosened.items() if count.count]
    if not found.rows or not changed:
        return []
    filters = ", ".join(f"where[{i}]" for i in loosen if where[i].get("column") is not None
                        and _same_column(where[i]["column"], changed))
    return [f"no row matches exactly; {len(found.rows)} would with loose: true on {filters} "
            f"(text compared ignoring case and white space, on {', '.join(changed)})"]


def _same_column(name: str, columns: list[str]) -> bool:
    folded = name.casefold()
    return any(column.casefold() == folded for column in columns)


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
