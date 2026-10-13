"""``xlsx_derive`` (figures computed here, kept as a dataset ``d:<id>``) and the dataset side
of ``xlsx_read`` (ROADMAP.md, "Tools"; "Design rules", 1 and 2).  Both in the loaded core.

**xlsx_derive** is :mod:`xlsx_agent.derive`'s engine on a table or region (``t:``) of the
workbook ``doc``, or on a dataset (``d:``; no ``doc`` needed): ``where`` and ``sort`` as
``xlsx_query`` writes them, ``group_by``, ``measures``, ``bins`` (edges as filter values
are written), ``pivot``, ``top``, ``round`` and ``coerce``.  The result is kept in the
session as ``d:<id>`` (:meth:`Call.define_dataset <ooxml_edit.tools.Call.define_dataset>`:
the :class:`~ooxml_edit.tools.Dataset` of :func:`xlsx_agent.dataset.as_dataset`, the
:class:`~xlsx_agent.derive.Derived` as its detail, each value's sources as range text).  A
dataset derived from a dataset inherits its source document and version (it goes stale
with it), and its values' sources are the original cells (:mod:`xlsx_agent.dataset`).

What the result gives: the dataset's id and source, its columns, types and key columns,
the rounding, the counts (``matched``, ``used``, ``groups``), what was left out and why
(``excluded`` per measure, ``left_out`` per column, by reason, with first addresses),
``coerced``, ``not_compared``, ``hidden``, the ``totals`` shares are fractions of, the
``pivot_columns``; and its first rows, each value typed with its address
(``d:by_region/#2/total``) and its sources -- by key when the table has a key column
(``t:Sales/{S-01,S-04}/Amount``), else ranges (:mod:`xlsx_agent.derive`) -- **at most**
:data:`SOURCES_SHOWN` **a value**, counted in units (a range, or one key of a key set),
the rest counted in ``more`` -- within ``Limits.max_list_items`` rows and
``Limits.max_result_chars``.  ``cut`` says what was left for ``xlsx_read``, and ``next``
gives the target and cursor that continue it there.

**xlsx_read of a dataset** (:func:`read_dataset`): ``d:by_region`` (its rows),
``d:by_region[total]``, ``d:by_region/#2``, ``d:by_region/North`` (one key column) and
``d:by_region/North/total``, paged as cells are, with ``sources`` adding each value's
first :data:`SOURCES_SHOWN` units; one value with ``sources`` pages through all of its
units instead, ``limit`` a page, a page's keys of one column written as one key set.  A dataset outlives its workbook: no ``doc`` is needed.  ``dataset`` says
where it came from and whether it is ``stale`` (its source has changed since; it is never
refreshed) or its source ``closed``; the toolbox adds a warning when it is stale.

**Content.**  Values, keys and pivot values are in results only; summaries name the
dataset and count, errors name columns, functions and types.
"""

from __future__ import annotations

from typing import Any, Iterator

from ooxml_edit.tools import (CORE, DATASET_PREFIX, Result, ToolError, array, integer, obj,
                              string, tool)

from ..address import DatasetAddress, parse_address
from ..dataset import DatasetTable, cut_sources, join_units, source_units
from ..derive import COERCIONS, FUNCTIONS, MAX_DIGITS, MAX_ROWS, derive
from .docformat import KIND, handler
from .query import SORT, WHERE
from .read import _MARGIN, READ_ROWS, _Annotations, _cursor, _page, _read_cursor, _size

#: The most source units (a range, or one key of a key set) a value shows in a result; the
#: rest are counted (ROADMAP.md, "Design rules", 2).
SOURCES_SHOWN = 8

#: How long a text value in a page of a dataset may be before it is cut.
_TEXT = 2000


def workbook_of(call: Any, doc: str | None, field: str = "doc") -> Any:
    """The workbook ``doc`` names: ``invalid_arguments`` without one, or for another
    kind of document."""
    entry = call.entry
    if doc is None or entry is None:
        raise ToolError("invalid_arguments", "doc names the workbook (only a dataset, d:, "
                        "needs none)", field=field)
    if entry.kind != KIND:
        raise ToolError("invalid_arguments", f"{call.tool.name} reads a workbook; {doc} is a "
                        f"{entry.kind} document", field=field)
    return entry.document


def _entry(call: Any, target: str, field: str) -> tuple[Any, DatasetAddress, DatasetTable]:
    entry = call.dataset(target, field=field)
    address = parse_address(target)
    return entry, address, DatasetTable(entry.dataset, entry.id)


def _about(entry: Any) -> dict[str, Any]:
    """Where a dataset came from: never its values."""
    out: dict[str, Any] = {"id": entry.address, "doc": entry.doc, "stale": entry.stale}
    if entry.closed:
        out["closed"] = True
    return out


# -- xlsx_derive -------------------------------------------------------------------------------


@tool("xlsx_derive",
      "Figures computed here, never by you, from a table's or dataset's rows, filtered, "
      "grouped, binned or pivoted. min/max of dates are dates; first/last by row order. Kept "
      "as dataset d:<id> (xlsx_read reads it, xlsx_derive derives from it); every value "
      "lists its source cells; cells left out are counted by reason.",
      {"doc": string("Document id; not needed when from is a dataset.", optional=True),
       "from": string("A table or region (t:Products) or a dataset (d:by_region)."),
       "id": string("The dataset's id: by_region (letters, digits, _ . -)."),
       "where": WHERE,
       "group_by": array(string(), "Columns to group by.", optional=True),
       "measures": array(obj({"fn": string(enum=FUNCTIONS),
                              "column": string("None: count or share of rows.", optional=True),
                              "as": string("Its column's name.", optional=True)}),
                         "Default: count of rows.", optional=True),
       "bins": obj({"column": string(),
                    "edges": array(string(), "Increasing, written as filter values."),
                    "as": string(optional=True)},
                   "Group by ranges of a number or date column: [a, b), the last [a, b].",
                   optional=True),
       "pivot": string("A column whose values become columns.", optional=True),
       "sort": SORT,
       "top": integer("Keep the first rows after sorting.", minimum=1, maximum=MAX_ROWS,
                      optional=True),
       "round": integer("Decimals, half away from zero; reported values only.",
                        minimum=-MAX_DIGITS, maximum=MAX_DIGITS, optional=True),
       "coerce": array(string(enum=COERCIONS), "Opt in: read numbers or dates stored as text.",
                       optional=True)},
      group=CORE, datasets=("from",))
@handler
def xlsx_derive(call: Any, id: str, doc: str | None = None, **arguments: Any) -> Result:
    """``xlsx_derive``'s handler, for a workbook or a dataset (see the module's notes)."""
    source = arguments.pop("from")
    options = _options(arguments)
    if isinstance(source, str) and source.startswith(DATASET_PREFIX):
        entry, address, table = _entry(call, source, "from")
        if address.kind != "dataset":
            raise ToolError("invalid_arguments", "from names a whole dataset (d:by_region), "
                            "not a column or row of one", field="from")
        derived = derive(None, table, id=id, **options)
        dataset = derived.dataset()
        name = call.define_dataset(id, dataset, source=entry)
    else:
        derived = workbook_of(call, doc).derive(source, id=id, **options)
        dataset = derived.dataset()
        name = call.define_dataset(id, dataset)
    return _derived_result(call, name, derived, DatasetTable(dataset, id))


def _options(arguments: dict[str, Any]) -> dict[str, Any]:
    """The engine's arguments from the tool's: nulls dropped from the objects, as a
    strict call may send them."""

    def clean(item: Any) -> Any:
        return {k: v for k, v in item.items() if v is not None} if isinstance(item, dict) \
            else item

    out: dict[str, Any] = {}
    for name in ("where", "measures", "sort"):
        if arguments.get(name) is not None:
            out[name] = [clean(item) for item in arguments[name]]
    if arguments.get("bins") is not None:
        out["bins"] = clean(arguments["bins"])
    for name in ("group_by", "pivot", "top", "round", "coerce"):
        if arguments.get(name) is not None:
            out[name] = arguments[name]
    return out


def _derived_result(call: Any, name: str, derived: Any, table: DatasetTable) -> Result:
    data: dict[str, Any] = {"dataset": name, "from": derived.address,
                            "columns": list(derived.columns),
                            "key_columns": derived.key_columns, "types": dict(derived.types),
                            "rounding": dict(derived.rounding), "matched": derived.matched,
                            "used": derived.used, "groups": derived.groups}
    for key in ("left_out", "excluded"):
        value = getattr(derived, key)
        if value:
            data[key] = {k: {r: c.as_dict() for r, c in v.items()} for k, v in value.items()}
    for key in ("coerced", "not_compared"):
        value = getattr(derived, key)
        if value:
            data[key] = {k: c.as_dict() for k, c in value.items()}
    if derived.hidden.count:
        data["hidden"] = derived.hidden.as_dict()
    if derived.pivot_columns:
        data["pivot_columns"] = {k: dict(v) for k, v in derived.pivot_columns.items()}
    if derived.totals:
        data["totals"] = {k: _shown(t.as_dict()) for k, t in derived.totals.items()}
    total = len(derived.rows)
    cut: dict[str, Any] = {}
    chars = call.limits.max_result_chars
    # what is about the dataset, not its rows, may take half the page at most
    for key in ("pivot_columns", "types", "columns"):
        if _size(data) <= chars // 2 or key not in data:
            continue
        cut[key] = len(data.pop(key))
    rows_limit = min(call.limits.max_list_items, READ_ROWS)
    budget = chars - _MARGIN - _size({**data, "cut": {**cut, **_CUT_ROOM},
                                      "next": {"target": name, "cursor": "x" * 24}}) - 80
    rows = _values(table, 1, min(total, rows_limit), _every(table), SOURCES_SHOWN)
    shown, cut_text, stop, full = _page(rows, 0, 0, budget, _NONE, False, _TEXT)
    if stop is None and total > len(shown):
        stop = (len(shown), 0)
    data["rows"] = shown
    if stop is not None:
        cut["rows"] = total - stop[0]
        if stop[1]:
            cut["cells"] = len(table.columns) - stop[1]
        if full:
            cut["chars"] = True
        data["next"] = {"target": name, "sources": True,
                        "cursor": _cursor(name, stop[0], stop[1])}
    if cut_text:
        cut["text"] = len(cut_text)
    if cut:
        data["cut"] = cut
    summary = f"Derived {name} from {derived.address}: {total} rows"
    if stop is not None:
        summary += f"; more with xlsx_read {name}"
    return Result(summary=summary, data=data, total=total)


def _shown(sourced: dict[str, Any]) -> dict[str, Any]:
    """A sourced value's dict with :data:`SOURCES_SHOWN` units at most, the rest
    counted."""
    shown, more = cut_sources(sourced.get("sources", []), SOURCES_SHOWN)
    if more:
        sourced["sources"] = list(shown)
        sourced["more"] = more
    return sourced


def _every(table: DatasetTable) -> tuple[int, ...]:
    return tuple(range(len(table.columns)))


def _values(table: DatasetTable, first: int, last: int, offsets: tuple[int, ...],
            sources: int | None) -> Iterator[list[Any]]:
    """Rows ``first`` to ``last`` of ``table``, the columns at ``offsets``, as values
    (each with its first ``sources`` ranges when given); :func:`~.read._page` starts a
    page's first row at its own column."""
    for position in range(first, last + 1):
        yield [table.value(position, o, sources=sources) for o in offsets]


#: A page needs no comments or links: a dataset has none.
_NONE = _Annotations(None, [], comments=False, hyperlinks=False)

#: A ``cut`` at its longest, to keep room for it.
_CUT_ROOM = {"rows": 1000, "cells": 10000, "chars": True, "text": 100000}


# -- xlsx_read of a dataset --------------------------------------------------------------------


def read_dataset(call: Any, target: str, *, sources: bool, limit: int | None,
                 max_chars: int | None, cursor: str | None) -> Result:
    """``xlsx_read`` of a ``d:`` target (see the module's notes)."""
    entry, address, table = _entry(call, target, "target")
    selection = table.select(address)
    chars = min(max_chars or call.limits.max_result_chars, call.limits.max_result_chars)
    rows_limit = min(limit or call.limits.max_list_items, READ_ROWS)
    data: dict[str, Any] = {"target": selection.text, "dataset": _about(entry)}
    if sources and selection.kind == "value":
        return _value_sources(table, selection, data, chars, rows_limit, cursor)
    total = selection.rows
    width = len(selection.offsets)
    start_row, start_column = _read_cursor(cursor, selection.text, total, width)
    data["columns"] = [table.columns[o] for o in selection.offsets]
    if _size(data["columns"]) > chars // 4:
        data.pop("columns")
    data["rows"] = []
    budget = chars - _MARGIN - _size({**data, "cut": _CUT_ROOM}) - 80
    first = selection.first + start_row
    last = min(selection.last, first + rows_limit - 1)
    values = _values(table, first, last, selection.offsets,
                     SOURCES_SHOWN if sources else None)
    cap = None if selection.kind == "value" else _TEXT
    shown, cut_text, stop, full = _page(values, start_row, start_column, budget, _NONE,
                                        False, cap)
    if stop is None and last < selection.last:
        stop = (last - selection.first + 1, 0)
    data["rows"] = shown
    cut: dict[str, Any] = {}
    next_cursor = None
    if stop is not None:
        cut["rows"] = total - stop[0]
        if stop[1]:
            cut["cells"] = width - stop[1]
        if full:
            cut["chars"] = True
        next_cursor = _cursor(selection.text, *stop)
    if cut_text:
        cut["text"] = len(cut_text)
    if cut:
        data["cut"] = cut
    return Result(summary=_summary(selection.text, start_row + 1, start_row + len(shown),
                                   total, next_cursor is not None, entry.stale),
                  data=data, total=total, next_cursor=next_cursor)


def _value_sources(table: DatasetTable, selection: Any, data: dict[str, Any], chars: int,
                   limit: int, cursor: str | None) -> Result:
    """One value and its sources, every one of them, paged ``limit`` units at a time (a
    range, or one key of a key set; a page's keys of one column joined in one key set)."""
    position, offset = selection.first, selection.offsets[0]
    every = source_units(table.cell_sources(position, offset))
    total = len(every)
    token = selection.text + " sources"
    start, _ = _read_cursor(cursor, token, max(total, 1), 1)
    data["value"] = table.value(position, offset).as_dict()
    budget = chars - _MARGIN - _size({**data, "cut": _CUT_ROOM}) - 80
    taken = 0
    used = 0
    for column, text in every[start:start + limit]:
        # a key adds itself and a comma; a range or a new key set its whole text
        size = _size(text) + 1 if column is None or taken == 0 \
            or every[start + taken - 1][0] != column else len(text) + 3
        if column is not None and (taken == 0 or every[start + taken - 1][0] != column):
            size += _size(column.text) + 4
        if taken and used + size > budget:
            break
        taken += 1
        used += size
    shown = join_units(every[start:start + taken])
    data["sources"] = shown
    end = start + taken
    next_cursor = None
    if end < total:
        data["cut"] = {"sources": total - end}
        next_cursor = _cursor(token, end, 0)
    summary = f"{selection.text}: sources {start + 1}-{end} of {total}" if total else \
        f"{selection.text}: no sources"
    if next_cursor is not None:
        summary += "; more with next_cursor"
    return Result(summary=summary, data=data, total=total, next_cursor=next_cursor)


def _summary(target: str, first: int, last: int, total: int, more: bool, stale: bool) -> str:
    text = f"{target}: no rows" if total == 0 or last < first else \
        f"{target}: rows {first}-{last} of {total}"
    if more:
        text += "; more with next_cursor"
    if stale:
        text += " (stale: its source has changed)"
    return text


TOOLS = [xlsx_derive]
