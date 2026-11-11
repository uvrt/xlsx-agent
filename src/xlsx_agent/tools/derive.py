"""``xlsx_derive`` (figures computed here, kept as a dataset ``d:<id>``) and the dataset side
of ``xlsx_read`` (ROADMAP.md, "Tools"; "Design rules", 1 and 2).  Both in the loaded core.

**xlsx_derive** is :mod:`xlsx_agent.derive`'s engine on a table or region (``t:``) of the
workbook ``doc``, or on a dataset (``d:``; no ``doc`` needed): ``where`` and ``sort`` as
``xlsx_query`` writes them (``sort`` by the result's columns: a ``group_by`` column, a
bin's or a measure's ``as``; never an input column a measure reads), ``group_by``, ``measures``, ``bins`` (edges as filter values
are written), ``pivot`` (with ``pivot_values``, its columns listed in order: round 27, X6;
``""`` the blank cells, ``(blank)``: round 28, X8),
``top``, ``round`` and ``coerce``.  The result is kept in the
session as ``d:<id>`` (:meth:`Call.define_dataset <ooxml_edit.tools.Call.define_dataset>`:
the :class:`~ooxml_edit.tools.Dataset` of :func:`xlsx_agent.dataset.as_dataset`, the
:class:`~xlsx_agent.derive.Derived` as its detail, each value's sources as range text).  A
dataset derived from a dataset inherits its source document and version (it goes stale
with it), and its values' sources are the original cells (:mod:`xlsx_agent.dataset`).

**join** (round 25, X1; :mod:`xlsx_agent.join`): ``{from, on, columns}`` adds another
table's (``t:``, read in ``doc``) or dataset's (``d:``) columns to each row of ``from``, by
key, before everything else; without ``group_by``, ``measures``, ``bins`` or ``pivot`` the
joined rows are the dataset.  The result's ``joined`` says what was matched.  A dataset
joined with a table needs ``doc``.  The dataset goes stale when either side's source
changes: ``from``'s, and the joined table's workbook or the joined dataset's source
(ooxml-edit's ``define_dataset(also=...)``).

**fill** (round 35, R22; :mod:`xlsx_agent.fill`): ``[{column, from, coerce}]`` gives each
blank cell of ``column`` the same row's ``from`` value, after the join (``from`` may be a
joined column) and before everything else; ``coerce`` (``number``, ``date``) reads text
as the derive's ``coerce`` does.  Each filled value's sources are the ``from`` cells and
it carries ``filled`` (the column it came from) in the result's rows and in a read of the
dataset; the result's ``filled`` gives per fill the cells filled, ``coerced`` and those
left blank (``unfilled``, by reason), each with first addresses.  Without ``group_by``,
``measures``, ``bins`` or ``pivot`` the filled rows are the dataset.

**where alone** (round 51): without ``group_by``, ``measures``, ``bins``, ``pivot``,
``join`` or ``fill``, the rows ``where`` keeps are the dataset, every column (``sort`` and
``top`` apply); ``measures: [{fn: count}]`` counts them instead, as the description says.

What the result gives: the dataset's id and source, its columns, types and key columns,
the rounding, the counts (``matched``, ``used``, ``groups``), what was left out and why
(``excluded`` per measure, ``left_out`` per column, by reason, with first addresses),
``coerced``, ``not_compared``, ``hidden``, the ``totals`` shares are fractions of, the
``pivot_columns`` (and ``pivot_unlisted``: the values not in ``pivot_values``, with their
row counts), ``cache`` (the formula cells read, by cache status and ``external``:
a result computed from ``invalidated`` caches is a stale dataset from the start, and the
toolbox warns); and its first rows, each value typed with its address
(``d:by_region/#2/total``) and its sources -- by key when the table has a key column
(``t:Sales/{S-01,S-04}/Amount``), else ranges (:mod:`xlsx_agent.derive`) -- **at most**
:data:`SOURCES_SHOWN` **a value**, counted in units (a range, or one key of a key set),
the rest counted in ``more`` -- within ``Limits.max_list_items`` rows and
``Limits.max_result_chars``.  ``cut`` says what was left for ``xlsx_read``, and ``next``
gives the target and cursor that continue it there.

**xlsx_read of a dataset** (:func:`read_dataset`): ``d:by_region`` (its rows),
``d:by_region[total]``, ``d:by_region/#2``, ``d:by_region/North`` (one key column) and
``d:by_region/North/total``, paged as cells are; ``columns`` picks its columns, as a
table's (round 28, X7: in the order given, the address's own column first; a name it has no
column of ``not_found``, with its columns; a cursor belongs to those columns), with
``sources`` adding each value's
first :data:`SOURCES_SHOWN` units; one value with ``sources`` pages through all of its
units instead, ``limit`` a page, a page's keys of one column written as one key set.
Text longer than ``text_chars`` (by default :data:`_TEXT`; one value: as far as the page
holds) is cut with its length, and ``cut.text`` gives the read that has it in full, as a
table's read does (:func:`~.read.text_cut`); so does ``xlsx_derive``'s ``cut.text``.  A dataset outlives its workbook: no ``doc`` is needed.  ``dataset`` says
where it came from and whether it is ``stale`` (its source has changed since, or it was
computed from invalidated formula caches; it is never refreshed) or its source ``closed``,
and its ``cache`` counts; the toolbox adds a warning when it is stale.

**Content.**  Values, keys and pivot values are in results only; summaries name the
dataset and count, errors name columns, functions and types.
"""

from __future__ import annotations

import json
from dataclasses import replace
from typing import Any, Iterator, Sequence

from ooxml_edit.tools import (CORE, DATASET_PREFIX, Result, ToolError, array, boolean, integer,
                              obj, string, tool)

from ..address import DatasetAddress, NameNotFound, parse_address
from ..clean import UNMATCHED, unmatched_warning
from ..dataset import DatasetSelection, DatasetTable, cut_sources, join_units, source_units
from ..read.regions import RegionError
from ..derive import COERCIONS, FUNCTIONS, MAX_DIGITS, MAX_ROWS, derive
from ..fill import FILL_TYPES
from .docformat import KIND, handler
from .query import WHERE
from .read import (_MARGIN, READ_ROWS, _Annotations, _cursor, _page, _read_cursor, _size, again,
                   cut_room, text_cut)

#: The most source units (a range, or one key of a key set) a value shows in a result; the
#: rest are counted (ROADMAP.md, "Design rules", 2).
SOURCES_SHOWN = 8

#: The sort keys as ``xlsx_derive`` takes them: ``xlsx_query``'s, by the result's columns.
SORT = array(obj({"column": string(), "desc": boolean(optional=True)}),
             "By output column (group_by, a measure's as), not input; first decides first. "
             "Blanks last.", optional=True)

#: How long a text value in a page of a dataset may be before it is cut (``xlsx_read``'s and
#: ``xlsx_query``'s ``text_chars`` changes it).
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
    assert isinstance(address, DatasetAddress)   # call.dataset took it as one
    return entry, address, DatasetTable(entry.dataset, entry.id)



def _about(entry: Any) -> dict[str, Any]:
    """Where a dataset came from: never its values."""
    out: dict[str, Any] = {"id": entry.address, "doc": entry.doc, "stale": entry.stale}
    if entry.closed:
        out["closed"] = True
    if entry.dataset.cache:
        out["cache"] = dict(entry.dataset.cache)
    return out


# -- xlsx_derive -------------------------------------------------------------------------------


@tool("xlsx_derive",
      "Figures computed here, never by you, from a table's or dataset's rows. min/max of "
      "dates are dates; first/last by row order. Kept as dataset d:<id> (xlsx_read reads "
      "it; derive from it again); every value lists its source cells; cells left out are "
      "counted by reason. where/join/fill without group_by, measures, bins, pivot: the rows "
      "(sort, top apply; a count: measures [{fn: count}]).",
      {"doc": string("Document id; not needed for a dataset (d:).", optional=True),
       "from": string("Table or region (t:Products), or dataset (d:by_region)."),
       "id": string("Dataset id: by_region (letters, digits, _ . -)."),
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
       "pivot_values": array(string(), "The pivot's columns, in this order, written as "
                             "filter values (\"\" for the blank cells); a value without rows "
                             "still gets its column (count 0, other measures blank). Rows of "
                             "a value not listed are left out and named in pivot_unlisted.",
                             optional=True),
       "sort": SORT,
       "top": integer("Keep the first rows after sorting.", minimum=1, maximum=MAX_ROWS,
                      optional=True),
       "round": integer("Decimals, half away from zero; reported values only.",
                        minimum=-MAX_DIGITS, maximum=MAX_DIGITS, optional=True),
       "coerce": array(string(enum=COERCIONS), "Opt in: read numbers or dates stored as text.",
                       optional=True),
       "join": obj({"from": string("The table (t:) or dataset (d:) to join."),
                    "on": array(string(), "Two names: [key column of the outer from (the "
                                "rows joined to), key column of this join's table]. A list "
                                "cell (a split column) matches by any of its items."),
                    "columns": array(string(), "This join's table's columns to add. Default: all but "
                                     "its key.",
                                     optional=True)},
                   "First, add another table's columns to each row of the outer from, by key, e.g. {from: \"t:Issues\", on: [\"Ref\", \"Items\"], columns: "
                   "[\"Issue\"]}. Every row kept (blank if unmatched); several matches: each "
                   "added cell a list.",
                   optional=True),
       "fill": array(obj({"column": string(), "from": string(),
                          "coerce": string(enum=FILL_TYPES, optional=True)}),
                     "After join: each blank of column takes the same row's value of "
                     "the fill's from column, e.g. {column: \"Limit\", from: \"Default\"}. "
                     "Sources kept, marked filled.",
                     optional=True)},
      group=CORE, datasets=("from",))
@handler
def xlsx_derive(call: Any, id: str, doc: str | None = None, **arguments: Any) -> Result:
    """``xlsx_derive``'s handler, for a workbook or a dataset (see the module's notes)."""
    source = arguments.pop("from")
    options = _options(arguments)
    join = options.get("join")
    also: list[Any] = []            # the joined side's source, when not from's own
    if join is not None:
        options["join"], joined_entry = _join_side(call, join, doc)
        if joined_entry is not None:
            also.append(joined_entry)
    if isinstance(source, str) and source.startswith(DATASET_PREFIX):
        entry, address, table = _entry(call, source, "from")
        if address.kind != "dataset":
            raise ToolError("invalid_arguments", "from names a whole dataset (d:by_region), "
                            "not a column or row of one", field="from")
        workbook = None
        if isinstance(options.get("join", {}).get("from"), str):    # a table: its workbook
            if doc is None:
                raise ToolError("invalid_arguments", "join's from is a table (t:): give doc, "
                                "the workbook it is in", field="doc")
            workbook = workbook_of(call, doc)
        if workbook is not None:
            also.append(doc)        # the joined table's workbook
        derived = derive(workbook, table, id=id, **options)
        dataset = derived.dataset()
        name = call.define_dataset(id, dataset, source=entry, also=also)
    else:
        derived = workbook_of(call, doc).derive(source, id=id, **options)
        dataset = derived.dataset()
        name = call.define_dataset(id, dataset, origin=derived.address, also=also)
    return _derived_result(call, name, derived, DatasetTable(dataset, id))


def _options(arguments: dict[str, Any]) -> dict[str, Any]:
    """The engine's arguments from the tool's: nulls dropped from the objects, as a
    strict call may send them."""

    def clean(item: Any) -> Any:
        return {k: v for k, v in item.items() if v is not None} if isinstance(item, dict) \
            else item

    out: dict[str, Any] = {}
    for name in ("where", "measures", "sort", "fill"):
        if arguments.get(name) is not None:
            out[name] = [clean(item) for item in arguments[name]]
    if arguments.get("bins") is not None:
        out["bins"] = clean(arguments["bins"])
    for name in ("group_by", "pivot", "pivot_values", "top", "round", "coerce"):
        if arguments.get(name) is not None:
            out[name] = arguments[name]
    if arguments.get("join") is not None:
        out["join"] = clean(arguments["join"])
    return out


def _join_side(call: Any, join: Any, doc: str | None) -> tuple[Any, Any]:
    """``join`` with its ``from`` a dataset's table when it names one (``d:``), and that
    dataset's entry, a source the result also read (it goes stale with it); a table
    (``t:``) is left for the workbook ``doc`` to read (no entry)."""
    if not isinstance(join, dict) or not isinstance(join.get("from"), str) \
            or not join["from"].startswith(DATASET_PREFIX):
        return join, None
    entry, address, table = _entry(call, join["from"], "join")
    if address.kind != "dataset":
        raise ToolError("invalid_arguments", "join's from names a whole dataset "
                        "(d:by_region), not a column or row of one", field="join")
    return {**join, "from": table}, entry


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
    for key in ("coerced", "cleaned", "not_compared", "loosened"):
        value = getattr(derived, key)
        if value:
            data[key] = {k: c.as_dict() for k, c in value.items()}
    warnings = list(derived.notes)
    if derived.unmatched:
        data[UNMATCHED] = {k: c.as_dict() for k, c in derived.unmatched.items()}
        warnings += [unmatched_warning(c.count, c.first[0], k)
                     for k, c in derived.unmatched.items()]
    if derived.hidden.count:
        data["hidden"] = derived.hidden.as_dict()
    if derived.pivot_columns:
        data["pivot_columns"] = {k: dict(v) for k, v in derived.pivot_columns.items()}
    if derived.pivot_unlisted:
        data["pivot_unlisted"] = [dict(u) for u in derived.pivot_unlisted]
    if derived.totals:
        data["totals"] = {k: _shown(t.as_dict()) for k, t in derived.totals.items()}
    if derived.cache:
        data["cache"] = dict(derived.cache)
    if derived.joined is not None:
        data["joined"] = derived.joined.as_dict()
    if derived.filled:
        data["filled"] = [f.as_dict() for f in derived.filled]
    total = len(derived.rows)
    cut: dict[str, Any] = {}
    chars = call.limits.max_result_chars
    # what is about the dataset, not its rows, may take half the page at most
    for key in ("pivot_unlisted", "pivot_columns", "types", "columns"):
        if _size(data) <= chars // 2 or key not in data:
            continue
        cut[key] = len(data.pop(key))
    rows_limit = min(call.limits.max_list_items, READ_ROWS)
    budget = chars - _MARGIN - _size({**data, "cut": {**cut, **_CUT_ROOM},
                                      "next": {"target": name, "cursor": "x" * 24}}) - 80
    rows = _values(table, range(1, min(total, rows_limit) + 1), _every(table), SOURCES_SHOWN)
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
        cut["text"] = text_cut(cut_text, shown, {"target": name},
                               call.limits.max_result_chars, warnings)
    if cut:
        data["cut"] = cut
    summary = f"Derived {name} from {derived.address}: {total} rows"
    if stop is not None:
        summary += f"; more with xlsx_read {name}"
    return Result(summary=summary, data=data, total=total, warnings=warnings)


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


def _values(table: DatasetTable, positions: Sequence[int], offsets: tuple[int, ...],
            sources: int | None) -> Iterator[list[Any]]:
    """The rows at ``positions`` of ``table``, in that order, the columns at ``offsets``,
    as values (each with its first ``sources`` ranges when given); :func:`~.read._page`
    starts a page's first row at its own column."""
    for position in positions:
        yield [table.value(position, o, sources=sources) for o in offsets]


#: A page needs no comments or links: a dataset has none.
_NONE = _Annotations(None, [], comments=False, hyperlinks=False)

#: A ``cut`` at its longest, to keep room for it: its ``text.call`` reads the dataset
#: (:func:`~.read.text_cut`).
_CUT_ROOM = {"rows": 1000, "cells": 10000, "chars": True,
             "text": {"count": 100000, "first": ["d:" + "x" * 40 + "/#100000/" + "x" * 40] * 3,
                      "longest": 100000, "call": {"target": "d:" + "x" * 40,
                                                  "text_chars": 100000}}}


# -- xlsx_read of a dataset --------------------------------------------------------------------


def read_dataset(call: Any, target: str, *, sources: bool, limit: int | None,
                 max_chars: int | None, cursor: str | None,
                 columns: list[str] | None = None, key: str | None = None,
                 text_chars: int | None = None) -> Result:
    """``xlsx_read`` of a ``d:`` target (see the module's notes); ``key`` the column
    holding the keys of a row by key or a key set (:meth:`DatasetTable.select`);
    ``text_chars`` the length a text is cut at (by default :data:`_TEXT`; one value: as
    far as the page holds)."""
    entry, address, table = _entry(call, target, "target")
    warnings: list[str] = []
    if key is not None and address.key is None and address.keys is None:
        warnings.append(f"key {key!r} not used: it names the column holding the keys of a "
                        f"row by key or a key set ({DatasetAddress(address.id).text}/{{a,b}}), and "
                        f"{address.text} names none")
        key = None
    selection = _selected(table, address, key)
    token = selection.text if key is None else f"{selection.text}|key={key}"
    if columns:
        selection = _picked(table, address, columns, key)
        token = f"{selection.text}|{','.join(map(str, selection.offsets))}"
    chars = min(max_chars or call.limits.max_result_chars, call.limits.max_result_chars)
    rows_limit = min(limit or call.limits.max_list_items, READ_ROWS)
    data: dict[str, Any] = {"target": selection.text, "dataset": _about(entry)}
    if sources and selection.kind == "value":
        return _value_sources(table, selection, data, chars, rows_limit, cursor)
    total = selection.rows
    width = len(selection.offsets)
    start_row, start_column = _read_cursor(cursor, token, total, width)
    data["columns"] = [table.columns[o] for o in selection.offsets]
    if _size(data["columns"]) > chars // 4:
        data.pop("columns")
    data["rows"] = []
    budget = chars - _MARGIN - _size({**data, "cut": cut_room(again(call))}) - 80
    order = selection.order
    page = order[start_row:start_row + rows_limit]
    values = _values(table, page, selection.offsets, SOURCES_SHOWN if sources else None)
    cap = text_chars or (None if selection.kind == "value" else _TEXT)
    shown, cut_text, stop, full = _page(values, start_row, start_column, budget, _NONE,
                                        False, cap)
    if stop is None and start_row + len(page) < len(order):
        stop = (start_row + len(page), 0)
    data["rows"] = shown
    cut: dict[str, Any] = {}
    next_cursor = None
    if stop is not None:
        cut["rows"] = total - stop[0]
        if stop[1]:
            cut["cells"] = width - stop[1]
        if full:
            cut["chars"] = True
        next_cursor = _cursor(token, *stop)
    if cut_text:
        cut["text"] = text_cut(cut_text, shown, again(call), call.limits.max_result_chars,
                               warnings)
    if cut:
        data["cut"] = cut
    return Result(summary=_summary(selection.text, start_row + 1, start_row + len(shown),
                                   total, next_cursor is not None, entry.stale),
                  data=data, total=total, next_cursor=next_cursor, warnings=warnings)


def _selected(table: DatasetTable, address: DatasetAddress,
              key: str | None) -> DatasetSelection:
    """:meth:`DatasetTable.select`, a dataset without a key column refused with the call
    that works: the same read, with ``key`` naming a column that can be the key (rule 2);
    and ``key`` naming no column refused with the dataset's columns."""
    try:
        return table.select(address, key=key)
    except NameNotFound as error:
        if key is None or error.what != "column" or error.name != key:
            raise
        raise ToolError("not_found", f"{error.message}; key names the column holding the "
                        "keys", field="key", valid_options=list(table.columns)[:50],
                        details={"what": "column"}) from None
    except RegionError as error:
        if error.code != "no_key_column" or not error.valid_options:
            raise
        call = {"target": address.text, "key": error.valid_options[0]}
        raise ToolError("refused", f"{error.message}. The call that works: xlsx_read "
                        f"{json.dumps(call)}", field="target",
                        valid_options=list(error.valid_options),
                        details={"code": "no_key_column", "call": call}) from None


def _picked(table: DatasetTable, address: DatasetAddress,
            columns: list[str], key: str | None = None) -> DatasetSelection:
    """The rows ``address`` names, of the ``columns`` asked for, in their order, as a
    table's are picked (round 28, X7): the address's own column first, each column once;
    ``not_found`` for a name the dataset has no column of, with its columns."""
    wanted = ([address.column] if address.column is not None else []) + list(columns)
    try:
        offsets = tuple(dict.fromkeys(table.column_index(name) for name in wanted))
    except NameNotFound as error:
        raise ToolError("not_found", f"{error.message}; columns picks some of them",
                        field="columns", valid_options=list(table.columns)[:50],
                        details={"what": "column"}) from None
    rows = table.select(replace(address, column=None, any_row=False, any_column=False),
                        key=key)
    return replace(rows, offsets=offsets)


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
