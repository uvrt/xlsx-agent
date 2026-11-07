"""``xlsx_trace`` (round 31, R7): rows of a table or dataset, each with the rows it rests on
and why, every value in full, within a size budget per reply.  On demand: its own deferred
group, ``xlsx_trace`` (a checking tool, not every run's).

What it traces is :mod:`xlsx_agent.trace`'s: ``rows`` are addresses of table rows
(``t:Items/I-4``, a key set ``t:Items/{I-4,I-7}``, a row range ``t:Items/#3:#9``, a whole
table, or a sheet's cells inside one table, as ``xlsx_read`` follows a citation) or of a
dataset's (``d:by_owner/#2``, ``d:by_owner/#2:#5``, a row by key, the whole dataset), each
row traced in order.  A dataset row's ``source`` links are the rows of ``doc`` its values
came from, so ``doc`` must be the workbook the dataset was derived from, and open.

**The result.**  ``traced``: per row, ``row`` its address and ``links`` (each ``row``,
``why`` -- ``source``, ``names``, ``lists`` -- and, where they apply, ``by``, ``column``,
``columns`` and ``depth``; ``more`` counts links past :data:`~xlsx_agent.trace.MAX_LINKS`);
``rows``: every row's values by address, each row once however many links name it, every
value in full (a declared split's items as a list; blanks left out).

**The budget.**  A reply holds at most ``max_chars`` characters (by default, and at most,
``Limits.max_result_chars``).  Rows are traced in order while their links fit; then the
values of the rows they name, in link order, while they fit.  What did not fit is listed,
never cut: ``left_out`` the rows named whose values are not in ``rows``, as addresses
``xlsx_read`` reads (one key set per table where the rows have keys), and ``untraced`` the
rows asked for that were not traced, to pass back as ``rows`` (key sets, and row ranges
where rows by position run on); ``next`` says so.  At most :data:`MAX_TRACED` rows are
traced per call; at most 200 addresses are listed in each, the rest counted
(``left_out_more``, ``untraced_more``).

**Content.**  The result carries cell content, as a read does; the summary and errors
carry addresses and counts only.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any

from ooxml_edit.tools import DATASET_PREFIX, Result, ToolError, array, integer, string, tool

from ..address import DatasetAddress, TableAddress, parse_address
from ..dataset import DatasetTable
from ..trace import DEPTH, MAX_DEPTH, Trace, Tracer
from .derive import workbook_of
from .docformat import handler
from .read import _MARGIN, DOC, MIN_CHARS, _size

#: The group ``xlsx_trace`` is in: deferred, loaded when a model looks for it.
GROUP = "xlsx_trace"

#: The most rows one call traces; the rest are ``untraced``.
MAX_TRACED = 50

#: The most addresses ``left_out`` and ``untraced`` list each; the rest are counted.
_LISTED = 200

_NEXT = "left_out: read them with xlsx_read; untraced: pass them back as rows."


@tool("xlsx_trace",
      "Rows of a table or dataset, each with the rows it rests on and why: source (its "
      "values' rows), names (rows its text names by key, and those they name), lists (rows "
      "whose text names its key). Every value in full, each row once; what did not fit is "
      "listed.",
      {"doc": DOC,
       "rows": array(string(), "t:Items/I-4, t:Items/{I-4,I-7}, t:Items/#3:#9, "
                     "d:by_owner/#2, a whole table or dataset.", min_items=1),
       "depth": integer(f"How far names are followed. Default {DEPTH}.", minimum=0,
                        maximum=MAX_DEPTH, optional=True),
       "max_chars": integer("Characters per reply. Default and at most 24000.",
                            minimum=MIN_CHARS, optional=True)},
      group=GROUP, datasets=("rows[]",))
@handler
def xlsx_trace(call: Any, doc: str, rows: list[str], depth: int | None = None,
               max_chars: int | None = None) -> Result:
    """``xlsx_trace``'s handler (see the module's notes)."""
    workbook = workbook_of(call, doc)
    tracer = Tracer(workbook, DEPTH if depth is None else depth)
    wanted = _wanted(call, tracer, doc, rows)
    chars = min(max_chars or call.limits.max_result_chars, call.limits.max_result_chars)
    budget = max(chars, MIN_CHARS) - _MARGIN - _size({"next": _NEXT}) - 200
    traced: list[dict[str, Any]] = []
    shown: dict[str, dict[str, Any]] = {}
    named: list[str] = []
    used = 0
    stop = None
    for index, (address, make) in enumerate(wanted):
        if index >= MAX_TRACED:
            stop = index
            break
        found: Trace = make()
        entry: dict[str, Any] = {"row": found.row,
                                 "links": [link.as_dict() for link in found.links]}
        if found.more:
            entry["more"] = found.more
        own = tracer.values.get(found.row, {})
        cost = _size(entry) + 1 + (0 if found.row in shown else _size({found.row: own}))
        if traced and used + cost > budget:
            stop = index
            break
        traced.append(entry)
        shown.setdefault(found.row, own)
        used += cost
        named.extend(link.row for link in found.links if link.row not in named)
    left_out: list[str] = []
    for address in named:
        if address in shown:
            continue
        values = tracer.values.get(address, {})
        cost = _size({address: values}) + 1
        if used + cost > budget:
            left_out.append(address)
            continue
        shown[address] = values
        used += cost
    data: dict[str, Any] = {"traced": traced, "rows": shown}
    untraced = [address for address, _ in wanted[stop:]] if stop is not None else []
    for name, listed in (("left_out", _key_sets(left_out)),
                         ("untraced", _key_sets(_runs(untraced)))):
        if listed:
            data[name] = listed[:_LISTED]
            if len(listed) > _LISTED:
                data[f"{name}_more"] = len(listed) - _LISTED
    if left_out or untraced:
        data["next"] = _NEXT
    links = sum(len(entry["links"]) + entry.get("more", 0) for entry in traced)
    summary = f"Traced {len(traced)} of {len(wanted)} row(s): {links} link(s)"
    if left_out or untraced:
        summary += (f"; {len(left_out)} row(s) named left out, {len(untraced)} not traced: "
                    "see next")
    warnings = []
    if tracer.index.keys_cut:
        warnings.append(f"{tracer.index.keys_cut} keys past the first "
                        f"{len(tracer.index.holding) - tracer.index.keys_cut} were not "
                        "looked for in text")
    return Result(summary=summary, data=data, warnings=warnings)


def _wanted(call: Any, tracer: Tracer, doc: str, rows: list[str]) -> list[tuple[str, Any]]:
    """Each row asked for, in order: its address and how to trace it."""
    out: list[tuple[str, Any]] = []
    for n, given in enumerate(rows):
        if isinstance(given, str) and given.startswith(DATASET_PREFIX):
            entry = call.dataset(given, field=f"rows[{n}]")
            if entry.doc != doc:
                raise ToolError("invalid_arguments", f"{entry.address} came from "
                                f"{entry.doc or 'no workbook'}, whose rows its sources name: "
                                f"give doc {entry.doc}" if entry.doc else
                                f"{entry.address} came from no workbook: trace a table's rows",
                                field="doc", valid_options=[entry.doc] if entry.doc else [])
            table = DatasetTable(entry.dataset, entry.id)
            selection = table.select(given)
            for position in selection.order:
                address = DatasetAddress(table.id, position=position).text
                out.append((address, lambda t=table, p=position: tracer.trace_dataset_row(t, p)))
            continue
        for address in tracer.table_rows(given):
            out.append((address, lambda a=address: tracer.trace_row(a)))
    return out


def _key_sets(addresses: list[str]) -> list[str]:
    """``addresses`` as ``xlsx_read`` reads them: rows by key of one table as one key set
    (``t:Items/{I-1,I-3}``), in the order first met; any other address as it is."""
    keyed: dict[str, tuple[TableAddress, list[str]]] = {}
    order: list[str] = []
    for text in addresses:
        try:
            address = parse_address(text)
        except Exception:  # noqa: BLE001 -- cells as the trace named them
            address = None
        if isinstance(address, TableAddress) and address.key is not None \
                and address.column is None and not set(address.key) & set(",{}"):
            whole = replace(address, key=None)
            if whole.text not in keyed:
                keyed[whole.text] = (whole, [])
                order.append(whole.text)
            keyed[whole.text][1].append(address.key)
            continue
        order.append(text)
    out = []
    for item in order:
        if item not in keyed:
            out.append(item)
            continue
        whole, keys = keyed[item]
        out.append((replace(whole, key=keys[0]) if len(keys) == 1
                    else replace(whole, keys=tuple(keys))).text)
    return out


def _runs(addresses: list[str]) -> list[str]:
    """Rows by position that run on as one row range (``d:by_owner/#51:#80``,
    ``t:Items/#3:#9``); any other address as it is."""
    out: list[str] = []
    run: tuple[Any, int, int] | None = None   # (whole, first, last)

    def flush() -> None:
        if run is None:
            return
        whole, first, last = run
        out.append(replace(whole, position=first).text if first == last
                   else replace(whole, row_range=(first, last)).text)

    for text in addresses:
        address = parse_address(text)
        position = getattr(address, "position", None)
        if isinstance(address, (TableAddress, DatasetAddress)) and position is not None:
            whole = replace(address, position=None)
            if run is not None and run[0] == whole and run[2] + 1 == position:
                run = (whole, run[1], position)
                continue
            flush()
            run = (whole, position, position)
            continue
        flush()
        run = None
        out.append(text)
    flush()
    return out


TOOLS = [xlsx_trace]
