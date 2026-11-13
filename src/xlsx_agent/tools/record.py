"""``xlsx_record_values`` (round 53, R108): values the caller judged from a column's text, one
per row by key, kept as a dataset ``d:<id>`` traced to the cells judged.  On demand: its own
deferred group, ``xlsx_record``; the core is unchanged.

The engine is :mod:`xlsx_agent.record`: the dataset's rows are the key, ``column`` (the
value as given) and ``note`` (the caller's reason, blank when none), one row per key
recorded, in the table's row order; each value's source is the source cell of its row, so
``xlsx_read`` with ``sources``, ``xlsx_trace`` (to the row judged), ``xlsx_derive`` (counts
by value, a join), a deck's or document's ``{{d:<id>/<key>/<column>}}`` and a table sync
take it as they take a derive's dataset.  It goes stale when the workbook changes, as a
derive's does.  Re-recording an id replaces the dataset (``replaced``, and a warning says
so); with ``add: true`` (round 54) the call's values are added to it instead (``added``):
the same table, key, source, column and type, else refused; a key already recorded with the
same value is no error, with another value refused naming the key; ``add`` to an id there
is not yet records it, with a warning.  ``type: "number"`` (round 54): each value a number
-- a JSON number, or text that is a plain number (``"2.5"``) -- kept in a number column
that ``xlsx_derive`` averages; ``value`` is a string or a number to the dispatcher
(ooxml-edit's ``orNumber``), a string to the provider.

**Repeated keys** (round 54, R113): keys are matched among the rows with text in
``source``; a heading or spacer row repeating a key is left out (a warning counts them).  A
key named that two or more rows with text hold names no one row: that value alone is not
recorded and the rest of the call is (round 55, R118, rule 5) -- ``refused`` lists it as a
refusal's ``details.problems`` would (``{"code": "key_shared", "item", "key", "rows"}``, by
place in ``values``), a warning names its place and the way out; such rows are listed in
``shared_keys``.

**Refused, nothing kept** (rule 2): a key no row has (keys are matched exactly; the rows
whose key is loosely equal by ooxml-edit's ``loose_key`` are given in ``like`` as
``{"row", "key"}`` -- the key as its cell holds it, to send so -- and their rows as
``valid_options``), a key twice, a value from an empty source cell, an empty value, a value
not a number (``type: "number"``), a key recorded earlier with another value (``add``);
every problem of the call at once in ``details.problems``, by place in ``values`` -- a
shared key's too, and a call whose every value names a shared key.

**The result** (rule 4): ``recorded``, ``counts`` per value (at most :data:`LISTED`
values, the rest counted in ``counts_more``), ``not_recorded`` -- the rows with text in the
source column and no value: ``count``, the first :data:`LISTED` ``keys`` and ``read``, the
``xlsx_read`` call that gives their key and text (a key set; the rest of them counted, ask
again after recording), or ``query``, an ``xlsx_query`` call, where another row holds one of
those keys too -- ``no_key``, rows with source text but no key, by position,
``shared_keys``, and ``refused`` (at most :data:`LISTED`, the rest counted in
``refused_more``).  The values themselves are not echoed: ``xlsx_read`` of ``d:<id>`` reads
them.

**Content.**  The result carries keys and the values given (a refusal's problems the key
cells' text); the summary and the errors' messages carry counts, places and rows by
position only.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any

from ooxml_edit.tools import (DATASET_PREFIX, Result, ToolError, array, boolean, obj, string,
                              tool)

from ..address import DatasetAddress, TableAddress, parse_address
from ..record import RecordError, record_values, shared_way
from .derive import workbook_of
from .docformat import handler
from .read import DOC

#: The group ``xlsx_record_values`` is in: deferred, loaded when a model looks for it.
GROUP = "xlsx_record"

#: The most keys ``not_recorded`` lists, and values ``counts`` gives; the rest are counted.
LISTED = 50

#: What a refusal's ``code`` becomes as a tool error.
CODES = {"unknown_key": "not_found", "key_shared": "refused", "empty_source": "refused",
         "already_recorded": "refused", "add_mismatch": "refused", "add_stale": "refused",
         "key_twice": "invalid_arguments", "empty_value": "invalid_arguments",
         "not_number": "invalid_arguments", "bad_item": "invalid_arguments",
         "bad_column": "invalid_arguments", "bad_table": "invalid_arguments",
         "bad_type": "invalid_arguments"}


@tool("xlsx_record_values",
      "Keep values you judged from a column's text, one per row by key, as dataset d:<id> "
      "{key, column, note}; each value's source is the cell judged. Lists rows with text "
      "but no value (not_recorded). Same id replaces it unless add.",
      {"doc": DOC,
       "table": string("Table (t:Products)."),
       "key": string("Key column; matched exactly."),
       "source": string("Column judged."),
       "column": string("Values' column name."),
       "id": string("Dataset id."),
       "values": array(obj({"key": string(), "value": string(or_number=True),
                            "note": string("Short reason.", optional=True)}),
                       "One per key.", min_items=1),
       "type": string("number: values must be numbers.", enum=["text", "number"],
                      optional=True),
       "add": boolean("Add to d:<id> instead of replacing it.", optional=True)},
      group=GROUP)
@handler
def xlsx_record_values(call: Any, doc: str, table: str, key: str, source: str, column: str,
                       id: str, values: list[dict[str, Any]], type: str = "text",  # noqa: A002
                       add: bool = False) -> Result:
    """``xlsx_record_values``'s handler (see the module's notes)."""
    workbook = workbook_of(call, doc)
    address = DATASET_PREFIX + id
    known = address in call.session.dataset_ids() or id in call.context.datasets
    warnings: list[str] = []
    add_to = None
    if add and known:
        add_to = call.dataset(address, field="id").dataset
    elif add:
        warnings.append(f"add: there is no {address} yet; recorded as a new dataset")
    try:
        recorded = record_values(workbook, table, key=key, source=source, column=column,
                                 values=values, type=type or "text", add_to=add_to)
    except RecordError as error:
        raise ToolError(CODES.get(error.code, "invalid_arguments"), error.message,
                        field=error.field, valid_options=error.valid_options,
                        details={"code": error.code, "problems": error.problems}) from None
    name = call.define_dataset(id, recorded.dataset, origin=recorded.address)
    warnings.extend(recorded.notes)
    data: dict[str, Any] = {"dataset": name, "from": recorded.address,
                            "columns": list(recorded.dataset.columns),
                            "recorded": recorded.recorded}
    if add_to is not None:
        data["added"] = recorded.added
    elif known:
        data["replaced"] = True
        warnings.append(f"{name} replaced: it holds this call's values only (add: true "
                        "adds to it)")
    counts = [(_key(value), n) for value, n in recorded.counts.items()]
    data["counts"] = dict(counts[:LISTED])
    if len(counts) > LISTED:
        data["counts_more"] = len(counts) - LISTED
    missing = recorded.not_recorded
    if missing:
        keys = [k for k, _ in missing[:LISTED]]
        whole = parse_address(recorded.address)
        assert isinstance(whole, TableAddress)
        listed: dict[str, Any] = {"count": len(missing), "keys": keys}
        columns = [recorded.key, recorded.source]
        if recorded.repeated.intersection(keys):
            # a key another row holds too (a heading, a spacer) names no one row to read:
            # the rows with text and those keys, by a query (round 54, R113)
            listed["query"] = {"doc": doc, "table": recorded.address, "columns": columns,
                               "where": [{"column": recorded.key, "op": "in",
                                          "values": list(dict.fromkeys(keys))},
                                         {"column": recorded.source, "op": "not_blank"}]}
        else:
            listed["read"] = {"doc": doc, "target": (replace(whole, keys=tuple(keys))
                                                     if len(keys) > 1
                                                     else replace(whole, key=keys[0])).text,
                              "key": recorded.key, "columns": columns}
        data["not_recorded"] = listed
    if recorded.no_key:
        data["no_key"] = {"count": len(recorded.no_key), "rows": list(recorded.no_key[:LISTED])}
    if recorded.shared:
        data["shared_keys"] = {"count": len(recorded.shared),
                               "rows": list(recorded.shared[:LISTED])}
    refused = recorded.refused
    if refused:
        # Per item (round 55, R118): the rest of the call is recorded.
        data["refused"] = [dict(p) for p in refused[:LISTED]]
        if len(refused) > LISTED:
            data["refused_more"] = len(refused) - LISTED
        places = ", ".join(f"values[{p['item']}]" for p in refused[:LISTED])
        more = f" and {len(refused) - LISTED} more" if len(refused) > LISTED else ""
        warnings.append(f"{len(refused)} value(s) not recorded, each key held by two or more "
                        f"rows with text to judge: {places}{more}; the rest are recorded. "
                        + shared_way("each refused item's", recorded.source))
    data["read"] = {"target": DatasetAddress(id).text, "sources": True}
    summary = f"Recorded {recorded.recorded} value(s) as {name} from {recorded.address}"
    if refused:
        summary += f"; {len(refused)} value(s) refused (refused)"
    if missing:
        summary += f"; {len(missing)} row(s) with text in {recorded.source} not recorded"
    return Result(summary=summary, data=data, total=recorded.recorded, warnings=warnings)


def _key(value: Any) -> str:
    """A value as a ``counts`` key: text as given, a whole number's digits."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return value if isinstance(value, str) else str(value)


TOOLS = [xlsx_record_values]
