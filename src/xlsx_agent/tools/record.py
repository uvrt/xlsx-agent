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
so): send every value of the dataset in one call.

**Refused, nothing kept** (rule 2): a key no row has (keys are matched exactly; the rows
whose key is loosely equal by ooxml-edit's ``loose_key`` are given in ``like`` and as
``valid_options``, by position: read their key and send it as held), a key twice, a value
from an empty source cell, an empty value, a key column where two rows share a key; every
problem of the call at once in ``details.problems``, by place in ``values``.

**The result** (rule 4): ``recorded``, ``counts`` per value (at most :data:`LISTED`
values, the rest counted in ``counts_more``), ``not_recorded`` -- the rows with text in the
source column and no value: ``count``, the first :data:`LISTED` ``keys`` and ``read``, the
``xlsx_read`` call that gives their key and text (a key set; the rest of them counted, ask
again after recording) -- and ``no_key``, rows with source text but no key, by position.
The values themselves are not echoed: ``xlsx_read`` of ``d:<id>`` reads them.

**Content.**  The result carries keys and the values given; the summary and errors carry
counts, places and rows by position only.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any

from ooxml_edit.tools import DATASET_PREFIX, Result, ToolError, array, obj, string, tool

from ..address import DatasetAddress, TableAddress, parse_address
from ..record import RecordError, record_values
from .derive import workbook_of
from .docformat import handler
from .read import DOC

#: The group ``xlsx_record_values`` is in: deferred, loaded when a model looks for it.
GROUP = "xlsx_record"

#: The most keys ``not_recorded`` lists, and values ``counts`` gives; the rest are counted.
LISTED = 50

#: What a refusal's ``code`` becomes as a tool error.
CODES = {"unknown_key": "not_found", "duplicate_keys": "refused", "empty_source": "refused",
         "key_twice": "invalid_arguments", "empty_value": "invalid_arguments",
         "bad_item": "invalid_arguments", "bad_column": "invalid_arguments",
         "bad_table": "invalid_arguments"}


@tool("xlsx_record_values",
      "Keep values you judged from a column's text, one per row by key, as dataset d:<id> "
      "{key, column, note}; each value's source is the cell judged. Lists rows with text "
      "but no value (not_recorded). Same id replaces it.",
      {"doc": DOC,
       "table": string("Table (t:Products)."),
       "key": string("Key column; matched exactly."),
       "source": string("Column judged."),
       "column": string("Values' column name."),
       "id": string("Dataset id."),
       "values": array(obj({"key": string(), "value": string(),
                            "note": string("Short reason.", optional=True)}),
                       "One per key.", min_items=1)},
      group=GROUP)
@handler
def xlsx_record_values(call: Any, doc: str, table: str, key: str, source: str, column: str,
                       id: str, values: list[dict[str, Any]]) -> Result:  # noqa: A002
    """``xlsx_record_values``'s handler (see the module's notes)."""
    workbook = workbook_of(call, doc)
    try:
        recorded = record_values(workbook, table, key=key, source=source, column=column,
                                 values=values)
    except RecordError as error:
        raise ToolError(CODES.get(error.code, "invalid_arguments"), error.message,
                        field=error.field, valid_options=error.valid_options,
                        details={"code": error.code, "problems": error.problems}) from None
    address = DATASET_PREFIX + id
    replaced = address in call.session.dataset_ids() or id in call.context.datasets
    name = call.define_dataset(id, recorded.dataset, origin=recorded.address)
    warnings = list(recorded.notes)
    data: dict[str, Any] = {"dataset": name, "from": recorded.address,
                            "columns": list(recorded.dataset.columns),
                            "recorded": recorded.recorded}
    if replaced:
        data["replaced"] = True
        warnings.append(f"{name} replaced: it holds this call's values only")
    counts = list(recorded.counts.items())
    data["counts"] = dict(counts[:LISTED])
    if len(counts) > LISTED:
        data["counts_more"] = len(counts) - LISTED
    missing = recorded.not_recorded
    if missing:
        keys = [k for k, _ in missing[:LISTED]]
        whole = parse_address(recorded.address)
        assert isinstance(whole, TableAddress)
        listed: dict[str, Any] = {"count": len(missing), "keys": keys}
        listed["read"] = {"doc": doc, "target": (replace(whole, keys=tuple(keys)) if len(keys) > 1
                                     else replace(whole, key=keys[0])).text,
                          "key": recorded.key, "columns": [recorded.key, recorded.source]}
        data["not_recorded"] = listed
    if recorded.no_key:
        data["no_key"] = {"count": len(recorded.no_key), "rows": list(recorded.no_key[:LISTED])}
    data["read"] = {"target": DatasetAddress(id).text, "sources": True}
    summary = f"Recorded {recorded.recorded} value(s) as {name} from {recorded.address}"
    if missing:
        summary += f"; {len(missing)} row(s) with text in {recorded.source} not recorded"
    return Result(summary=summary, data=data, total=recorded.recorded, warnings=warnings)


TOOLS = [xlsx_record_values]
