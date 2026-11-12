"""Values a caller judged, one per row, kept as a dataset traced to the cells judged (round 53,
R108): what ``xlsx_record_values`` keeps as ``d:<id>``::

    recorded = wb.record_values("t:Shelf", key="Ref", source="notes", column="outcome",
                                values=[{"key": "S-01", "value": "done"},
                                        {"key": "S-02", "value": "partly done",
                                         "note": "says 'en partie'"}])
    recorded.dataset            # columns (Ref, outcome, note), a row per key recorded
    recorded.dataset.sources[0][1]    # ("t:Shelf/S-01/notes",): the cell judged
    recorded.not_recorded       # rows with text in notes and no value given

An agent reads a column of free text (``notes``, worded many ways), decides a value for
each row (``done``, ``partly done``) and wants it kept like any figure: in a dataset whose
every value names where it came from.  The library cannot judge text; it can keep the
judgement honest: one row per key, each value tied to the source cell it was judged from,
and every row whose text was not judged listed.

**What is kept.**  A row per recorded key, in the table's row order (not the order given):
the key column's value (its name and type as in the table), ``column`` (the value as given,
text) and ``note`` (the caller's short reason, blank when none).  ``keys`` is 1, so a row is
read by key (``d:outcome/S-01``) as a derive's grouped rows are.  Sources: the key's own
cell for the key, the source cell for the value and the note -- by key where the table has
a key column (``t:Shelf/S-01/notes``), else the cell's address, as
:func:`~xlsx_agent.dataset.table_dataset` writes them -- so ``xlsx_trace`` of a row reaches
the row judged, and a value placed in a deck or document cites the cell.

**Refused before anything is kept** (:class:`RecordError`, every problem of the call at
once, each named by its place in ``values`` and the row by position, never by content): a
key no row has (``unknown_key``; the rows whose key is loosely equal -- ooxml-edit's
:func:`~ooxml_edit.tools.loose_key`: case, white space, separators, leading zeros -- are
suggested by position, ``t:Shelf/#4/Ref``), a key given twice (``key_twice``), a value
whose source cell is empty (``empty_source``: nothing was there to judge), an empty value
(``empty_value``: leave the key out instead).  A key column where two rows share a key is
refused first (``duplicate_keys``), with those rows: a judgement per key needs keys that
name one row.  The output ``column`` may not be the key column's name or ``note``.

**Listed, never refused.**  :attr:`Recorded.not_recorded`: the rows whose source cell holds
text (anything but blank or white space) and got no value, by key, in row order;
:attr:`Recorded.no_key`: rows with source text but no key (a blank key cell), by position.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, Mapping, Sequence

from ooxml_edit.tools import Dataset, loose_key

from .address import TableAddress, parse_address
from .dataset import table_dataset

if TYPE_CHECKING:
    from .workbook import Workbook

#: The name of the column holding each value's note.
NOTE = "note"

#: The most problems a refusal names; the rest are counted.
NAMED = 20

#: The most rows suggested for one unknown key.
SUGGESTED = 3


class RecordError(ValueError):
    """A ``record_values`` call refused before anything was kept: ``code`` the first
    problem's kind (``unknown_key``, ``key_twice``, ``empty_source``, ``empty_value``,
    ``duplicate_keys``, ``bad_column``), ``field`` the argument at fault, ``problems``
    every problem as ``{"code", "item", "row"}`` (``item`` the place in ``values``, from 0;
    ``row`` the table row by position) with ``like`` (rows whose key is loosely equal) for
    an unknown key; ``valid_options`` the rows to read.  The message names places and rows,
    never a key or a value."""

    def __init__(self, code: str, message: str, *, field: str = "values",
                 problems: Sequence[Mapping[str, Any]] = (),
                 valid_options: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.field = field
        self.problems = [dict(p) for p in problems]
        self.valid_options = list(valid_options)


@dataclass(frozen=True)
class Recorded:
    """What :func:`record_values` kept: ``dataset`` (ooxml-edit's :class:`Dataset`),
    ``address`` the table's, ``key``/``source``/``column`` the column names,
    ``not_recorded`` the rows with source text and no value as ``(key, source cell)``,
    ``no_key`` the rows with source text and no key (``t:Shelf/#7``), ``counts`` how many
    rows got each value, in the order first met; ``notes`` what was read as meant."""

    dataset: Dataset
    address: str
    key: str
    source: str
    column: str
    not_recorded: tuple[tuple[str, str], ...] = ()
    no_key: tuple[str, ...] = ()
    counts: Mapping[str, int] = field(default_factory=dict)
    notes: tuple[str, ...] = ()

    @property
    def recorded(self) -> int:
        return len(self.dataset.rows)


def _text(value: Any) -> str | None:
    """A given key or value as text: a whole number's digits, as a key cell's are
    (:func:`~xlsx_agent.read.regions.key_text`); ``None`` for nothing."""
    if value is None:
        return None
    if isinstance(value, float) and not isinstance(value, bool) and value.is_integer():
        return str(int(value))
    return value if isinstance(value, str) else str(value)


def _blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def record_values(workbook: "Workbook", table: "str | TableAddress", *, key: str, source: str,
                  column: str, values: Sequence[Mapping[str, Any]]) -> Recorded:
    """Keep ``values`` (``[{"key", "value", "note"?}]``), judged by the caller from the
    ``source`` column of ``table``'s rows named by ``key``, as a dataset (see the module's
    notes); :class:`RecordError` for a call refused, :class:`~xlsx_agent.NameNotFound`
    for a table or column there is not."""
    notes: list[str] = []
    target = parse_address(table) if isinstance(table, str) else table
    if not isinstance(target, TableAddress):
        raise RecordError("bad_table", "table names a table or region (t:Shelf)", field="table")
    whole = replace(target, column=None, key=None, keys=None, position=None, any_row=False,
                    any_column=False, row_range=None)
    if whole != target:
        notes.append(f"table: {target.text} read as the whole table {whole.text}; key and "
                     "source name its columns")
    region = workbook.region(whole)
    key_name = region.columns[region.column_index(key)]
    source_name = region.columns[region.column_index(source)]
    address = region.address
    if not isinstance(column, str) or not column.strip():
        raise RecordError("bad_column", "column names the output column (outcome)",
                          field="column")
    if column.casefold() in (key_name.casefold(), NOTE):
        raise RecordError("bad_column", f"column {column!r} is taken in the dataset: its "
                          f"columns are the key ({key_name!r}), yours and {NOTE!r}; name "
                          "it otherwise", field="column")
    index = workbook.regions.key_index(region, key_name)
    if index.duplicates:
        shared = [replace(address, position=p, column=key_name).text
                  for positions in index.duplicates.values() for p in positions]
        raise RecordError("duplicate_keys", f"{len(index.duplicates)} key(s) of column "
                          f"{key_name!r} are held by more than one row of {address.text} "
                          f"({len(shared)} rows): a value per key needs keys naming one row; "
                          "the rows are the valid options", field="key",
                          problems=[{"code": "duplicate_keys", "row": r}
                                    for r in shared[:NAMED]],
                          valid_options=shared[:50])
    keys_ds = table_dataset(workbook, replace(address, column=key_name))
    source_ds = table_dataset(workbook, replace(address, column=source_name))
    assert keys_ds.sources is not None and source_ds.sources is not None
    position_of: dict[str, int] = {}
    for position, k in enumerate(index.keys, 1):
        if k is not None:
            position_of[k] = position

    def row(position: int, name: str) -> str:
        return replace(address, position=position, column=name).text

    problems: list[dict[str, Any]] = []
    given: dict[int, tuple[str, str | None]] = {}       # position -> (value, note)
    item_of: dict[int, int] = {}
    loose: dict[tuple, list[int]] | None = None
    for n, item in enumerate(values):
        if not isinstance(item, Mapping):
            problems.append({"code": "bad_item", "item": n})
            continue
        wanted = _text(item.get("key"))
        value = _text(item.get("value"))
        note = _text(item.get("note"))
        found = position_of.get(wanted) if wanted is not None else None
        if found is None:
            if loose is None:
                loose = {}
                for p, k in enumerate(index.keys, 1):
                    if k is not None:
                        loose.setdefault(loose_key(k), []).append(p)
            like = loose.get(loose_key(wanted), []) if wanted else []
            problems.append({"code": "unknown_key", "item": n,
                             "like": [row(p, key_name) for p in like[:SUGGESTED]]})
            continue
        position = found
        if position in item_of:
            problems.append({"code": "key_twice", "item": n, "row": row(position, key_name),
                             "first": item_of[position]})
            continue
        item_of[position] = n
        if _blank(source_ds.rows[position - 1][0]):
            problems.append({"code": "empty_source", "item": n,
                             "row": row(position, source_name)})
            continue
        if _blank(value):
            problems.append({"code": "empty_value", "item": n, "row": row(position, key_name)})
            continue
        assert value is not None
        given[position] = (value, None if _blank(note) else note)
    if problems:
        raise _refusal(problems, address.text, key_name, source_name)
    rows: list[tuple[Any, ...]] = []
    sources: list[tuple[tuple[str, ...], ...]] = []
    not_recorded: list[tuple[str, str]] = []
    no_key: list[str] = []
    counts: dict[str, int] = {}
    for position in range(1, len(index.keys) + 1):
        k = index.keys[position - 1]
        cell = source_ds.sources[position - 1][0]
        if position in given:
            value, note = given[position]
            rows.append((keys_ds.rows[position - 1][0], value, note))
            sources.append((keys_ds.sources[position - 1][0], cell, cell))
            counts[value] = counts.get(value, 0) + 1
        elif not _blank(source_ds.rows[position - 1][0]):
            if k is None:
                no_key.append(replace(address, position=position).text)
            else:
                not_recorded.append((k, cell[0] if cell else row(position, source_name)))
    key_type = keys_ds.types[0] if keys_ds.types else None
    dataset = Dataset((key_name, column, NOTE), tuple(rows), sources=tuple(sources), keys=1,
                      types=(key_type, "text", "text"), cache=_cache(keys_ds, source_ds))
    return Recorded(dataset, address.text, key_name, source_name, column,
                    tuple(not_recorded), tuple(no_key), counts, tuple(notes))


def _cache(*read: Dataset) -> dict[str, int] | None:
    """The formula cells the key and source columns were read from, by cache status."""
    out: dict[str, int] = {}
    for dataset in read:
        for status, count in (dataset.cache or {}).items():
            out[status] = out.get(status, 0) + count
    return out or None


_SAID = {"unknown_key": "unknown key", "key_twice": "key given twice",
         "empty_source": "source cell empty (nothing there was judged)",
         "empty_value": "empty value (leave the key out to record nothing for it)",
         "bad_item": "not an object {key, value, note}"}


def _refusal(problems: list[dict[str, Any]], table: str, key: str,
             source: str) -> RecordError:
    """One refusal naming every problem by kind and place (rule 2), the way out each."""
    kinds: dict[str, list[int]] = {}
    for p in problems:
        kinds.setdefault(p["code"], []).append(p["item"])
    parts = []
    for code, items in kinds.items():
        shown = ", ".join(f"values[{i}]" for i in items[:NAMED])
        more = f" and {len(items) - NAMED} more" if len(items) > NAMED else ""
        parts.append(f"{_SAID[code]} ({len(items)}): {shown}{more}")
    ways = []
    if "unknown_key" in kinds:
        ways.append(f"keys are matched exactly as column {key!r} holds them; problems' like "
                    "names the rows whose key is loosely equal (case, spaces, separators, "
                    "leading zeros): read their key and send it as held")
    if "empty_source" in kinds:
        ways.append(f"leave out the keys whose {source!r} is empty")
    options = [r for p in problems for r in p.get("like", [])][:50]
    return RecordError(problems[0]["code"], f"nothing was recorded from {table}: "
                       + "; ".join(parts) + (". " + "; ".join(ways) if ways else ""),
                       problems=problems[:NAMED * 2], valid_options=options)
