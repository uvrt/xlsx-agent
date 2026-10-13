"""Golden transcripts: three read-and-derive tasks, as a model did them, replayed offline.

``g1`` summarises a sales table by region with ``xlsx_derive`` and reads the dataset back;
``g2`` compares a generated workbook with the copy a reviewer changed; ``g3`` checks a
workbook against a declared schema.  Each transcript in ``golden/transcripts`` was recorded
by ``tools/record_goldens.py`` (claude-haiku-5-5, 2026-10-13): the model's calls, in order,
with what every call returned (``Result.to_json()``), and its final answer for the reader.
Replaying the calls on a fresh session (fixed clock) must give the same results, value for
value -- including the calls that failed, whose errors are part of the record.  No provider
is called here.

The inputs are rebuilt by ``golden_inputs.py`` (every value and date fixed there) and must
hash as they did when the transcripts were recorded: a package by its entries everywhere
(``content_sha256``), and by its bytes too where deflate is zlib's own (``zip_content.py``).
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
from dataclasses import replace
from pathlib import Path

import pytest

from ooxml_edit.tools import Toolbox

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

import golden_inputs
from zip_content import content_sha256, stock_deflate, zlib_ng

HERE = Path(__file__).resolve().parent
TRANSCRIPTS = sorted((HERE / "golden" / "transcripts").glob("*.json"))
#: The clock the transcripts were recorded on (tools/record_goldens.py).
CLOCK = dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)
STOCK_DEFLATE = stock_deflate()


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def test_deflate_is_zlibs_unless_this_python_says_zlib_ng():
    """The bytes are skipped only where they cannot match: never by a probe gone wrong."""
    assert STOCK_DEFLATE or zlib_ng()


@pytest.mark.parametrize("path", TRANSCRIPTS, ids=lambda path: path.stem)
def test_a_golden_transcript_replays_to_the_same_results(path):
    transcript = _load(path)
    task = transcript["task"]
    builders = golden_inputs.BUILDERS[task]
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=lambda: CLOCK)
        for item in transcript["inputs"]:
            data = builders[item["file"]]()
            assert content_sha256(data) == item["content_sha256"], \
                f"{item['file']} is not the input the transcript was recorded with"
            if STOCK_DEFLATE:
                assert hashlib.sha256(data).hexdigest() == item["sha256"], item["file"]
            assert item["as"] == "document"
            session.open(data, item["file"])
        for index, step in enumerate(transcript["calls"]):
            result = toolbox.dispatch(session, step["tool"], step["arguments"])
            got = json.loads(json.dumps(result.to_json()))
            if step["tool"] == "xlsx_compare" and got["ok"]:
                got = _as_recorded_before_x10(got)
            assert got == step["result"], (index, step["tool"])


def _as_recorded_before_x10(result: dict) -> dict:
    """An ``xlsx_compare`` result as it was before X10, when g2 was recorded: X10 added each
    table's summary (``columns``, ``added``, ``removed`` and their ``more_``; ``cut``'s
    ``summaries``) and names a difference by its address by key (``at``,
    ``t:Stock/SKU-0002/Price``) in place of ``table``, ``key`` and ``column`` -- a row's
    ``at`` was its address by position (``t:Stock/#8``), read back here from its first
    cell's row.  Converted back, the replay still checks every value recorded, and the
    model's answer stands on the same facts; the summary is tested in test_compare.  Not
    re-recorded: recording calls a model, and is the main session's to decide."""
    from xlsx_agent.address import RangeAddress, TableAddress

    data = result["data"]
    starts = {}
    for table in data["tables"]:
        for name in ("columns", "added", "removed", "more_columns", "more_added",
                     "more_removed"):
            table.pop(name, None)
        for side in ("a", "b"):
            area = RangeAddress.parse(table[f"range_{side}"]).area
            starts[side] = area.first.row + 1       # one header row in g2
    if "cut" in data:
        data["cut"].pop("summaries", None)
        if not data["cut"]:
            del data["cut"]
    old = []
    for item in data["differences"]:
        if "at" in item and ("cell" in item or "row" in item):
            address = TableAddress.parse(item.pop("at"))
            named = {"table": address.whole.text, "key": address.key}
            if "cell" in item:
                item = {**named, "cell": item.pop("cell"), "column": address.column, **item}
            else:
                side = "b" if item["row"] == "added" else "a"
                row = RangeAddress.parse(item["cells"][0]["a"]).area.first.row
                position = replace(address.whole, position=row - starts[side] + 1)
                item = {**named, "row": item.pop("row"), "at": position.text, **item}
        old.append(item)
    data["differences"] = old
    return result


def test_g2s_comparison_now_opens_on_its_summary():
    """What X10 adds to g2's comparison: each column's changes counted, rows by key."""
    transcript = _load(HERE / "golden" / "transcripts" / "g2.json")
    builders = golden_inputs.BUILDERS[transcript["task"]]
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=lambda: CLOCK)
        for item in transcript["inputs"]:
            session.open(builders[item["file"]](), item["file"])
        step = next(s for s in transcript["calls"] if s["tool"] == "xlsx_compare")
        result = toolbox.dispatch(session, "xlsx_compare", step["arguments"])
    (table,) = result.data["tables"]
    assert table["columns"] == {"Qty": {"changed": 1},
                                "Price": {"changed": 1, "up": 1,
                                          "max_up": "t:Stock/SKU-0002/Price +0.01"}}
    assert table["added"] == ["t:Stock/SKU-0009"] and table["removed"] == ["t:Stock/SKU-0006"]


def test_the_goldens_cover_the_three_tasks_in_few_calls():
    transcripts = {_load(p)["task"]: _load(p) for p in TRANSCRIPTS}
    assert sorted(transcripts) == sorted(golden_inputs.BUILDERS)
    names = {tool.name for tool in TOOLS}
    for task, transcript in transcripts.items():
        assert set(transcript) == {"task", "recorded", "prompt", "inputs", "calls", "answer"}
        assert transcript["recorded"]["how"] in ("recorded", "hand-written")
        assert 1 <= len(transcript["calls"]) <= 10, task
        assert all(step["tool"] in names for step in transcript["calls"])
        assert all(set(step) == {"tool", "arguments", "result"} for step in transcript["calls"])


def test_each_task_ends_on_the_call_it_is_about():
    """The last successful call of each task is the one the task asks for: the dataset read
    back (g1), the comparison (g2), the check (g3)."""
    wanted = {"g1-sales-by-region": ("xlsx_read", "d:by_region"),
              "g2-compare-review": ("xlsx_compare", None),
              "g3-validate-orders": ("xlsx_validate_data", None)}
    for path in TRANSCRIPTS:
        transcript = _load(path)
        tool, target = wanted[transcript["task"]]
        done = [s for s in transcript["calls"] if s["result"]["ok"]]
        assert done[-1]["tool"] == tool, transcript["task"]
        if target:
            assert done[-1]["arguments"]["target"] == target
