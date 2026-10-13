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
            assert got == step["result"], (index, step["tool"])


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
