"""Golden transcripts: three read-and-derive tasks, as a model did them, replayed offline.

``g1`` summarises a sales table by region with ``xlsx_derive`` and reads the dataset back;
``g2`` compares a generated workbook with the copy a reviewer changed; ``g3`` checks a
workbook against a declared schema.  Each transcript in ``golden/transcripts`` was recorded
by ``tools/record_goldens.py`` (claude-haiku-5-5; g1 and g3 2026-10-13, g2 re-recorded
2026-10-14 after X10 changed ``xlsx_compare``'s result): the model's calls, in order,
with what every call returned (``Result.to_json()``), and its final answer for the reader.
Replaying the calls on a fresh session (fixed clock) must give the same results, value for
value -- including the calls that failed, whose errors are part of the record -- once what
the tools have added since the recording is applied to the recorded results
(:data:`SINCE_RECORDED`: each addition named, its values written out here, never derived
from the replay).  No provider is called here.

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
#: The clock the transcripts were recorded on (tools/record_goldens.py), the same fixed
#: instant for all three whatever day each was recorded.
CLOCK = dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)
STOCK_DEFLATE = stock_deflate()


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


#: g2's rows by key and the text of the column that labels them (Item), as g2's own
#: ``describe`` calls recorded them.
G2_LABELS = {"t:Stock/SKU-0002": "Wood screw 4x40", "t:Stock/SKU-0005": "Spirit level 60 cm",
             "t:Stock/SKU-0006": "Masking tape 25 mm", "t:Stock/SKU-0009": "Paint brush 50 mm"}


def _compare_labels(column: str, labels: dict[str, str]):
    """T19b (after g2 was recorded): ``xlsx_compare`` names each table's label ``column``,
    after its key, and gives each row and cell difference its row's label, after ``at``."""
    def apply(result: dict) -> dict:
        data = result["data"]
        for table in data["tables"]:
            items = list(table.items())
            at = [name for name, _ in items].index("key") + 1
            table.clear()
            table.update(items[:at] + [("label", column)] + items[at:])
        for index, difference in enumerate(data["differences"]):
            row = "/".join(difference["at"].split("/")[:2])
            items = list(difference.items())
            at = [name for name, _ in items].index("at") + 1
            data["differences"][index] = dict(items[:at] + [("label", labels[row])]
                                              + items[at:])
        return result
    return apply


#: What each tool's results have gained since a transcript was recorded, by task and tool:
#: applied to the recorded result before the replay's is compared with it, so a transcript
#: stays as the model saw it (its answer was written from the result as recorded) and is
#: not recorded again for an addition.
SINCE_RECORDED = {"g2-compare-review": {"xlsx_compare": _compare_labels("Item", G2_LABELS)}}


def _expected(task: str, step: dict) -> dict:
    expected = json.loads(json.dumps(step["result"]))
    change = SINCE_RECORDED.get(task, {}).get(step["tool"])
    return change(expected) if change and expected["ok"] else expected


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
            expected = _expected(task, step)
            assert got == expected, (index, step["tool"])
            assert json.dumps(got) == json.dumps(expected), (index, "order")


def test_g2s_comparison_opens_on_its_summary():
    """g2's comparison as recorded gives each column's changes counted and the rows by key
    first, the summary X10 added; the replay above checks every value of it."""
    transcript = _load(HERE / "golden" / "transcripts" / "g2.json")
    step = next(s for s in transcript["calls"] if s["tool"] == "xlsx_compare")
    (table,) = step["result"]["data"]["tables"]
    assert list(table)[-3:] == ["columns", "added", "removed"]
    assert table["columns"] == {"Qty": {"changed": 1},
                                "Price": {"changed": 1, "up": 1,
                                          "max_up": "t:Stock/SKU-0002/Price +0.01"}}
    assert table["added"] == ["t:Stock/SKU-0009"] and table["removed"] == ["t:Stock/SKU-0006"]
    assert all("at" in item for item in step["result"]["data"]["differences"])


def test_g2s_replay_labels_the_row_its_answer_misnamed():
    """g2's answer, written from the comparison as recorded (rows by key alone), named
    SKU-0005 by another row's item; the replay now gives each row its label (T19b)."""
    transcript = _load(HERE / "golden" / "transcripts" / "g2.json")
    assert "SKU-0005** (Wall plug" in transcript["answer"]
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=lambda: CLOCK)
        for item in transcript["inputs"]:
            session.open(golden_inputs.BUILDERS[transcript["task"]][item["file"]](), item["file"])
        step = next(s for s in transcript["calls"] if s["tool"] == "xlsx_compare")
        data = toolbox.dispatch(session, step["tool"], step["arguments"]).data
    assert data["tables"][0]["label"] == "Item"
    assert {d["at"]: d["label"] for d in data["differences"]} == {
        "t:Stock/SKU-0009": "Paint brush 50 mm", "t:Stock/SKU-0006": "Masking tape 25 mm",
        "t:Stock/SKU-0002/Price": "Wood screw 4x40", "t:Stock/SKU-0005/Qty": "Spirit level 60 cm"}


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
