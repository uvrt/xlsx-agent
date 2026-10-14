"""Record the golden transcripts (tests/golden/transcripts) by running each task on Claude.

    ANTHROPIC_API_KEY=... python tools/record_goldens.py --spend spend.txt --work DIR [task ...]

With no task named, all three are recorded; name one (``g2-compare-review``) to record it
alone, the others left as they are.

Each task's inputs are built by ``tests/golden_inputs.py`` and opened in a fresh session
(fixed clock); the model gets the tools as ``Toolbox.definitions("anthropic")`` sends them,
the format's system prompt, and the task's prompt, and its calls are run until it answers.
What is kept is the transcript a replay needs: the inputs' hashes, every call's tool and
arguments, and what it returned (``Result.to_json()``), plus the model's final answer for
the reader.  No key, header or request is written anywhere.

Spending is capped (``CAPS``) over every run together: the usage of each response is added
to the ``--spend`` file (numbers only), and a request is not sent when it could cross a
cap; a task that cannot finish within what is left is reported and not written.  Prompt
caching keeps a run cheap: the tools carry ``cache_control`` on the last core definition
(the adapter's own), the system prompt one more, and the newest message one, so each turn
reads the conversation so far from the cache.  Plain HTTPS, no SDK, as the online test.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
TESTS = HERE.parent / "tests"
sys.path.insert(0, str(TESTS))

from ooxml_edit.tools import Toolbox  # noqa: E402

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

import golden_inputs  # noqa: E402
from zip_content import content_sha256  # noqa: E402

API = "https://api.anthropic.com/v1/messages"
MODEL = "claude-haiku-5-5"
#: The fixed clock every session runs on (test_golden.py's), whatever day a task is recorded.
CLOCK = dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)
MAX_TOKENS = 512
MAX_TURNS = 8
#: Over every recording run together, summed from each response's usage.
CAPS = {"uncached": 40_000, "read": 200_000, "output": 8_000, "requests": 30}
#: What one more request is assumed to cost at most, before it is sent: new input (a turn's
#: tool results, or the tools and system prompt on a cold cache), cache reads, output.
WORST = {"uncached": 8_000, "read": 20_000, "output": MAX_TOKENS}

#: The schema g3 is checked against, as an application would declare it.
ORDERS_SCHEMA = {"tables": [{"table": "t:Orders", "key": ["Order"], "columns": [
    {"name": "Customer", "type": "text", "required": True},
    {"name": "Status", "enum": ["Open", "Shipped", "Closed"]},
    {"name": "Qty", "type": "number", "min": "1"},
    {"name": "Discount", "type": "number", "min": "0", "max": "0.5"}]}]}

PROMPTS = {
    "g1-sales-by-region": (
        "The workbook sales-july.xlsx is open as d1. Summarise its sales by region: the "
        "total revenue and the number of orders for each region, the largest revenue "
        "first, revenue rounded to 2 decimals, kept as the dataset by_region. Once it is "
        "kept, read d:by_region back with its sources, and report each region's figures "
        "with their addresses."),
    "g2-compare-review": (
        "d1 is inventory-generated.xlsx, as a tool wrote it, and d2 is "
        "inventory-reviewed.xlsx, the copy a reviewer changed. Compare the Stock table of "
        "the two by SKU and report what the reviewer changed, with addresses."),
    "g3-validate-orders": (
        "orders.xlsx is open as d1. Check it against this schema (pass it as the schema "
        "object) and report every problem with its address: "
        + json.dumps(ORDERS_SCHEMA)),
}


# -- spending ------------------------------------------------------------------------------


def read_spend(path: Path) -> dict[str, int]:
    """The tally so far: ``uncached creation read output`` per request, one line each."""
    total = {"uncached": 0, "read": 0, "output": 0, "requests": 0}
    if path.exists():
        for line in path.read_text().splitlines():
            fields = line.split()
            if len(fields) == 4 and all(f.isdigit() for f in fields):
                uncached, creation, read, output = map(int, fields)
                total["uncached"] += uncached + creation
                total["read"] += read
                total["output"] += output
                total["requests"] += 1
    return total


def can_send(total: dict[str, int]) -> str | None:
    """Why one more request could cross a cap, or ``None``."""
    if total["requests"] + 1 > CAPS["requests"]:
        return "requests"
    for key in ("uncached", "read", "output"):
        if total[key] + WORST[key] > CAPS[key]:
            return key
    return None


def add_spend(path: Path, usage: dict) -> None:
    line = (f"{usage.get('input_tokens', 0)} {usage.get('cache_creation_input_tokens', 0)} "
            f"{usage.get('cache_read_input_tokens', 0)} {usage.get('output_tokens', 0)}\n")
    with path.open("a") as handle:
        handle.write(line)


# -- the request -----------------------------------------------------------------------------


def post(body: dict) -> dict:
    headers = {"x-api-key": os.environ["ANTHROPIC_API_KEY"],
               "anthropic-version": "2023-06-01", "content-type": "application/json"}
    request = urllib.request.Request(API, data=json.dumps(body).encode(), method="POST",
                                     headers=headers)
    failure = None
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            return json.loads(response.read())
    except urllib.error.HTTPError as error:
        failure = f"HTTP {error.code}: {error.read().decode(errors='replace')}"
    raise SystemExit(failure)        # outside the handler: no headers in a traceback


def cached_last(messages: list[dict]) -> list[dict]:
    """The messages with ``cache_control`` on the newest one's last block (a copy)."""
    out = [dict(m) for m in messages]
    content = out[-1]["content"]
    if isinstance(content, str):
        content = [{"type": "text", "text": content}]
    content = [dict(block) for block in content]
    content[-1]["cache_control"] = {"type": "ephemeral"}
    out[-1]["content"] = content
    return out


def record(task: str, root: Path, spend: Path) -> dict | None:
    golden_inputs.build(task, root)
    folder = root / task
    inputs = []
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=lambda: CLOCK)
        for name in golden_inputs.BUILDERS[task]:
            data = (folder / name).read_bytes()
            session.open(data, name)
            inputs.append({"file": name, "as": "document",
                           "sha256": hashlib.sha256(data).hexdigest(),
                           "content_sha256": content_sha256(data)})
        tools = toolbox.definitions("anthropic")
        system = [{"type": "text", "text": toolbox.system_prompt(),
                   "cache_control": {"type": "ephemeral"}}]
        messages: list[dict] = [{"role": "user", "content": PROMPTS[task]}]
        calls = []
        answer = None
        for _ in range(MAX_TURNS):
            total = read_spend(spend)
            over = can_send(total)
            if over:
                print(f"{task}: stopped before a request that could cross the {over} cap")
                return None
            reply = post({"model": MODEL, "max_tokens": MAX_TOKENS, "tools": tools,
                          "system": system, "messages": cached_last(messages),
                          "thinking": {"type": "disabled"},
                          "output_config": {"effort": "low"}})
            add_spend(spend, reply.get("usage", {}))
            content = [block for block in reply["content"] if block["type"] in ("text", "tool_use")]
            messages.append({"role": "assistant", "content": content})
            uses = [block for block in content if block["type"] == "tool_use"]
            if not uses:
                answer = "\n".join(b["text"] for b in content if b["type"] == "text")
                break
            results = []
            for use in uses:
                result = toolbox.dispatch(session, use["name"], use["input"])
                calls.append({"tool": use["name"], "arguments": use["input"],
                              "result": json.loads(json.dumps(result.to_json()))})
                results.append((use["id"], result))
            messages.append(toolbox.render_results("anthropic", results))
        if answer is None:
            print(f"{task}: no answer within {MAX_TURNS} turns")
            return None
    return {"task": task, "recorded": {"how": "recorded", "model": MODEL,
                                       "date": dt.date.today().isoformat()},
            "prompt": PROMPTS[task], "inputs": inputs, "calls": calls, "answer": answer}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("tasks", nargs="*", default=list(PROMPTS))
    parser.add_argument("--spend", type=Path, required=True)
    parser.add_argument("--work", type=Path, required=True, help="where the inputs are built")
    parser.add_argument("--out", type=Path, default=TESTS / "golden" / "transcripts")
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    for task in args.tasks:
        transcript = record(task, args.work, args.spend)
        if transcript is None:
            raise SystemExit(1)
        path = args.out / f"{task.split('-')[0]}.json"
        path.write_text(json.dumps(transcript, indent=1, ensure_ascii=False) + "\n",
                        encoding="utf-8")
        print(f"{task}: {len(transcript['calls'])} calls -> {path.name}")
        print(f"tally: {read_spend(args.spend)}")


if __name__ == "__main__":
    main()
