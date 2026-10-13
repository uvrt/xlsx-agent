"""The workbook tools against Claude itself (marked ``provider``): the loaded core's counted
size, and a request with the definitions as the adapter sends them.

Skipped unless ``ANTHROPIC_API_KEY`` is set; run with ``python -m pytest -m provider``.
``ANTHROPIC_MODEL`` picks the model (default ``claude-sonnet-5-5``, the siblings' trial
model) and ``ANTHROPIC_WORKSPACE_ID``, when set, is sent as the ``anthropic-workspace-id``
header.  Plain HTTPS, no SDK, as in ooxml-edit's and the siblings' online checks; neither
the key nor any header is printed.

**The budget (ROADMAP.md, "Tools"):** what a request loads -- the core definitions, the
hidden tool-use prompt included -- is at most 6,000 tokens as count-tokens counts them (the
siblings' per-format budget).  The offline proxy and the all-definitions guard are in
``test_tools_schema.py``.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.adapters import anthropic_problems

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

pytestmark = [
    pytest.mark.provider,
    pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                       reason="ANTHROPIC_API_KEY is not set"),
]

MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")
API = "https://api.anthropic.com/v1"
#: Tokens the loaded core may add to a request, counted online.
CORE_COUNTED_BUDGET = 6000


def _post(path: str, body: dict) -> dict:
    headers = {"x-api-key": os.environ["ANTHROPIC_API_KEY"],
               "anthropic-version": "2023-06-01", "content-type": "application/json"}
    workspace = os.environ.get("ANTHROPIC_WORKSPACE_ID", "").strip()
    if workspace:
        headers["anthropic-workspace-id"] = workspace
    request = urllib.request.Request(f"{API}/{path}", data=json.dumps(body).encode(),
                                     method="POST", headers=headers)
    failure = None
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            reply = json.loads(response.read())
    except urllib.error.HTTPError as error:
        failure = f"{path}: HTTP {error.code}: {error.read().decode(errors='replace')}"
    if failure is not None:
        pytest.fail(failure, pytrace=False)   # outside the handler: no headers in a traceback
    if "usage" in reply:
        print(f"\n{path} usage: {json.dumps(reply['usage'])}")
    return reply


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def test_the_loaded_core_is_within_its_counted_budget(toolbox):
    messages = [{"role": "user", "content": "x"}]
    core = [d for d in toolbox.definitions("anthropic", groups="core") if "input_schema" in d]
    base = _post("messages/count_tokens", {"model": MODEL, "messages": messages})
    loaded = _post("messages/count_tokens", {"model": MODEL, "messages": messages,
                                             "tools": core})
    counted = loaded["input_tokens"] - base["input_tokens"]
    print(f"\n{MODEL}: the core, {len(core)} definitions, counts {counted} tokens")
    assert counted <= CORE_COUNTED_BUDGET, counted


def test_the_definitions_as_sent_are_accepted(toolbox):
    tools = toolbox.definitions("anthropic")      # core loaded, the rest deferred, strict planned
    assert anthropic_problems(tools) == []
    reply = _post("messages", {"model": MODEL, "max_tokens": 32, "tools": tools,
                               "system": toolbox.system_prompt(),
                               "output_config": {"effort": "low"},
                               "messages": [{"role": "user", "content": "Reply with: ready"}]})
    assert reply["type"] == "message"


# -- strict decoding keeps every field given -------------------------------------------------------


def _missing(given, got, where=""):
    """The dotted paths of ``given`` that ``got`` lacks: objects compared key by key, an
    array by its first item."""
    if isinstance(given, dict):
        if not isinstance(got, dict):
            return [where or "the arguments"]
        lost = []
        for key, value in given.items():
            path = f"{where}.{key}" if where else key
            lost += [path] if key not in got else _missing(value, got[key], path)
        return lost
    if isinstance(given, list) and given:
        if not isinstance(got, list) or not got:
            return [where]
        return _missing(given[0], got[0], where + "[]")
    return []


def _ask_for(name, schema, arguments):
    """A user message asking for one call with every argument, the optional ones listed
    first: a model writes the arguments it must give first, which is what lost the
    siblings' trial 3 its optional fields when they were listed before a required one."""
    required = set(schema.get("required", ()))
    order = [k for k in arguments if k not in required] + [k for k in arguments if k in required]
    lines = "\n".join(f"- {key}: {json.dumps(arguments[key])}" for key in order)
    return (f"This is a test of the tool interface; nothing is read. Call {name} exactly "
            f"once, with exactly these arguments and values, none dropped or changed, and "
            f"nothing else:\n{lines}")


def test_strict_tools_keep_every_field_given(toolbox):
    """Each strict tool, as the adapter sends it, gets one scripted request asking for a
    call that uses every optional field; every field must arrive (the siblings' finding S1
    of trial 3).

    Claude's strict plan (``ooxml_edit.tools.adapters.anthropic_strict_plan``) gives strict
    mode to writing tools only, and 0.1's tools all read, so none is sent strict yet: the
    test skips until 0.2's writing tools are, rather than passing on nothing."""
    from ooxml_edit.tools.schema import example_arguments

    strict = [d for d in toolbox.definitions("anthropic", defer=False) if d.get("strict")]
    if not strict:
        pytest.skip("no tool is sent strict: 0.1's tools all read (strict comes with 0.2)")
    lost = {}
    for definition in strict:
        arguments = example_arguments(definition["input_schema"])
        reply = _post("messages", {
            "model": MODEL, "max_tokens": 1500, "tools": [definition],
            "tool_choice": {"type": "auto"}, "output_config": {"effort": "low"},
            "messages": [{"role": "user", "content": _ask_for(
                definition["name"], definition["input_schema"], arguments)}]})
        calls = [b for b in reply["content"] if b["type"] == "tool_use"]
        assert [c["name"] for c in calls] == [definition["name"]], reply["content"]
        missing = _missing(arguments, calls[0]["input"])
        if missing:
            lost[definition["name"]] = missing
    print(f"\nstrict tools checked: {[d['name'] for d in strict]}")
    assert lost == {}, lost
