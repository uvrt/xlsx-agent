"""The xlsx tool definitions against each provider's rules, and the token budget.

Offline: every canonical schema is in the common strict subset (``ooxml_edit.tools`` refuses
a tool that is not, when it is made); the Anthropic, OpenAI Responses and Chat Completions
adapters' output passes each provider's documented strict rules; no parameter is a path;
the groups are as the roadmap lists them.

**The budget** (ROADMAP.md, "Tools"; the siblings' per-format budget): what a request
loads -- the core -- is at most 6,000 tokens as Anthropic's count-tokens counts them.
Offline, characters / 3.5 of each definition's compact JSON stands in for it, times 1.48,
the ratio of the counted to the estimated size the siblings measured on Sonnet 5.5 (T5 and
T5b).  The online count is R14's (``test_tools_online.py``, marked ``provider``).

R6 (ooxml-edit 0.18.0): the core is the shared ``open_document`` and ``describe``, 274
estimated, about 406 counted by the proxy.  R7 adds ``xlsx_read`` and ``xlsx_profile`` to
the core (the read tools are loaded): 912 estimated, about 1,350 counted.  R8 adds
``xlsx_query``: 1,366 estimated, about 2,022 counted.  R9 adds the shared ``find_text``
(its one definition, 289 estimated, about 428 counted; nothing of the workbook's own):
1,654 estimated, about 2,448 counted.  R11 adds ``xlsx_derive`` (789 estimated) and
``xlsx_read``'s dataset side (``doc`` optional, ``sources``): 2,488 estimated, about 3,682
counted.  R12 adds ``xlsx_compare`` (329 estimated, about 487 counted): 2,806 estimated,
about 4,153 counted, so the check tools stay in the core (ROADMAP.md, "Tools").  R13 adds
``xlsx_validate_data`` (630 estimated, about 932 counted): 3,425 estimated, about 5,069
counted, still under 6,000, so it is loaded too and no ``xlsx_check`` group is needed.
X4 and X5 add ``first``, ``last``, ``date_as_text``, ``words_gt`` and a few words of
description: 3,452 estimated, about 5,109 counted.  X2 adds only ``xlsx_read``'s mention
of ``*`` (``t:Products/*/Price``): 3,462 estimated, about 5,124 counted.  X10 describes
``xlsx_compare``'s summary per column (330 estimated): 3,474 estimated, about 5,142 counted.
Every definition (``close_document`` and ``read_blob`` deferred in ``shared_misc``, and the
tool-search tool) is 3,696 estimated (3,647 before X4 and X5); the guard was 3,700, a
regression guard and not a budget, raised as the read tools arrive.  T19 says what
``key_column`` names, where ``sources`` works and that ``xlsx_derive`` sorts by output
columns (its own ``sort``): ``xlsx_read`` 492 -> 501, ``xlsx_derive`` 794 -> 801, the core
3,490 estimated, about 5,165 counted; every definition 3,712, the guard 3,716.
ooxml-edit 0.36.0 (round 6, R6): find_text's ``slides`` takes slide ids as render's does
(a string with ``orInteger``; decks only): the core 3,490 -> 3,496, about 5,174 counted;
every definition 3,712 -> 3,718, the guard 3,720.
L8 (round 7): ``xlsx_read``'s ``key_column`` is ``key`` (``key_column`` and ``row`` are
aliases, never sent), and ``columns`` picks a ``t:`` target's columns: ``xlsx_read`` 501 ->
530, the core 3,496 -> 3,525 estimated, about 5,174 -> 5,217 counted; every definition
3,718 -> 3,747, the guard 3,720 -> 3,750.
"""

from __future__ import annotations

import json
import math

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems
from ooxml_edit.tools.schema import is_path_name, subset_problems, walk_properties

from xlsx_agent.tools import CORE_TOOLS, FORMAT, GROUPS, TOOLS

CHARS_PER_TOKEN = 3.5
COUNTED_PER_ESTIMATED = 1.48
CORE_COUNTED_BUDGET = 6000
ALL_GUARD = 3750


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def estimate(definitions: list[dict]) -> int:
    """Characters / 3.5 of the compact JSON of each definition as sent."""
    return math.ceil(sum(len(json.dumps(d, separators=(",", ":"), ensure_ascii=False))
                         for d in definitions) / CHARS_PER_TOKEN)


def test_every_canonical_schema_is_in_the_common_strict_subset():
    for tool in TOOLS:
        assert subset_problems(tool.canonical, allow_free=not tool.strict) == [], tool.name


@pytest.mark.parametrize("options", [{}, {"defer": True}], ids=["all-loaded", "deferred"])
def test_the_anthropic_definitions_keep_claudes_strict_rules(box, options):
    definitions = box.definitions("anthropic", **options)
    assert anthropic_problems(definitions) == []
    # Claude's strict mode is planned for the writing tools (0.2); the read tools need none
    assert len([d for d in definitions if d.get("strict")]) <= 20
    if options:
        loaded = [d for d in definitions if "input_schema" in d and not d.get("defer_loading")]
        assert {d["name"] for d in loaded} == set(CORE_TOOLS)


def test_the_openai_definitions_keep_the_strict_rules(box):
    responses = box.definitions("openai-responses")
    chat = box.definitions("openai-chat")
    assert openai_problems(responses) == [] and openai_problems(chat, chat=True) == []
    namespaces = box.definitions("openai-responses", namespaces=True, defer=True)
    assert openai_problems(namespaces) == []


def test_no_parameter_is_a_path_and_blobs_are_handles():
    for tool in TOOLS:
        for name, node, _ in walk_properties(tool.canonical):
            assert not is_path_name(name.split(".")[-1].rstrip("[]")), (tool.name, name)
    blobs = [t.name for t in TOOLS for n, _, _ in walk_properties(t.canonical) if "blob" in n]
    assert blobs and all("handle" in json.dumps(t.canonical) for t in TOOLS if t.name in blobs)


def test_the_groups_are_the_roadmaps():
    groups: dict[str, set[str]] = {}
    for tool in TOOLS:
        groups.setdefault(tool.group, set()).add(tool.name)
    # the read tools are the core; edits (0.2) will be deferred groups of their own
    assert groups == {"core": {"open_document", "describe", "find_text", "xlsx_read",
                               "xlsx_profile", "xlsx_query", "xlsx_derive", "xlsx_compare",
                               "xlsx_validate_data"},
                      "shared_misc": {"close_document", "read_blob"}}


def test_the_loaded_core_is_within_budget_and_every_definition_within_its_guard(box):
    core = estimate(box.definitions("anthropic", groups="core"))
    every = estimate(box.definitions("anthropic", defer=False))
    assert core * COUNTED_PER_ESTIMATED <= CORE_COUNTED_BUDGET, core
    assert every <= ALL_GUARD, every
