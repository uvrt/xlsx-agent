"""The xlsx tool definitions against each provider's rules, and the token budget.

Offline: every canonical schema is in the common strict subset (``ooxml_edit.tools`` refuses
a tool that is not, when it is made); the Anthropic, OpenAI Responses and Chat Completions
adapters' output passes each provider's documented strict rules; no parameter is a path;
the groups are as the roadmap lists them.

**The budget** (ROADMAP.md, "Tools"; the siblings' per-format budget): what a request
loads -- the core -- is at most 6,500 tokens as Anthropic's count-tokens counts them.
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
528, the core 3,496 -> 3,523 estimated, about 5,174 -> 5,214 counted; every definition
3,718 -> 3,746, the guard 3,720 -> 3,750.
X2 of the second production lane (E1, ooxml-edit 0.51.0): the shared ``undo`` (158),
``batch`` (202) and ``save_document`` (186) join the core, as the roadmap's 0.2 table has
them: the core 3,523 -> 4,069 estimated, about 5,214 -> 6,022 counted by the proxy, under
6,500; every definition 3,746 -> 4,291, the guard 3,750 -> 4,300.
X3b (ooxml-edit 0.54.0): ``xlsx_read``'s ``markdown`` (text as Markdown) and
``xlsx_set_cells``' ``markdown`` value: ``xlsx_read`` 528 -> 556, the core 4,069 -> 4,096
estimated, about 6,022 -> 6,062 counted by the proxy; ``xlsx_set_cells`` 347 -> 378;
every definition 4,627 -> 4,686, the guard 4,650 -> 4,690.
X4 (guards): ``xlsx_set_cells``' address by key, each cell's ``expect`` (a type) and
``expect_value`` (the value as read), ``new_rows`` and ``revision``: ``xlsx_set_cells`` 378
-> 533 (deferred: the core unchanged, 4,096); every definition 4,686 -> 4,840, the guard
4,690 -> 4,850.  (A first form with ``expect`` an object of typed fields, 553, was refused
online as too complex for strict mode: see the strict complexity proxy below.)
X5 (integrity compare, validation inspection): ``xlsx_compare``'s ``cells`` and
``formula_text`` (330 -> 395) and ``xlsx_read``'s ``validations`` (556 -> 593): the core
4,096 -> 4,198 estimated, about 6,062 -> 6,213 counted by the proxy, under 6,500; every
definition 4,840 -> 4,942, the guard 4,850 -> 4,950.  No strict tool changed.
X6 (the contract's gaps): ``xlsx_set_cells``' ``new_rows`` (the next free row) gives way to
each cell's ``row`` (the capacity row the application allocates to a new key): 533 -> 529,
the strict tool's optional parameters still 13, its free strings still 8; ``xlsx_compare``'s
``cells`` names the structure it compares (395 -> 409): the core 4,198 -> 4,212 estimated,
about 6,213 -> 6,234 counted by the proxy, under 6,500; every definition 4,942 -> 4,953,
the guard 4,950 -> 4,960.
Round 10, item 4 (ooxml-edit 0.55.0): ``xlsx_query`` takes a dataset (O5: ``doc`` optional,
``table`` and the description name ``d:``), 457 -> 474, and the shared ``undo``'s ``scope``
names a sheet as well as a slide, 158 -> 159: the core 4,212 -> 4,230 estimated, about
6,234 -> 6,261 counted by the proxy, under 6,500; every definition 4,953 -> 4,971, the
guard 4,960 -> 4,975.  The strict set is unchanged: neither tool is strict.
Round 16, F11 (ooxml-edit 0.60.0): the shared ``save_document``'s ``clean`` (documents:
what Word never draws taken out of the file saved; a workbook's handler takes none, so
ooxml-edit refuses ``clean: true`` for it), +42: the core 4,230 -> 4,272 estimated, about
6,261 -> 6,323 counted by the proxy, under 6,500; every definition 4,971 -> 5,013, the
guard 4,975 -> 5,015.  The strict set's members are unchanged: save_document is strict on
OpenAI only, where ``clean`` is one more nullable boolean.
Round 23, V3 (ooxml-edit 0.70.0, Tool.kind_fields): the shared ``find_text`` is sent
without a deck's ``slides`` and the ``stories`` of decks and documents, which a workbook's
handler refuses, and its ``range`` described for a workbook alone; ``save_document`` without
a document's ``clean`` and ``undo``'s ``scope`` described for a sheet: the core 4,272 ->
4,122 estimated, about 6,323 -> 6,101 counted by the proxy; every definition 5,013 -> 4,862,
the guard 5,015 -> 4,865.  None of them is strict on Claude: the strict set is unchanged.
Round 25, X0 and X1: ``xlsx_read`` reads the table rows a sheet's cell or range lies in
(``target`` says so, with an example of a cited cell; ``columns`` picks their columns;
``cells_only``), 593 -> 688; ``xlsx_derive``'s ``join`` (``from``, ``on``, ``columns``, with
an example), 801 -> 1,045: the core 4,122 -> 4,460 estimated, about 6,101 -> 6,601 counted by
the proxy; counted online, 6,334 (claude-sonnet-5-5, 2026-10-30; 5,894 before).  The user lets the core grow past 6,500 for a good tool (2026-10-30), so the core's
offline check is a regression guard from here (:data:`CORE_GUARD`), and the online count is
the main session's.  Every definition 4,862 -> 5,201, the guard 4,865 -> 5,205.  Neither tool
is strict on Claude: the strict set is unchanged (``xlsx_set_cells``, 13 optional).
Round 27, X6: ``xlsx_derive``'s ``pivot_values`` (the pivot's columns, listed in order; what a
column without rows holds; the values not listed named in ``pivot_unlisted``), 1,045 ->
1,110: the core 4,460 -> 4,525 estimated, about 6,601 -> 6,697 counted by the proxy (online,
about 6,334 -> 6,430 at the same ratio; the main session counts it); every definition
5,201 -> 5,266; the guards 4,465 -> 4,530 and 5,205 -> 5,270.  ``xlsx_derive`` is not
strict (12 optional parameters): the strict set is unchanged.
Round 28, X7 and X8: ``xlsx_read``'s ``columns`` names a ``d:`` target too, 688 -> 690;
``xlsx_derive``'s ``pivot_values`` says ``""`` is the blank cells, 1,110 -> 1,116: the core
4,525 -> 4,533 estimated (online 6,418 on 2026-11-03, about 6,430 at the same ratio; the
main session counts it); every definition 5,266 -> 5,274; the guards 4,530 -> 4,535 and
5,270 -> 5,275.  Neither tool is strict: the strict set is unchanged.
Round 31, R7: ``xlsx_trace`` (rows traced to the rows they rest on) in a deferred group of
its own, ``xlsx_trace`` (242 estimated with its group's line in the tool-search tool, about
358 counted by the proxy): the core unchanged, 4,533; every definition 5,274 -> 5,505, the
guard 5,275 -> 5,510.  It reads, so it is not strict: the strict set is unchanged.
Round 32, R15: ooxml-edit's ``read_batch`` (many reads in one call, each with its own budget)
in the core, where a builder reading record after record reaches for it, 184 estimated: the
core 4,533 -> 4,717, every definition 5,505 -> 5,688; the guards 4,535 -> 4,720 and 5,510 ->
5,690.  Its ops are free-form, so it is not strict: the strict set is unchanged.  Counted
online, 6,431 -> 6,678 (claude-sonnet-5-5, 2026-11-05); the user keeps it in the core and
raises the online budget 6,500 -> 6,800 (``test_tools_online.py``).
Round 35, R22: ``xlsx_derive``'s ``fill`` (a column's blanks given another column's value of
the same row, after the join), the note that a join alone gives the rows moved to the tool's
description (now for a fill too): ``xlsx_derive`` 1,116 -> 1,224, the core 4,717 -> 4,824
estimated, counted online 6,838 (claude-sonnet-5-5, 2026-11-05), past the 6,800 budget; so
xlsx_derive's texts were shortened, meaning kept (``doc`` worded as xlsx_read's and
xlsx_query's, ``join``'s and ``pivot_values``' sentences tightened): 1,224 -> 1,182, the
core 4,824 -> 4,783 (about 6,780 online at the counted ratio, 1.4175; the main session
counts it), every definition 5,688 -> 5,754; the guards 4,720 -> 4,785 and 5,690 -> 5,755.
``xlsx_derive`` is not strict (13 optional parameters): the strict set is unchanged.
Round 35: user (2026-11-06): the budget is guidance, clear tools first; 6,800 -> 7,000;
join.on / pivot_values / coerce wording restored (``join.on`` names the outer ``from``,
``join.from`` says it is the table to join, ``pivot_values`` reads as sentences again,
``coerce`` says "Opt in:"): ``xlsx_derive`` 1,182 -> 1,228, the core 4,783 -> 4,829
estimated (about 6,845 online at the counted ratio, 1.4175; the main session counts it),
every definition 5,754 -> 5,800; the guards 4,785 -> 4,830 and 5,755 -> 5,800.  Then
``fill`` says its ``from`` is the fill's own column of the same row, ``join`` names "the
outer from", ``join.columns`` says they are this join's table's, and ``xlsx_compare``'s and
``xlsx_validate_data``'s ``cursor`` says "the same call" (neither has a target):
``xlsx_derive`` 1,228 -> 1,245, the core 4,829 -> 4,844 (about 6,866 online), every
definition 5,800 -> 5,816; the guards 4,830 -> 4,845 and 5,800 -> 5,820.  ``xlsx_query``'s
``cursor`` says "the same call" too (its parameter is ``table``): every definition 5,816 ->
5,815, the core unchanged.
Round 40, R43: the filters' ``contains_any`` (text holding any of ``values``) and ``loose``
(text compared ignoring case and white space), in ``xlsx_query``'s and ``xlsx_derive``'s
``where``: ``xlsx_query`` 473 -> 513, ``xlsx_derive`` 1,245 -> 1,286, the core 4,844 -> 4,924
estimated (about 6,970 online at the counted ratio, 6,854 / 4,844; the main session counts it), every
definition 5,815 -> 5,896; the guards 4,845 -> 4,925 and 5,820 -> 5,900.  Neither tool is
strict: the strict set is unchanged.
Round 41: ``text_chars`` (the characters per text in a page) on ``xlsx_read`` and
``xlsx_query``: ``xlsx_read`` 690 -> 713, ``xlsx_query`` 513 -> 537, the core 4,924 -> 4,972
estimated, every definition 5,896 -> 5,943; counted online 7,024 (the main session,
2026-11-08), past the 7,000 budget.  So ``xlsx_query``'s description drops "A filter on a
column of another type is refused, naming the type." (the refusal says so itself):
``xlsx_query`` 537 -> 518, the core 4,972 -> 4,953 (about 6,997 online at the counted ratio;
the main session counts it again), every definition 5,943 -> 5,924; the guards 4,925 ->
4,955 and 5,900 -> 5,925.  Neither tool is strict: the strict set is unchanged.
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
#: Round 35: user (2026-11-06): the budget is guidance, clear tools first; 6,800 -> 7,000;
#: join.on / pivot_values / coerce wording restored.
CORE_COUNTED_BUDGET = 7000
#: The core's estimate, a regression guard (round 25: past the proxy's budget, by choice).
#: Round 35: the restored wording, 4,783 -> 4,829 and 5,754 -> 5,800; then fill's, join's,
#: join.columns' and two cursors' wording, -> 4,844 and 5,816.
#: Round 40, R43: contains_any and loose in the filters, -> 4,924 and 5,896.
#: Round 41: text_chars on xlsx_read and xlsx_query, and xlsx_query's sentence on refused
#: filters dropped, -> 4,953 and 5,924.
CORE_GUARD = 4955
ALL_GUARD = 5925


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
    # the read tools, save_document, undo and batch are the core; edits (0.2) are
    # deferred groups of their own: xlsx_edit from E3
    assert groups == {"core": {"open_document", "describe", "find_text", "xlsx_read",
                               "xlsx_profile", "xlsx_query", "xlsx_derive", "xlsx_compare",
                               "xlsx_validate_data", "save_document", "undo", "batch",
                               "read_batch"},
                      "shared_misc": {"close_document", "read_blob"},
                      "xlsx_edit": {"xlsx_set_cells"},
                      "xlsx_trace": {"xlsx_trace"}}


def test_find_text_is_sent_without_a_decks_or_documents_parameters(box):
    """Round 23, V3: no slides and no stories, which the handler refuses as before."""
    sent = next(d for d in box.definitions("anthropic", defer=False) if d["name"] == "find_text")
    properties = sent["input_schema"]["properties"]
    assert set(properties) == {"doc", "text", "regex", "range", "cursor"}
    assert properties["range"]["description"] == "Only within this sheet, range or table."
    sent = {d["name"]: d["input_schema"]["properties"] for d in box.definitions("anthropic", defer=False)}
    assert "clean" not in sent["save_document"]
    assert sent["undo"]["scope"]["description"] == (
        "A sheet's (Sales) changes only. Default: the latest, anywhere.")


def test_the_loaded_core_is_within_budget_and_every_definition_within_its_guard(box):
    core = estimate(box.definitions("anthropic", groups="core"))
    every = estimate(box.definitions("anthropic", defer=False))
    assert core <= CORE_GUARD, core
    assert every <= ALL_GUARD, every


# -- Claude's strict grammar: the complexity proxy (X4) -------------------------------------------

#: What the strict tools may reach, summed over every one sent strict.  Claude documents 24
#: optional parameters and an undocumented "Schema is too complex" (HTTP 400) beyond them:
#: X4's first ``xlsx_set_cells`` (17 optional, 9 free strings, an optional object inside
#: the array's items: three object levels) was refused online on Sonnet 5.5, 2026-10-20,
#: while X3b's (9 optional, 7 free strings, two levels) was accepted.  Assumed limits, from
#: those two points: no object below an array item's (two object levels at most), and no
#: more optional parameters or free-text strings than the flat X4 schema has (13 and 8).
#: Raise them only after the online suite (tests/test_tools_online.py) accepts more.
STRICT_OBJECT_DEPTH = 2
STRICT_OPTIONAL = 13
STRICT_FREE_STRINGS = 8


def _object_depth(node: dict, depth: int = 1) -> int:
    children = list((node.get("properties") or {}).values())
    if "items" in node:
        children.append(node["items"])
    return max([depth] + [_object_depth(child, depth + (child.get("type") == "object"))
                          for child in children])


def _optional(node: dict) -> int:
    properties = node.get("properties") or {}
    required = set(node.get("required", ()))
    count = sum(1 for name in properties if name not in required)
    count += sum(_optional(child) for child in properties.values())
    return count + (_optional(node["items"]) if "items" in node else 0)


def test_the_strict_tools_stay_within_the_complexity_claude_compiled(box):
    from ooxml_edit.tools.adapters import free_strings

    strict = [d for d in box.definitions("anthropic", defer=False) if d.get("strict")]
    assert [d["name"] for d in strict] == ["xlsx_set_cells"]
    for definition in strict:
        assert _object_depth(definition["input_schema"]) <= STRICT_OBJECT_DEPTH, \
            definition["name"]
    assert sum(_optional(d["input_schema"]) for d in strict) <= STRICT_OPTIONAL
    assert sum(free_strings(d["input_schema"]) for d in strict) <= STRICT_FREE_STRINGS
