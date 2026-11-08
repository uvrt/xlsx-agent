"""The workbook tools by role (ooxml-edit 0.45.0, round 7, item 9).

Every xlsx tool but the shared ``undo`` and E3's ``xlsx_set_cells`` reads (E1 adds
``undo``, ``batch`` and ``save_document``): all of them are ``review``'s, or any role's,
and a reviewer, ``["review"]`` or ``["review", "resolve"]``, is sent the default's
definitions without ``undo`` and ``xlsx_set_cells``, which are ``edit``'s, with ``accept``
too.  A role set without ``review`` -- ``["resolve"]``, ``["accept"]`` -- is sent only the
session's ``open_document``, ``close_document``, ``read_blob``, ``batch`` and
``save_document``, and ``["edit"]`` ``undo`` and ``xlsx_set_cells`` too; every read is
refused.  A reviewer's
prompt has the role line without the sentence about comments: no workbook tool places one.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import ROLES, Toolbox
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from build import WorkbookBuilder
from test_tools_schema import ALL_GUARD, CORE_GUARD, estimate


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _book() -> bytes:
    book = WorkbookBuilder()
    sheet = book.sheet("Stock")
    sheet.rows("A2", [["A-1", 1], ["A-2", 2]])
    sheet.table("Lines", "A1:B3", ["SKU", "Qty"])
    return book.to_bytes()


#: The workbook tools that are the edit role's, and the group of its own they make
#: (an OpenAI namespace a reviewer is not sent).
EDITING = {"undo", "xlsx_set_cells", "xlsx_format"}
EDITING_GROUPS = {"xlsx_edit", "xlsx_format"}


def test_every_workbook_tool_is_a_reviewer_s():
    assert all(tool.roles <= {"review"} or tool.role == "any" for tool in TOOLS
               if tool.name not in EDITING)
    assert all(t.roles == {"edit"} for t in TOOLS if t.name in EDITING)


@pytest.mark.parametrize("provider", ["anthropic", "openai-responses", "openai-chat"])
@pytest.mark.parametrize("roles", [["review"], ["review", "resolve"], ["review", "accept"],
                                   ROLES], ids=str)
def test_a_reviewer_is_sent_the_default_s_definitions(box, provider, roles):
    for groups in (None, "core"):
        default = box.definitions(provider, groups=groups)
        if "edit" not in roles:
            default = [d for d in default if _name(d) not in EDITING | EDITING_GROUPS]
        assert box.definitions(provider, groups=groups, roles=roles) == default


def _name(definition: dict) -> str:
    return definition.get("name") or definition.get("function", {}).get("name", "")


def test_a_reviewer_s_core_is_the_core_and_within_the_budget(box):
    core = estimate(box.definitions("anthropic", groups="core", roles=["review"]))
    everything = estimate(box.definitions("anthropic", defer=False, roles=["review"]))
    assert core <= CORE_GUARD, core
    assert everything <= ALL_GUARD, everything


@pytest.mark.parametrize("roles", [["edit"], ["resolve"], ["accept"]], ids="+".join)
def test_a_set_without_review_reads_nothing(box, roles):
    sent = {d["name"] for d in box.definitions("anthropic", defer=False, roles=roles)}
    assert sent == {"open_document", "close_document", "read_blob", "batch",
                    "save_document"} | (EDITING if "edit" in roles else set())
    assert anthropic_problems(box.definitions("anthropic", roles=roles)) == []
    assert openai_problems(box.definitions("openai-responses", roles=roles)) == []
    session = box.session(roles=roles)
    session.open(_book(), "stock.xlsx")
    refused = box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Lines"})
    assert refused.error.code == "refused" and refused.error.details["role"] == "review"
    assert "needs the review role" in refused.error.message


def test_a_reviewer_reads_and_derives(box):
    session = box.session(roles=["review"])
    session.open(_book(), "stock.xlsx")
    assert box.dispatch(session, "describe", {"doc": "d1"}).ok
    assert box.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Lines"}).ok


def test_a_reviewer_s_prompt_has_the_role_line_without_comments(box):
    prompt = box.system_prompt(roles=["review"])
    assert "Role: review. Do not change the document" in prompt
    assert "comment" not in prompt.split("Role: review.")[1]
    assert box.system_prompt() == box.system_prompt(roles=ROLES)
    assert "Role:" not in box.system_prompt()


def test_two_agents_views_of_one_session_share_its_workbook_and_datasets(box):
    """ooxml-edit 0.46.0 (round 8, M1): agents on one session, each through a view."""
    session = box.session()                                  # every role
    session.open(_book(), "stock.xlsx")
    analyst = session.view(roles=["review"], context={"agent": "analyst-1"})
    writer = session.view(roles=["edit"], context={"agent": "writer-1"})
    derived = box.dispatch(analyst, "xlsx_derive", {
        "doc": "d1", "from": "t:Lines", "measures": [{"fn": "sum", "column": "Qty"}],
        "id": "stock"})
    assert derived.ok and session.dataset_ids() == ["d:stock"]
    # The dataset is the session's: every view sees it.
    assert writer.dataset("d:stock") is analyst.dataset("d:stock")
    refused = box.dispatch(writer, "xlsx_read", {"doc": "d1", "target": "t:Lines"})
    assert refused.error.code == "refused" and refused.error.details == {
        "role": "review", "roles": ["edit"]}
    assert [record.context for record in session.log] == [{"agent": "analyst-1"},
                                                          {"agent": "writer-1"}]
    with pytest.raises(ValueError, match="never widens"):
        writer.view(roles=["review"])


def test_a_reviewer_reads_many_records_in_one_read_batch(box):
    """ooxml-edit 0.78.0 (round 32, R15): ``read_batch`` is a reader's, in the core: records
    by key, a query and the catalogue in one call, each answered as alone; a change in it
    is refused with the way out, the others answered."""
    assert "read_batch" in {d["name"] for d in box.definitions("anthropic", groups="core",
                                                               roles=["review"])}
    assert "read_batch makes many reads" in box.system_prompt(roles=["review"])
    session = box.session()
    session.open(_book(), "stock.xlsx")
    reader = session.view(roles=["review"], context={"agent": "reader-1"})
    alone = box.dispatch(reader, "xlsx_read", {"doc": "d1", "target": "t:Lines/A-2",
                                               "key": "SKU"})
    result = box.dispatch(reader, "read_batch", {"ops": [
        {"tool": "xlsx_read", "arguments": {"doc": "d1", "target": "t:Lines/A-1",
                                            "key": "SKU"}},
        {"tool": "xlsx_read", "arguments": {"doc": "d1", "target": "t:Lines/A-2",
                                            "key": "SKU"}},
        {"tool": "xlsx_query", "arguments": {"doc": "d1", "table": "t:Lines"}},
        {"tool": "describe", "arguments": {"doc": "d1"}},
        {"tool": "xlsx_set_cells", "arguments": {"doc": "d1",
                                                 "cells": [{"address": "Stock!B2", "number": 5}]}}]})
    assert result.ok
    answers = result.data["ops"]
    assert [answer.get("error") is None for answer in answers] == [True] * 4 + [False]
    assert answers[1]["data"] == alone.data
    assert answers[4]["error"]["details"]["role"] == "edit"
    assert session.entry("d1").version == 0


def test_a_builder_s_change_in_a_read_batch_is_sent_to_batch(box):
    session = box.session()
    session.open(_book(), "stock.xlsx")
    result = box.dispatch(session, "read_batch", {"ops": [
        {"tool": "xlsx_set_cells", "arguments": {"doc": "d1",
                                                 "cells": [{"address": "Stock!B2", "number": 5}]}},
        {"tool": "xlsx_read", "arguments": {"doc": "d1", "target": "Stock!B2"}}]})
    refused, read = result.data["ops"]
    assert refused["error"]["code"] == "refused"
    assert refused["error"]["valid_options"] == ["batch", "xlsx_set_cells"]
    assert "error" not in read and session.entry("d1").version == 0
