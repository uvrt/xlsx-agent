"""The workbook tools by role (ooxml-edit 0.45.0, round 7, item 9).

Every xlsx tool but the shared ``undo`` reads (E1 adds ``undo``, ``batch`` and
``save_document``; the writing tools come with E3): all of them are ``review``'s, or any
role's, and a reviewer, ``["review"]`` or ``["review", "resolve"]``, is sent the default's
definitions without ``undo``, which is ``edit``'s (no strict tool), with ``accept`` too.  A
role set without ``review`` -- ``["resolve"]``, ``["accept"]`` -- is sent only the
session's ``open_document``, ``close_document``, ``read_blob``, ``batch`` and
``save_document``, and ``["edit"]`` ``undo`` too; every read is refused.  A reviewer's
prompt has the role line without the sentence about comments: no workbook tool places one.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import ROLES, Toolbox
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

from build import WorkbookBuilder
from test_tools_schema import ALL_GUARD, COUNTED_PER_ESTIMATED, CORE_COUNTED_BUDGET, estimate


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


def test_every_workbook_tool_is_a_reviewer_s():
    assert all(tool.roles <= {"review"} or tool.role == "any" for tool in TOOLS
               if tool.name != "undo")
    assert next(t for t in TOOLS if t.name == "undo").roles == {"edit"}


@pytest.mark.parametrize("provider", ["anthropic", "openai-responses", "openai-chat"])
@pytest.mark.parametrize("roles", [["review"], ["review", "resolve"], ["review", "accept"],
                                   ROLES], ids=str)
def test_a_reviewer_is_sent_the_default_s_definitions(box, provider, roles):
    for groups in (None, "core"):
        default = box.definitions(provider, groups=groups)
        if "edit" not in roles:
            default = [d for d in default if _name(d) != "undo"]
        assert box.definitions(provider, groups=groups, roles=roles) == default


def _name(definition: dict) -> str:
    return definition.get("name") or definition.get("function", {}).get("name", "")


def test_a_reviewer_s_core_is_the_core_and_within_the_budget(box):
    core = estimate(box.definitions("anthropic", groups="core", roles=["review"]))
    everything = estimate(box.definitions("anthropic", defer=False, roles=["review"]))
    assert core * COUNTED_PER_ESTIMATED <= CORE_COUNTED_BUDGET, core
    assert everything <= ALL_GUARD, everything


@pytest.mark.parametrize("roles", [["edit"], ["resolve"], ["accept"]], ids="+".join)
def test_a_set_without_review_reads_nothing(box, roles):
    sent = {d["name"] for d in box.definitions("anthropic", defer=False, roles=roles)}
    assert sent == {"open_document", "close_document", "read_blob", "batch",
                    "save_document"} | ({"undo"} if "edit" in roles else set())
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
