"""The xlsx handlers of the shared tools (``ooxml_edit.tools.shared``): one definition each
for every format, dispatched by the document's kind.

``open_document``, ``close_document`` and ``read_blob`` are the session's own
(``shared.SESSION_TOOLS``): they work for a workbook as for any document, through the
xlsx :data:`~xlsx_agent.tools.FORMAT` (its ``detect``, ``open``, ``problems`` and
``summary``).  This module adds the workbook's ``describe``: the catalogue
(:mod:`xlsx_agent.catalogue`).  ``find_text`` (R9), ``save_document``, ``undo`` and
``batch`` (E1) and ``new_document`` (E9) come with their milestones.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, shared

from .docformat import KIND, handler

#: The session tools a workbook uses so far; ``undo`` and ``batch`` wait for edits (E1).
SESSION_NAMES = ("open_document", "close_document", "read_blob")


# -- S16 describe ------------------------------------------------------------------------------


@shared.handler("describe", kind=KIND)
@handler
def describe(call: Any, doc: str) -> Result:
    """The shared ``describe``'s xlsx handler: the workbook's catalogue, within its token
    target (:data:`~xlsx_agent.catalogue.DEFAULT_TOKENS`)."""
    entry = call.entry
    data = {"name": entry.name, **call.document.catalogue()}
    return Result(summary=f"Described {entry.name}", data=data)


TOOLS = [*[tool for tool in shared.SESSION_TOOLS if tool.name in SESSION_NAMES], describe]
