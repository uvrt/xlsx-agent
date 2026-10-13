"""Excel tools for an agent tool layer (``ooxml_edit.tools``): definitions and handlers.

An application puts these in a :class:`ooxml_edit.tools.Toolbox` with the xlsx
:data:`FORMAT` (and, beside them, another library's tools for its documents)::

    from ooxml_edit.tools import Toolbox
    from xlsx_agent.tools import FORMAT, GROUPS, TOOLS

    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=my_clock)
    b1 = session.add_blob(xlsx_bytes, name="inventory.xlsx")
    tools = toolbox.definitions("anthropic", groups="core")
    openai_tools = toolbox.definitions("openai-responses")   # or "openai-chat"
    result = toolbox.dispatch(session, "open_document", {"blob": b1})      # d1
    result = toolbox.dispatch(session, "describe", {"doc": "d1"})

The locale (how text that looks like a number or a date is read) is the application's
setting, never the model's: ``format(locale=Locale(decimal=",", thousands="."))`` in place
of :data:`FORMAT`.

Everything is in memory: workbooks and inputs are bytes under handles.  So far (R6) the
tools are the shared ``open_document``, ``close_document`` and ``read_blob``, and the
workbook's ``describe`` (:mod:`.shared`); the read and derive tools follow (ROADMAP.md,
"Milestones").  The library's address errors become tool errors with their valid
options (:mod:`.docformat`).
"""

from __future__ import annotations

from ooxml_edit.tools import CORE, DocumentFormat, shared as _shared

from ..read.types import Locale
from .docformat import (ERRORS, KIND, PROMPT, REGION_CODES, STRICT_FIRST, FORMAT, detect,
                        tool_error, workbook_format)
from .shared import TOOLS as SHARED_TOOLS

GROUPS = [*_shared.GROUPS]

#: Every tool a workbook has, the shared ones included, in the order they are offered.
TOOLS = [*SHARED_TOOLS]

CORE_TOOLS = [t.name for t in TOOLS if t.group == CORE]


def format(*, locale: Locale | None = None) -> DocumentFormat:  # noqa: A001 (ROADMAP's name)
    """The xlsx format with the application's ``locale`` (ROADMAP.md, "Values and types":
    ``xlsx_agent.tools.format(locale=...)``); :data:`FORMAT` is ``format()``."""
    return workbook_format(locale=locale)


__all__ = ["CORE_TOOLS", "ERRORS", "FORMAT", "GROUPS", "KIND", "PROMPT", "REGION_CODES",
           "STRICT_FIRST", "TOOLS", "detect", "format", "tool_error"]
