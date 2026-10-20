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
    result = toolbox.dispatch(session, "xlsx_read", {"doc": "d1", "target": "t:Products"})
    result = toolbox.dispatch(session, "xlsx_query", {
        "doc": "d1", "table": "t:Products", "columns": ["SKU", "Price"],
        "where": [{"column": "Category", "op": "eq", "value": "Hardware"}],
        "sort": [{"column": "Price", "desc": True}]})
    result = toolbox.dispatch(session, "find_text", {"doc": "d1", "text": "Hardware",
                                                     "range": "Stock!"})
    result = toolbox.dispatch(session, "xlsx_derive", {
        "doc": "d1", "from": "t:Products", "id": "by_category", "group_by": ["Category"],
        "measures": [{"fn": "mean", "column": "Price", "as": "average"}], "round": 2})
    result = toolbox.dispatch(session, "xlsx_read", {"target": "d:by_category[average]",
                                                     "sources": True})
    result = toolbox.dispatch(session, "xlsx_compare", {"a": "d1", "b": "d2",
                                                        "tables": ["t:Products"], "key": "SKU"})
    result = toolbox.dispatch(session, "xlsx_validate_data", {"doc": "d1", "schema": {
        "tables": [{"table": "t:Products", "key": ["SKU"], "columns": [
            {"name": "Price", "type": "number", "min": "0"}]}]}})

The locale (how text that looks like a number or a date is read, and the language of the
language-dependent built-in number formats) is the application's setting, never the
model's: the toolbox's or the session's, as for every format --
``Toolbox(..., locale=Locale(decimal_separator=",", thousands_separator=".",
date_order="dmy", language="nl-BE", list_separator=";"))`` with
:class:`ooxml_common.locales.Locale`, or ``toolbox.session(locale=...)`` per user; en-US
unless given.  The reader's own settings go on the format:
``format(local_dates=True)`` in place of :data:`FORMAT` reads text dates in the locale's
date order, as Excel does (:mod:`xlsx_agent.read.types`).

Tables the application knows from its own configuration -- sheets by name or glob, a
table over several sheets, the header row, the key, column ids and descriptions, columns
hidden from the model, key-value tables, whole sheets hidden and terms redacted
(:mod:`xlsx_agent.redact`) -- are declared as plain data, never by the model,
so they cost no definition tokens: ``format(tables=[...])`` for every workbook, or
``declare(session, [...])`` for one session's (:mod:`xlsx_agent.declare`)::

    toolbox = Toolbox(TOOLS, formats=[format(tables={"tables": [
        {"name": "checks", "sheets": "Area *", "header_row": 3, "key": "Ref",
         "fields": [{"header": "Reviewer notes", "hidden": True}]}]})], groups=GROUPS)

Everything is in memory: workbooks and inputs are bytes under handles.  In 0.1.0 the
tools are the shared ``open_document``, ``close_document`` and ``read_blob``, the
workbook's ``describe`` and ``find_text`` (:mod:`.shared`), ``xlsx_read`` and ``xlsx_profile``
(:mod:`.read`), ``xlsx_query`` (:mod:`.query`), and ``xlsx_derive`` (:mod:`.derive`), which
keeps its figures in the session as a dataset, ``d:<id>``: ``xlsx_read`` reads it and
``xlsx_derive`` derives from it, after its workbook is closed too, and a dataset whose
workbook has changed since is marked stale, never refreshed (ooxml-edit's datasets, O2);
``xlsx_compare`` (:mod:`.compare`), two open workbooks compared by key; and
``xlsx_validate_data`` (:mod:`.validate`), a workbook checked against a declared schema.
From 0.2 (E1, E3) the shared ``save_document``, ``undo`` and ``batch`` are in the core, and
the deferred ``xlsx_edit`` group holds ``xlsx_set_cells`` (:mod:`.edit`): cells written
typed and guarded, in one step (the ``edit`` role's).  Where the model may write is the
application's (X4): ``policy(session, {"cells": [...], "tables": [...]})`` before the
session opens a workbook (:mod:`xlsx_agent.edit.policy`; ``{"cells": []}`` lets nothing be
written), ``narrow(session, ...)`` to narrow it later; ``describe`` shows it
(``writable``).  Without one, the session writes as its ``edit`` role allows; a protected
sheet is never written.  Every formula value read carries
its cache status (``cache``: ``cached_unverified``, ``invalidated`` once the workbook was
written, ``absent``).  The library's address and query errors become tool errors with
their valid options (:mod:`.docformat`).
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import CORE, DocumentFormat, ToolGroup, shared as _shared

from .docformat import (ERRORS, KIND, PROMPT, QUERY_CODES, REGION_CODES, STRICT_FIRST, FORMAT,
                        declare, detect, narrow, policy, tool_error, workbook_format)
from .compare import TOOLS as COMPARE_TOOLS
from .derive import TOOLS as DERIVE_TOOLS
from .edit import TOOLS as EDIT_TOOLS
from .query import TOOLS as QUERY_TOOLS
from .read import TOOLS as READ_TOOLS
from .shared import TOOLS as SHARED_TOOLS
from .validate import TOOLS as VALIDATE_TOOLS

GROUPS = [*_shared.GROUPS,
          ToolGroup("xlsx_edit", "Excel: write cells (text, numbers, dates, formulas, blanks).")]

#: Every tool a workbook has, the shared ones included, in the order they are offered.
TOOLS = [*SHARED_TOOLS, *READ_TOOLS, *QUERY_TOOLS, *DERIVE_TOOLS, *COMPARE_TOOLS,
         *VALIDATE_TOOLS, *EDIT_TOOLS]

CORE_TOOLS = [t.name for t in TOOLS if t.group == CORE]


def format(*, grouping: bool = True, local_dates: bool = False,  # noqa: A001
           tables: Any = None) -> DocumentFormat:
    """The xlsx format with the application's reader settings (ROADMAP.md, "Values and
    types"; :meth:`Workbook.open <xlsx_agent.Workbook.open>`) and its declared ``tables``
    (:mod:`xlsx_agent.declare`), declared on every workbook it opens; its workbooks are
    read in the session's locale.  :data:`FORMAT` is ``format()``."""
    return workbook_format(grouping=grouping, local_dates=local_dates, tables=tables)


__all__ = ["CORE_TOOLS", "ERRORS", "FORMAT", "GROUPS", "KIND", "PROMPT", "QUERY_CODES",
           "REGION_CODES",
           "STRICT_FIRST", "TOOLS", "declare", "detect", "format", "narrow", "policy",
           "tool_error"]
