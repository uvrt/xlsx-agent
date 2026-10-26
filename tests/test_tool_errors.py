"""Which error code a tool call's failure gets (review, S9).

ooxml-edit 0.66.0 types its refusals (``ooxml_edit.errors``): its dispatcher gives an
``ArgumentError`` ``invalid_arguments`` and a ``UnitError`` (a value out of its range)
``unit``, in every format, a workbook's too.  Before, every plain ``ValueError`` was
``unit`` -- removing a diagram's last node among them, in a deck and a Word document; a
workbook has no diagram tool, so here it is the same refusal from the sheets ooxml-edit
writes.  The library's own errors keep the codes ``ERRORS`` gives them.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.errors", reason="needs ooxml-edit 0.66 (typed refusals)")

from ooxml_edit.errors import ArgumentError, UnitError  # noqa: E402
from ooxml_edit.sheets import Cell, column_letters  # noqa: E402
from ooxml_edit.tools import Toolbox  # noqa: E402

from xlsx_agent import (AddressError, DeclarationError, EditError, NameNotFound,  # noqa: E402
                        OpenError, PolicyError, QueryError, RegionError)
from xlsx_agent.schema import SchemaError  # noqa: E402
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from xlsx_agent.tools.docformat import ERRORS  # noqa: E402


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _raised(fn, *args):
    try:
        fn(*args)
    except ValueError as error:
        return error
    raise AssertionError(f"{fn.__name__}{args} raised nothing")


def test_ooxml_edits_refusal_is_invalid_arguments_and_its_unit_error_unit(box):
    refused = _raised(Cell.parse, "A0B")                   # not a cell reference
    assert isinstance(refused, ArgumentError) and not isinstance(refused, UnitError)
    assert box.map_exception(refused).code == "invalid_arguments"
    outside = _raised(column_letters, 0)                   # a column outside the sheet
    assert isinstance(outside, UnitError)
    assert box.map_exception(outside).code == "unit"
    assert box.map_exception(ArgumentError("a diagram keeps at least one node")).code \
        == "invalid_arguments"


# The library's own, as ERRORS maps them (unchanged by S9).
OWN = [
    (AddressError("not an address"), "invalid_arguments"),
    (NameNotFound("no sheet 'X'", what="sheet", name="X"), "not_found"),
    (RegionError("bad_table", "no table"), "refused"),
    (OpenError("xls"), "refused"),
    (QueryError("bad_filter", "no such test"), "invalid_arguments"),
    (SchemaError("columns", "no such column"), "invalid_arguments"),
    (DeclarationError("bad", "tables", "not a table"), "refused"),
    (PolicyError("writes", "not a policy"), "refused"),
    (EditError("rolled_back"), "refused"),
]


@pytest.mark.parametrize("error, code", OWN, ids=lambda item: type(item).__name__
                         if isinstance(item, BaseException) else item)
def test_the_librarys_own_errors_keep_their_codes(box, error, code):
    assert box.map_exception(error).code == code


def test_every_mapped_error_is_tested_and_none_is_an_argument_error():
    assert {type(error) for error, _ in OWN} == set(ERRORS)
    assert not any(issubclass(cls, ArgumentError) for cls in ERRORS)


def test_pythons_own_errors_keep_the_fallback(box):
    assert box.map_exception(ValueError("invalid literal")).code == "unit"
    assert box.map_exception(TypeError("unsupported operand")).code == "internal"
