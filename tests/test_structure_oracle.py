"""Inserting, deleting and moving rows and columns against Excel itself (round 44).

Opt in with ``pytest -m oracle`` (tests/oracle.py says how Excel is driven).  For every
case in ``golden/structure`` (see ``test_structure.py``):

* Excel is asked again to do the operation itself (``Session.operate``: its own insert,
  delete, or Cut and "Insert Cut Cells", through ``tools/excel_structure.applescript``), and
  what it saves must still be the recorded facts -- or it must still refuse;
* xlsx-agent's result of the same operation must open in Excel without a repair prompt, and
  Excel's re-save of it must hold the same facts as Excel's own operation did.

One difference is known and left out: a chart series reference that became ``#REF!`` is
written by both alike, but Excel drops such a series reference when it opens a file
(cases ``set2-E2``, ``E4``, ``E8``).  Without Excel, or with Excel in use or its lock held,
every test here skips.
"""

from __future__ import annotations

import sys

import pytest

import oracle
import structure_book
from structure_facts import differences, facts, measured
from test_structure import cases, operate

#: Cases whose chart series turn #REF!: compared without the charts (see the module).
REF_SERIES = {"set2-E2", "set2-E4", "set2-E8"}


@pytest.fixture(scope="module")
def excel():
    if sys.platform != "darwin" or not oracle.available():
        pytest.skip("the Excel oracle needs Microsoft Excel for Mac")
    try:
        with oracle.session() as current:
            yield current
    except oracle.ExcelBusy as busy:
        pytest.skip(str(busy))


@pytest.mark.oracle
@pytest.mark.parametrize("book, case", list(cases()))
def test_excel_agrees(excel, book, case, request):
    name = request.node.callspec.id
    source = getattr(structure_book, book)(**case["variant"]).to_bytes()
    theirs = excel.operate(source, case["steps"], name=f"structure-{name}")
    if case["excel"] == "refused":
        if theirs.ok:      # some refusals leave the file as it was, saved
            base = excel.resave(source, name=f"structure-{name}-base")
            assert not differences(measured(facts(theirs.saved.read_bytes())),
                                   measured(facts(base.saved.read_bytes())))
        else:
            assert theirs.outcome in ("error", "refused"), theirs
        return
    assert theirs.ok, theirs
    recorded = case["facts"]
    assert differences(measured(facts(theirs.saved.read_bytes())), recorded) == []
    ours = excel.resave(operate(book, case), name=f"structure-ours-{name}")
    assert ours.ok, ours          # opens without a repair prompt
    found = measured(facts(ours.saved.read_bytes()))
    if name in REF_SERIES:
        found["charts"] = recorded["charts"] = None
    assert differences(found, recorded) == []
