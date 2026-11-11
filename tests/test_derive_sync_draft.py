"""A batch that derives a dataset from a workbook, then syncs a Word table from it, runs on a
draft of the Word document (round 51; ooxml-edit 0.108.0).

One session holds a made-up report (docx-agent, a test-only dependency) with three tables and a
made-up workbook of sales and products.  Each batch derives with a join and a group
(``xlsx_derive``: the workbook only read), syncs one table from the dataset and sets one
paragraph.  On a draft, the workbook is read at its version as the derive runs; the Word change
is applied, and the dataset kept, only while the workbook is at that version -- else the batch
is prepared again.  So the answers, the report, the workbook and the datasets are those of the
same calls one at a time, in the order they were applied.
"""

from __future__ import annotations

import copy
import hashlib
import threading

import pytest
from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.session import DocumentEntry

docx_agent = pytest.importorskip("docx_agent")
docx_tools = pytest.importorskip("docx_agent.tools")

from xlsx_agent import Workbook  # noqa: E402
from xlsx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

REGIONS = ["North", "South", "East", "West"]


def book() -> bytes:
    sales = [[f"S-{i:03}", REGIONS[i % 4], f"P-{i % 5}", (i * 37) % 101 + 1] for i in range(40)]
    products = [[f"P-{k}", f"Part {k}"] for k in range(5)]
    return Workbook.from_tables({
        "Sales": {"columns": ["Ref", "Region", "Product", "Amount"], "table": "Sales", "rows": sales},
        "Products": {"columns": ["Product", "Name"], "table": "Products", "rows": products},
    }).to_bytes()


def report() -> bytes:
    markdown = "\n\n".join(f"# Part {k}\n\nFigures for part {k}.\n\nRead with care.\n\n"
                           f"| Product | Name | Total |\n|---|---|---|\n| x | y | 0 |\n\nSource: the book.\n\n"
                           f"More on part {k}." for k in range(3))
    with Toolbox(docx_tools.TOOLS, formats=[docx_tools.FORMAT], groups=docx_tools.GROUPS, workers=0) as box:
        session = box.session()
        session.open(docx_agent.Document.new().to_bytes(), "report.docx")
        first = session.entry("d1").document.paragraphs()[0].id
        assert box.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}",
                                                             "markdown": markdown}).ok
        return session.entry("d1").document.to_bytes()


def _box(mode: str) -> Toolbox:
    return Toolbox([*docx_tools.TOOLS, *TOOLS], formats=[docx_tools.FORMAT, FORMAT],
                   groups=[*docx_tools.GROUPS, *GROUPS], workers=0, block_edits=mode)


def _session(box, word: bytes, workbook: bytes):
    session = box.session()
    session.open(word, "report.docx")        # d1
    session.open(workbook, "sales.xlsx")     # d2
    document = session.entry("d1").document
    tables = [table.id for table in document.tables()]
    notes = [p.id for p in document.paragraphs() if p.text.startswith("Source:")]
    return session, tables, notes


def batch(k: int, c: int, tables: list[str], notes: list[str]) -> dict:
    """Derive (join and group: never a filter alone), sync table k, set its note."""
    id_ = f"part{k}"
    return {"ops": [
        {"tool": "xlsx_derive", "arguments": {
            "doc": "d2", "from": "t:Sales", "id": id_,
            "where": [{"column": "Region", "op": "in", "values": REGIONS[k:k + 2]}],
            "join": {"from": "t:Products", "on": ["Product", "Product"], "columns": ["Name"]},
            "group_by": ["Product", "Name"], "measures": [{"fn": "sum", "column": "Amount", "as": "Total"}],
            "sort": [{"column": "Product"}], "top": 2 + (c % 3)}},
        {"tool": "word_edit_table", "arguments": {"doc": "d1", "table": tables[k], "action": "sync",
                                                  "from": f"d:{id_}", "key": "Product", "remove": True}},
        {"tool": "word_set_text", "arguments": {"doc": "d1", "target": notes[k],
                                                "text": f"Source: the book, run {c}."}},
    ]}


def normal(result) -> dict:
    out = copy.deepcopy(result.to_json())
    for op in (out.get("data") or {}).get("ops") or []:
        if isinstance(op.get("data"), dict):
            op["data"].pop("timing", None)
    if isinstance(out.get("data"), dict):
        out["data"].pop("timing", None)
    out["warnings"] = [w for w in out.get("warnings") or [] if "changed while this call read it" not in w]
    return out


def datasets(session) -> dict:
    return {id_: (entry.doc, entry.version, entry.stale, [list(map(str, row)) for row in entry.dataset.rows])
            for id_, entry in session.datasets.items()}


def digest(session, doc: str) -> str:
    return hashlib.sha256(session.entry(doc).document.to_bytes()).hexdigest()


@pytest.fixture(scope="module")
def made():
    return report(), book()


def test_a_derive_then_sync_batch_runs_on_a_draft_as_under_both_locks(made):
    out = {}
    for mode in ("off", "always"):
        with _box(mode) as box:
            session, tables, notes = _session(box, *made)
            results = [box.dispatch(session, "batch", batch(0, 0, tables, notes)),
                       box.dispatch(session, "xlsx_set_cells", {"doc": "d2", "cells": [
                           {"address": "Sales!D3", "number": 500}]}),
                       box.dispatch(session, "batch", batch(1, 1, tables, notes)),
                       box.dispatch(session, "batch", {**batch(0, 2, tables, notes), "atomic": True})]
            assert all(r.ok for r in results), [r.to_json() for r in results if not r.ok]
            out[mode] = ([normal(r) for r in results], digest(session, "d1"), digest(session, "d2"),
                         datasets(session))
            blocks = [r.block for r in results]
    assert out["off"] == out["always"]
    assert blocks[1] is None                                    # the workbook's own change
    assert all(b and b.get("lock") is None and b["tries"] == 1 for k, b in enumerate(blocks) if k != 1), blocks
    assert out["always"][0][0]["doc"] == "d2"                   # names its first document, as before


def test_a_workbook_changed_while_the_batch_is_prepared_makes_it_prepare_again(made, monkeypatch):
    from docx_agent import blocks

    ready, go = threading.Event(), threading.Event()
    original = blocks.Blocks.prepared

    def prepared(self, draft, result, arguments):
        patch = original(self, draft, result, arguments)
        if not ready.is_set():
            ready.set()
            assert go.wait(60)
        return patch

    monkeypatch.setattr(blocks.Blocks, "prepared", prepared)
    with _box("always") as box:
        session, tables, notes = _session(box, *made)
        out: dict = {}
        thread = threading.Thread(target=lambda: out.setdefault("r", box.dispatch(
            session, "batch", batch(1, 0, tables, notes))))
        thread.start()
        assert ready.wait(60)
        # S-001 is South, product P-1: among part 1's rows (South, East).
        assert box.dispatch(session, "xlsx_set_cells", {"doc": "d2", "cells": [
            {"address": "Sales!D3", "number": 900}]}).ok
        go.set()
        thread.join()
        result = out["r"]
        assert result.ok and result.block["tries"] == 2 and result.block["conflicts"] == ["d2 changed"], result.block
        kept = session.datasets["part1"]
        assert kept.version == session.entry("d2").version and not kept.stale
        view = session.entry("d1").document.table(tables[1])
        totals = {view.cell(r, 0).text: view.cell(r, 2).text for r in range(1, len(view.rows))}
        assert int(totals["P-1"]) >= 900                        # the workbook as it is now


class Recorder:
    """The commits in order, by the call that made each (``note_change`` runs under the
    document's lock, as the change is made)."""

    def __init__(self) -> None:
        self.local = threading.local()
        self.commits: list[int] = []
        self.lock = threading.Lock()
        self.original = DocumentEntry.note_change

    def __enter__(self):
        recorder = self

        def note_change(entry, *args):
            recorder.original(entry, *args)
            call = getattr(recorder.local, "call", None)
            if call is not None:
                with recorder.lock:
                    if call not in recorder.commits:
                        recorder.commits.append(call)

        DocumentEntry.note_change = note_change  # type: ignore[method-assign]
        return self

    def __exit__(self, *exc) -> None:
        DocumentEntry.note_change = self.original  # type: ignore[method-assign]


def test_derive_and_sync_batches_from_threads_replay_serially(made):
    """A short replay (4 threads): three each derive and sync their own table three times, a
    fourth changes the workbook; then the same calls one at a time, in the order applied."""
    with _box("auto") as box, Recorder() as recorder:
        session, tables, notes = _session(box, *made)
        plans = [[("batch", batch(k, c, tables, notes)) for c in range(3)] for k in range(3)]
        plans.append([("xlsx_set_cells", {"doc": "d2", "cells": [
            {"address": f"Sales!D{c + 3}", "number": 200 + c}]}) for c in range(3)])
        results: dict = {}

        def worker(k: int) -> None:
            for c, (name, arguments) in enumerate(plans[k]):
                recorder.local.call = (k, c)
                try:
                    results[(k, c)] = box.dispatch(session, name, arguments)
                finally:
                    recorder.local.call = None

        threads = [threading.Thread(target=worker, args=(k,)) for k in range(4)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        assert all(r.ok for r in results.values()), [r.to_json() for r in results.values() if not r.ok]
        assert sorted(recorder.commits) == sorted(results)        # each call made one change
        drafted = sum(1 for r in results.values() if r.block and r.block.get("lock") is None)
        final = (digest(session, "d1"), digest(session, "d2"), datasets(session))
    with _box("off") as box:
        session, _tables, _notes = _session(box, *made)
        again = [normal(box.dispatch(session, *plans[k][c])) for k, c in recorder.commits]
        assert [normal(results[call]) for call in recorder.commits] == again
        assert (digest(session, "d1"), digest(session, "d2"), datasets(session)) == final
    assert drafted > 0
