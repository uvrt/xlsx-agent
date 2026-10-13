"""Time the common read paths, cold, through a real Toolbox (as a model calls them).

    python tools/bench.py                 # every path, both workbooks
    python tools/bench.py --rows 50000    # a smaller big sheet
    python tools/bench.py --profile describe:big   # cProfile one path, top 25 by cumtime

Two made-up workbooks: ``big``, one table of ``--rows`` rows by six columns (a unique code
and a category as shared strings, an amount, a date, a whole number and a flag), and
tests/build.py's ``warehouse`` (eight tables of 24 columns, the largest test fixture).
``compare`` compares ``big`` with a copy in which every hundredth row's amount differs.

Each path runs in a new session on a newly opened workbook, so nothing it reads is kept
from an earlier path; the steps within a path (a first page, then a later one) share it.
Times are wall clock on this machine; compare runs on the same machine only.
"""

from __future__ import annotations

import argparse
import cProfile
import datetime as dt
import pstats
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

from build import NS_MAIN, NS_R, WorkbookBuilder, zip_parts  # noqa: E402
from test_catalogue import warehouse  # noqa: E402

from ooxml_edit.tools import Toolbox  # noqa: E402

from xlsx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
CATEGORIES = ["Hardware", "Tools", "Garden", "Paint", "Lighting", "Plumbing"]
COLUMNS = ["Code", "Category", "Amount", "Day", "Qty", "Open"]


def big(rows: int, *, changed: bool = False) -> bytes:
    """One table, ``Big``, of ``rows`` data rows by six columns; ``changed``: every
    hundredth row's amount is one more."""
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    day = book.style(num_fmt="yyyy-mm-dd")
    sheet.table("Big", f"A1:F{rows + 1}", COLUMNS)
    parts = book.parts()
    names = COLUMNS + CATEGORIES
    codes = len(names)
    items = "".join(f"<si><t>{n}</t></si>" for n in names)
    items += "".join(f"<si><t>SKU-{r:07d}</t></si>" for r in range(rows))
    count = len(names) + rows
    parts["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{NS_MAIN}" '
        f'count="{count}" uniqueCount="{count}">{items}</sst>').encode()
    body = ['<row r="1">' + "".join(f'<c r="{chr(65 + k)}1" t="s"><v>{k}</v></c>'
                                    for k in range(6)) + "</row>"]
    for r in range(2, rows + 2):
        i = r - 2
        amount = round(i * 0.37 % 1000, 2) + (1 if changed and i % 100 == 0 else 0)
        body.append(
            f'<row r="{r}"><c r="A{r}" t="s"><v>{codes + i}</v></c>'
            f'<c r="B{r}" t="s"><v>{len(COLUMNS) + i % len(CATEGORIES)}</v></c>'
            f'<c r="C{r}"><v>{amount}</v></c><c r="D{r}" s="{day}"><v>{45000 + i % 900}</v></c>'
            f'<c r="E{r}"><v>{i % 250}</v></c><c r="F{r}" t="b"><v>{i % 3 == 0:d}</v></c></row>')
    parts["xl/worksheets/sheet1.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><sheetData>{"".join(body)}'
        f"</sheetData></worksheet>").encode()
    if not any(k.endswith("sharedStrings.xml") for k in book.parts()):
        raise SystemExit("the builder wrote no shared strings part to replace")
    return zip_parts(parts)


class Bench:
    def __init__(self) -> None:
        self.toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0)

    def session(self, *blobs: tuple[str, bytes]):
        session = self.toolbox.session(clock=CLOCK)
        for name, data in blobs:
            self.call(session, "open_document", {"blob": session.add_blob(data, name=name)})
        return session

    def call(self, session, name: str, arguments: dict):
        result = self.toolbox.dispatch(session, name, arguments)
        if not result.ok:
            raise SystemExit(f"{name} failed: {result.error}")
        return result


def paths(bench: Bench, books: dict[str, bytes], rows: int):
    """``(name, steps)``: each step ``(label, fn(session))``, run in one session."""
    table = {"big": "t:Big", "warehouse": "t:Products"}
    key = {"big": "Code", "warehouse": "Products column 1"}
    group = {"big": "Category", "warehouse": "Products column 2"}
    amount = {"big": "Amount", "warehouse": "Products column 3"}
    late = {"big": f"Data!A{rows - 50}:F{rows}", "warehouse": "Archive!A600:X650"}
    needle = {"big": "SKU-00001", "warehouse": "Tools"}
    for book in books:
        if book == "compare":
            continue
        yield f"describe:{book}", [("describe", lambda s: bench.call(s, "describe", {"doc": "d1"}))]
        yield f"read:{book}", [
            ("page 1", lambda s, b=book: bench.call(s, "xlsx_read", {"doc": "d1", "target": table[b]})),
            ("late page", lambda s, b=book: bench.call(s, "xlsx_read", {"doc": "d1", "target": late[b]})),
            ("late page again", lambda s, b=book: bench.call(s, "xlsx_read", {"doc": "d1", "target": late[b]}))]
        yield f"find:{book}", _find_steps(bench, needle[book])
        yield f"query:{book}", _query_steps(bench, table[book], group[book], amount[book])
        yield f"derive:{book}", [("derive", lambda s, b=book: bench.call(s, "xlsx_derive", {
            "doc": "d1", "from": table[b], "id": "g", "group_by": [group[b]],
            "measures": [{"fn": "sum", "column": amount[b], "as": "total"},
                         {"fn": "mean", "column": amount[b], "as": "average"}]}))]
        yield f"validate:{book}", [("validate", lambda s, b=book: bench.call(s, "xlsx_validate_data", {
            "doc": "d1", "schema": {"tables": [{"table": table[b], "key": [key[b]], "columns": [
                {"name": amount[b], "type": "number", "min": "0"},
                {"name": group[b], "type": "text", "required": True}]}]}}))]
    if "compare" in books:
        yield "compare:big", [("compare", lambda s: bench.call(s, "xlsx_compare", {
            "a": "d1", "b": "d2", "tables": ["t:Big"], "key": "Code"}))]


def _find_steps(bench: Bench, needle: str):
    state = {}

    def first(s):
        state["r"] = bench.call(s, "find_text", {"doc": "d1", "text": needle})
        return state["r"]

    def second(s):
        cursor = state["r"].next_cursor
        if cursor is None:
            return None
        return bench.call(s, "find_text", {"doc": "d1", "text": needle, "cursor": cursor})
    return [("page 1", first), ("page 2", second)]


def _query_steps(bench: Bench, table: str, group: str, amount: str):
    state = {}
    arguments = {"doc": "d1", "table": table,
                 "where": [{"column": amount, "op": "gte", "value": "10"}],
                 "sort": [{"column": amount, "desc": True}]}

    def first(s):
        state["r"] = bench.call(s, "xlsx_query", arguments)
        return state["r"]

    def second(s):
        cursor = state["r"].next_cursor
        if cursor is None:
            return None
        return bench.call(s, "xlsx_query", {**arguments, "cursor": cursor})
    return [("page 1", first), ("page 2", second)]


def run(bench: Bench, books: dict[str, bytes], rows: int, only: str | None, profile: bool):
    out = []
    for name, steps in paths(bench, books, rows):
        if only and name != only:
            continue
        book = name.split(":")[1]
        blobs = [(f"{book}.xlsx", books[book])]
        if name.startswith("compare"):
            blobs.append(("big2.xlsx", books["compare"]))
        started = time.perf_counter()
        session = bench.session(*blobs)
        opened = time.perf_counter() - started
        timings = [("open", opened)]
        profiler = cProfile.Profile() if profile else None
        for label, step in steps:
            started = time.perf_counter()
            if profiler:
                profiler.enable()
            step(session)
            if profiler:
                profiler.disable()
            timings.append((label, time.perf_counter() - started))
        out.append((name, timings))
        print(f"{name:20s} " + "  ".join(f"{label} {seconds:6.2f}s" for label, seconds in timings),
              flush=True)
        if profiler:
            pstats.Stats(profiler).sort_stats("cumulative").print_stats(25)
            pstats.Stats(profiler).sort_stats("tottime").print_stats(25)
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--rows", type=int, default=200_000)
    parser.add_argument("--only", help="one path, as describe:big")
    parser.add_argument("--profile", metavar="PATH", help="cProfile one path, as describe:big")
    parser.add_argument("--no-big", action="store_true", help="the warehouse only")
    args = parser.parse_args()
    books = {"warehouse": warehouse().to_bytes()}
    if not args.no_big:
        started = time.perf_counter()
        books["big"] = big(args.rows)
        books["compare"] = big(args.rows, changed=True)
        print(f"built {args.rows} rows x 6 in {time.perf_counter() - started:.1f} s "
              f"({len(books['big']) / 1e6:.1f} MB zipped)", flush=True)
    run(Bench(), books, args.rows, args.profile or args.only, bool(args.profile))


if __name__ == "__main__":
    main()
