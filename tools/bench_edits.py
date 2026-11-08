"""Time the edit steps on a large sheet, each in a new process (round 44, lane P).

    python tools/bench_edits.py                    # every step, best of three
    python tools/bench_edits.py set_cells undo     # some steps
    python tools/bench_edits.py --rows 20000 --runs 1
    python tools/bench_edits.py --profile set_cells   # cProfile one step, top 30

The workbook is made up (tests/build.py): a sheet ``Data`` of ``--rows`` rows by four
columns (a text, a whole number, a decimal, a flag), a formula every tenth row in column F,
and a sheet ``Other`` of 100 formulas summing ``Data``'s.  Each step runs in a process of
its own on a newly opened workbook, and is timed alone (the open and what the step needs
first are not timed): CPU and wall seconds, and the process's peak resident memory (the
open's included).  Times are this machine's; compare runs on the same machine only.
"""

from __future__ import annotations

import argparse
import cProfile
import io
import pstats
import resource
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

from build import Formula, WorkbookBuilder  # noqa: E402

from xlsx_agent import Workbook  # noqa: E402


def big(rows: int) -> bytes:
    book = WorkbookBuilder()
    sheet = book.sheet("Data")
    sheet.rows("A1", [["Name", "Qty", "Price", "Flag"]]
               + [[f"item {n}", n, n * 0.5, "x"] for n in range(1, rows + 1)])
    for n in range(1, rows // 10 + 1):
        r = n * 10
        sheet.cell(f"F{r}", Formula(f"B{r}*C{r}+SUM(B{r}:B{r + 5})", cached=1))
    other = book.sheet("Other")
    for n in range(1, 101):
        r = n * (rows // 500 or 1)
        other.cell(f"A{n}", Formula(f"SUM(Data!B{r}:B{r + 50})", cached=1))
    return book.to_bytes()


def _one_cell(wb):
    return lambda: wb.set_cells({"Data!A2": "changed"})


def _then_undo(wb):
    wb.set_cells({"Data!A2": "changed"})
    return wb.undo


def _then_to_bytes(wb):
    wb.set_cells({"Data!A2": "changed"})
    return wb.to_bytes


def _then_save(wb):
    wb.set_cells({"Data!A2": "changed"})
    return lambda: wb.save(io.BytesIO())


def _five_steps(wb):
    def run():
        for n in range(5):
            wb.set_cells({f"Data!A{n + 2}": f"step {n}"})
    return run


#: Each step: what it does to a workbook just opened, as a function to time.
STEPS = {
    "read": lambda wb: lambda: list(wb.read(f"Data!A1:D{wb.used_range('Data').last.row}")),
    "validate": lambda wb: wb.validate,
    "set_cells": _one_cell,
    "set_cells_1000": lambda wb: lambda: wb.set_cells(
        {f"Data!B{r}": r * 3 for r in range(2000, 3000)}),
    "append_rows_1000": lambda wb: lambda: wb.append_rows(
        "Data", [[f"new {n}", n, n, "y"] for n in range(1000)]),
    "write_table_1000": lambda wb: lambda: wb.write_table(
        "Other", "C1", ["A", "B", "C"], [[n, f"t{n}", n * 2] for n in range(1000)]),
    "format_column": lambda wb: lambda: wb.format("Data!C:C", number_format="0.00"),
    "insert_rows": lambda wb: lambda: wb.insert_rows("Data", 100, 5),
    "delete_rows": lambda wb: lambda: wb.delete_rows("Data", 100, 5),
    "move_rows": lambda wb: lambda: wb.move_rows("Data", "400:410", 50),
    "undo": _then_undo,
    "to_bytes": _then_to_bytes,
    "save": _then_save,
    "set_cells_x5": _five_steps,
}


def run_one(step: str, rows: int, profile: bool) -> None:
    wb = Workbook.open(big(rows))
    run = STEPS[step](wb)
    cpu, wall = time.process_time(), time.perf_counter()
    if profile:
        profiler = cProfile.Profile()
        profiler.runcall(run)
    else:
        run()
    cpu, wall = time.process_time() - cpu, time.perf_counter() - wall
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    peak = peak / 1e6 if sys.platform == "darwin" else peak / 1e3     # bytes, or KiB
    print(f"{step:18s} cpu {cpu:7.3f} s  wall {wall:7.3f} s  peak {peak:6.0f} MB", flush=True)
    if profile:
        pstats.Stats(profiler).sort_stats("cumulative").print_stats(30)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("steps", nargs="*", help=", ".join(STEPS))
    parser.add_argument("--rows", type=int, default=50_000)
    parser.add_argument("--runs", type=int, default=3)
    parser.add_argument("--profile", action="store_true")
    parser.add_argument("--one", help=argparse.SUPPRESS)
    options = parser.parse_args()
    unknown = [step for step in options.steps if step not in STEPS]
    if unknown:
        parser.error(f"no step {', '.join(unknown)}: one of {', '.join(STEPS)}")
    if options.one:
        run_one(options.one, options.rows, options.profile)
        return
    for step in options.steps or list(STEPS):
        if options.profile:
            run_one(step, options.rows, True)
            continue
        lines = [subprocess.run([sys.executable, __file__, "--one", step, "--rows",
                                 str(options.rows)], check=True, capture_output=True,
                                text=True).stdout.strip() for _ in range(options.runs)]
        print(min(lines, key=lambda line: float(line.split()[2])), flush=True)


if __name__ == "__main__":
    main()
