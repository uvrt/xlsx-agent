"""Driving Microsoft Excel for Mac -- one instance per machine -- as the oracle.

Excel is the only authoritative answer to "will this workbook open without a repair?"
(ROADMAP.md, E1; the second production team's contract, P0-1: Excel for Mac decides), and
there is one of it.  Modelled on docx-agent's Word oracle (its ``tests/oracle.py``):

* **One machine-wide lock.**  Every launch, open, save and recovery happens while holding
  an advisory ``flock`` on :data:`LOCK`, a file in the Office group container.  It is not
  waited for: if someone holds it, the oracle is skipped.
* **Excel already running means someone is using it**, so :func:`session` refuses to start
  (:class:`ExcelBusy`) rather than wait, kill or share; the tests skip with "Excel is in
  use".
* **Recovery only of what it started**: the Excel processes that appeared while the
  session ran (by pid), quit politely first.  An AppleScript call is never killed while it
  may still be talking to Excel unless it is stuck: a killed ``osascript`` can leave an
  AppleEvent queued that starts Excel again later, so after quitting the session watches
  for a few seconds and quits an Excel that comes back -- Excel is left not running.
* **The repair prompt** is found by polling System Events for a dialog with a Yes/No pair
  (Ja/Nee in a Dutch Excel), which is clicked No: the outcome is ``repair``, and the open
  fails on its own (``tools/excel_resave.applescript``).  Alerts stay on for that.
* **Cached by content.**  Outcomes (and the files Excel saved) are kept by the SHA-256 of
  the input and the script in :data:`ORACLE_DIR`, outside the repository: an Excel-saved
  file is never committed.
* **Paths under the sandbox.**  Excel may read and write the Office group container
  without a grant; workbooks are staged there.

Opt in with ``pytest -m oracle`` (never under xdist: one Excel).  Needs the terminal's
Accessibility permission (System Events reads Excel's windows).
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import os
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

try:
    import fcntl
except ImportError:  # Windows: there is no Excel oracle there (it drives Excel for Mac)
    fcntl = None

EXCEL_APP = Path("/Applications/Microsoft Excel.app")
PROCESS = "Microsoft Excel"
SCRIPT = Path(__file__).resolve().parents[1] / "tools" / "excel_resave.applescript"

GROUP_CONTAINER = Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office"
#: The machine-wide Excel lock.
LOCK = GROUP_CONTAINER / "excel-oracle.lock"
#: Where workbooks are staged for Excel and its answers cached.
ORACLE_DIR = GROUP_CONTAINER / "xlsx-agent-oracle"

#: The buttons that answer Excel's repair prompt No, in the languages met.
NO_BUTTONS = ("No", "Nee")
#: How long a run may take: a cold start is 8-13 s (measured), a repair prompt about 2 s.
TIMEOUT = 90.0
#: How long Excel is watched after it was quit, for a queued AppleEvent starting it again.
SETTLE = 6.0

_DIALOG_BUTTONS = ('tell application "System Events" to tell process "Microsoft Excel" to '
                   'get name of buttons of (every window whose subrole is "AXDialog")')
_CLICK = ('tell application "System Events" to tell process "Microsoft Excel" to click '
          '(first button of (first window whose subrole is "AXDialog") whose name is "{}")')


class ExcelBusy(RuntimeError):
    """Excel is in use by someone else (or the lock is held): the oracle does not run."""


def available() -> bool:
    return EXCEL_APP.exists() and SCRIPT.exists() and fcntl is not None


def excel_pids() -> set[int]:
    found = subprocess.run(["pgrep", "-x", PROCESS], capture_output=True, text=True)
    return {int(pid) for pid in found.stdout.split()}


@dataclass
class Outcome:
    """``ok`` (Excel opened and saved it), ``repair`` (Excel asked to repair it; answered
    No) or ``error`` (anything else: ``detail`` says what).  ``saved`` is the file Excel
    wrote, for ``ok``."""

    outcome: str
    detail: str = ""
    saved: Path | None = None
    seconds: float = 0.0
    cached: bool = False

    @property
    def ok(self) -> bool:
        return self.outcome == "ok"


@contextlib.contextmanager
def session() -> Iterator["Session"]:
    """Hold the machine-wide lock for a run of Excel work; :class:`ExcelBusy` if Excel or
    the lock is taken.  Excel is left not running on exit (only an Excel this session
    started is ever quit)."""
    if not available():
        raise ExcelBusy("Microsoft Excel or tools/excel_resave.applescript is not available")
    GROUP_CONTAINER.mkdir(parents=True, exist_ok=True)
    handle = open(LOCK, "a+")
    try:
        try:
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise ExcelBusy("another job holds the Excel lock") from None
        if excel_pids():
            raise ExcelBusy("Excel is in use")
        ORACLE_DIR.mkdir(exist_ok=True)
        current = Session()
        try:
            yield current
        finally:
            current.recover()
    finally:
        fcntl.flock(handle, fcntl.LOCK_UN)
        handle.close()


class Session:
    """Excel work under the lock (made by :func:`session`)."""

    def __init__(self) -> None:
        #: The Excel processes this session started (it began with none running).
        self.started: set[int] = set()
        #: An ``osascript`` was killed mid-call: an AppleEvent of it may still be queued.
        self.orphaned = False

    def resave(self, data: bytes, *, name: str = "book") -> Outcome:
        """``data`` opened by Excel and saved again as a .xlsx; cached by content."""
        digest = hashlib.sha256(data + SCRIPT.read_bytes()).hexdigest()[:16]
        record = ORACLE_DIR / f"{name}-{digest}.json"
        saved = ORACLE_DIR / f"{name}-{digest}-saved.xlsx"
        if record.exists():
            known = json.loads(record.read_text())
            if known["outcome"] != "ok" or saved.exists():
                return Outcome(known["outcome"], known.get("detail", ""),
                               saved if known["outcome"] == "ok" else None, cached=True)
        staged = ORACLE_DIR / f"{name}-{digest}-in.xlsx"
        staged.write_bytes(data)
        saved.unlink(missing_ok=True)
        try:
            outcome = self._run(staged, saved)
        finally:
            staged.unlink(missing_ok=True)
            self._clear_locks()
        if outcome.outcome in ("ok", "repair"):     # an error is not kept: run it again
            record.write_text(json.dumps({"outcome": outcome.outcome,
                                          "detail": outcome.detail}))
        return outcome

    def _run(self, staged: Path, saved: Path) -> Outcome:
        start = time.monotonic()
        before = excel_pids()
        process = subprocess.Popen(["osascript", str(SCRIPT), str(staged), str(saved)],
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        outcome: str | None = None
        detail = ""
        while time.monotonic() - start < TIMEOUT:
            self.started |= excel_pids() - before
            if process.poll() is not None:
                break
            buttons = _dialog_buttons()
            if buttons:
                answer = next((b for b in NO_BUTTONS if b in buttons), None)
                if answer is None:
                    outcome, detail = "error", f"a dialog with {buttons}"
                    break
                _osascript(_CLICK.format(answer))
                outcome = "repair"
            time.sleep(0.5)
        try:
            # After a click the open fails and the script returns by itself; it is killed
            # only when it is stuck.
            out, err = process.communicate(timeout=30 if outcome else max(1.0, TIMEOUT - (
                time.monotonic() - start)))
        except subprocess.TimeoutExpired:
            process.kill()
            self.orphaned = True
            out, err = process.communicate()
            if outcome is None:
                outcome, detail = "error", f"osascript did not return within {TIMEOUT:.0f} s"
        seconds = time.monotonic() - start
        if outcome is None:
            if out.strip() == "done" and saved.exists() and saved.stat().st_size:
                return Outcome("ok", saved=saved, seconds=seconds)
            outcome, detail = "error", (err or out).strip()[:300]
        # Anything but a clean save: start the next file in a fresh Excel.
        self.recover()
        saved.unlink(missing_ok=True)
        return Outcome(outcome, detail, seconds=seconds)

    def recover(self) -> None:
        """Quit the Excel this session started, and see that it stays quit.  After an
        ``osascript`` had to be killed, an Excel that starts within :data:`SETTLE` seconds
        is that call's queued AppleEvent -- this session's own -- and is quit too; otherwise
        an Excel that was not seen starting during a run is never touched."""
        deadline = time.monotonic() + (SETTLE if self.orphaned else 0.0)
        while True:
            ours = excel_pids() & self.started
            if ours:
                self._quit(ours)
                continue
            if self.orphaned and time.monotonic() < deadline:
                fresh = excel_pids() - self.started
                if fresh:
                    self.started |= fresh
                    deadline = time.monotonic() + SETTLE
                    continue
                time.sleep(0.5)
                continue
            break
        self.orphaned = False
        self._clear_locks()

    def _quit(self, pids: set[int]) -> None:
        # "if running" first: a tell to an Excel that has just quit would start it again.
        try:
            _osascript(f'if application "{PROCESS}" is running then tell application '
                       f'"{PROCESS}" to quit saving no', timeout=30)
        except subprocess.TimeoutExpired:
            pass        # blocked on a dialog: signalled below
        for signal in (None, "-TERM", "-KILL"):
            if signal is not None:
                for pid in pids & excel_pids():
                    subprocess.run(["kill", signal, str(pid)], check=False)
            for _ in range(50):
                if not pids & excel_pids():
                    return
                time.sleep(0.2)

    @staticmethod
    def _clear_locks() -> None:
        for lock in ORACLE_DIR.glob("~$*"):
            lock.unlink(missing_ok=True)


def _osascript(script: str, timeout: float = 10) -> str:
    done = subprocess.run(["osascript", "-e", script], capture_output=True, text=True,
                          timeout=timeout)
    return done.stdout.strip()


def _dialog_buttons() -> list[str]:
    """The buttons of Excel's open dialogs, by name; empty when there are none (or Excel
    is not running yet)."""
    try:
        text = _osascript(_DIALOG_BUTTONS, timeout=5)
    except subprocess.TimeoutExpired:
        return []
    return [name.strip() for name in text.split(",") if name.strip()]


__all__ = ["ExcelBusy", "LOCK", "ORACLE_DIR", "Outcome", "Session", "available",
           "excel_pids", "session"]

if __name__ == "__main__":      # pragma: no cover -- a manual check: python tests/oracle.py f
    import sys

    with session() as excel:
        for path in sys.argv[1:]:
            print(path, excel.resave(Path(path).read_bytes(), name=os.path.basename(path)))
