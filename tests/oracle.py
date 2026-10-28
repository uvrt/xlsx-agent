"""Driving Microsoft Excel for Mac -- one instance per machine -- as the oracle.

Excel is the only authoritative answer to "will this workbook open without a repair?"
(ROADMAP.md, E1; the second production team's contract, P0-1: Excel for Mac decides), and
there is one of it.  Modelled on docx-agent's Word oracle (its ``tests/oracle.py``):

* **One machine-wide lock.**  Every launch, open, save and recovery happens while holding
  an advisory ``flock`` on :data:`LOCK`, a file in the Office group container.  It is not
  waited for: if someone holds it, the oracle is skipped.
* **Excel already running means someone is using it**, so :func:`session` refuses to start
  (:class:`ExcelBusy`) rather than wait, kill or share; the tests skip with "Excel is in
  use".  The one exception is an Excel this oracle started in a run that was killed
  outright (``kill -9``: no cleanup can run): its pid and start time are kept in
  :data:`STARTED` while it runs, and the next session quits it when it is still that
  process and has no workbook open.
* **Recovery only of what it started**: the Excel processes that appeared while the
  session ran (by pid), quit politely first.  An AppleScript call is never killed while it
  may still be talking to Excel unless it is stuck: a killed ``osascript`` can leave an
  AppleEvent queued that starts Excel again later, so after quitting the session watches
  for a few seconds and quits an Excel that comes back -- Excel is left not running.
* **Interrupted, Excel is still quit.**  Ctrl-C, ``SIGTERM`` or ``SIGHUP`` during a run
  ends the session through its cleanup (the ``osascript`` runs in its own process group,
  so Ctrl-C reaches the oracle, not it); signals that arrive while Excel is being quit are
  held until it is; at exit, a session that was never closed is recovered.
* **The repair prompt** is found by polling System Events for a dialog with a Yes/No pair
  (Ja/Nee in a Dutch Excel), which is clicked No: the outcome is ``repair``, and the open
  fails on its own (``tools/excel_resave.applescript``).  Alerts stay on for that.  With
  the screen locked System Events sees no window of Excel's (measured, 2026-10-28), so the
  prompt is found through the window server instead (:func:`_alert_shown`: an untitled
  window of a dialog's size; ScreenCaptureKit images of it, taken while locked, showed the
  repair prompt each time) and cannot be answered: the call is stopped and Excel quit, and
  the outcome is ``repair`` all the same, its ``detail`` saying so.
* **What Excel saved is the input**: the script saves the workbook of the input's file
  name, never the active one, and the saved file must have the input's sheets, by name and
  in order; otherwise the outcome is ``mismatch``.  Measured (2026-10-27/28): Excel can
  start with a blank workbook ("Map1", sheet "Blad1" in a Dutch Excel) that stays active
  after the input is opened, and saving the active workbook saved it -- reported ``ok``.
* **Known causes, named, before Excel.**  Measured causes of the repair prompt
  (:func:`repair_causes`: sheet names Excel treats as equal, rows or cells out of order)
  are found in the file, and such a file is not given to Excel (``preflight``, naming the
  cause) unless asked.  Excel not answering with no dialog found is ``hang``, never a
  plain ``error``.
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
import io
import json
import os
import posixpath
import re
import signal
import subprocess
import threading
import time
import unicodedata
import weakref
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Sequence
from xml.etree import ElementTree

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
#: The file in :data:`ORACLE_DIR` naming the Excel a session started while it runs.
STARTED = "excel-started.json"

#: The buttons that answer Excel's repair prompt No, in the languages met.
NO_BUTTONS = ("No", "Nee")
#: How long a run may take: a cold start is 8-13 s (measured), a repair prompt about 2 s.
TIMEOUT = 90.0
#: How long Excel is watched after it was quit, for a queued AppleEvent starting it again.
SETTLE = 6.0

#: What a run can come to (:class:`Outcome`).
OUTCOMES = ("ok", "repair", "mismatch", "hang", "preflight", "error")
#: An Excel window of a dialog's size, in points (the repair prompt measured 273 x 274-288).
ALERT_SIZE = (200, 600)

_DIALOG_BUTTONS = ('tell application "System Events" to tell process "Microsoft Excel" to '
                   'get name of buttons of (every window whose subrole is "AXDialog")')
_CLICK = ('tell application "System Events" to tell process "Microsoft Excel" to click '
          '(first button of (first window whose subrole is "AXDialog") whose name is "{}")')
_WINDOWS = ('tell application "System Events" to tell process "Microsoft Excel" to '
            'count windows')
# "if running" first: a tell to an Excel that is not running would start it.
_WORKBOOKS = (f'if application "{PROCESS}" is running then tell application "{PROCESS}" to '
              'count workbooks')
#: AppleScript's "AppleEvent timed out": Excel did not answer (the script prints it).
_TIMED_OUT = "error -1712"


class ExcelBusy(RuntimeError):
    """Excel is in use by someone else (or the lock is held): the oracle does not run."""


def available() -> bool:
    return EXCEL_APP.exists() and SCRIPT.exists() and fcntl is not None


def excel_pids() -> set[int]:
    found = subprocess.run(["pgrep", "-x", PROCESS], capture_output=True, text=True)
    return {int(pid) for pid in found.stdout.split()}


@dataclass
class Outcome:
    """What Excel made of a workbook, one of :data:`OUTCOMES`:

    * ``ok``: Excel opened it and saved it, and the file saved is the input (its sheets);
    * ``repair``: Excel asked to repair it (answered No; with the screen locked, seen
      through the window server and not answered, ``detail`` says so);
    * ``mismatch``: Excel saved another workbook than the input, or had none of its name;
    * ``hang``: Excel stopped answering and no dialog was found; ``detail`` names a known
      cause found in the file (:func:`repair_causes`) and what Excel showed;
    * ``preflight``: not given to Excel, because the file holds a known cause of the
      repair prompt (``detail``); ``resave(preflight=False)`` gives it to Excel anyway;
    * ``error``: anything else (``detail`` says what).

    ``saved`` is the file Excel wrote, for ``ok``."""

    outcome: str
    detail: str = ""
    saved: Path | None = None
    seconds: float = 0.0
    cached: bool = False

    @property
    def ok(self) -> bool:
        return self.outcome == "ok"


# -- what the file says, before and after Excel ---------------------------------------------------

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]+)$")

#: Characters Excel leaves out when it compares sheet names.  Measured (Excel for Mac
#: 16.106): ``a<SOFT HYPHEN>b`` and ``a<BOM>b`` are equal, and ``ab`` and
#: ``a<ZERO WIDTH SPACE>b``.
_IGNORED = dict.fromkeys(map(ord, "\u00ad\ufeff\u200b"))


def same_name(name: str) -> str:
    """``name`` as Excel compares sheet names, as far as measured (Excel for Mac 16.106,
    2026-10-27/28; each pair in one workbook makes Excel ask to repair it): case aside (``R``/``r``,
    ``XFD9``/``xfd9``), compatibility forms as their plain ones -- Unicode's NFKC: a
    superscript and a fullwidth one (``a¹b``/``a１b``), a no-break space and a space, a
    halfwidth and a fullwidth katakana KA -- and the characters of :data:`_IGNORED` left
    out.  Wider than that is not measured: a pair this tells apart may still be equal to
    Excel (Excel then answers ``repair``, not the pre-flight)."""
    return unicodedata.normalize("NFKC", name.translate(_IGNORED)).lower()


def _resolve(source: str, target: str) -> str:
    """A relationship's target as a part name (no leading ``/``), ``source`` the part
    whose relationship it is."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def _rels(archive: zipfile.ZipFile, source: str) -> dict[str, str]:
    """``source``'s relationships, id to part name; empty when it has none."""
    folder, base = posixpath.split(source)
    name = posixpath.join(folder, "_rels", base + ".rels")
    try:
        root = ElementTree.fromstring(archive.read(name))
    except (KeyError, ElementTree.ParseError):
        return {}
    return {rel.get("Id", ""): _resolve(source, rel.get("Target", ""))
            for rel in root.iter(_REL + "Relationship")
            if rel.get("TargetMode") != "External"}


def _sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]] | None:
    """The workbook's sheets in order, (name, part name; ``""`` when it has none); ``None``
    when there is no workbook part to read."""
    package = ElementTree.fromstring(archive.read("_rels/.rels"))
    main = next((_resolve("", rel.get("Target", "")) for rel in
                 package.iter(_REL + "Relationship")
                 if rel.get("Type", "").endswith("/officeDocument")), None)
    if main is None:
        return None
    root = ElementTree.fromstring(archive.read(main))
    parts = _rels(archive, main)
    return [(sheet.get("name", ""), parts.get(sheet.get(_R_ID, ""), ""))
            for sheet in root.iter(_MAIN + "sheet")]


def sheet_names(data: bytes) -> list[str] | None:
    """The workbook's sheet names in order; ``None`` when the file cannot be read as one."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            sheets = _sheets(archive)
    except (zipfile.BadZipFile, KeyError, ElementTree.ParseError):
        return None
    return None if sheets is None else [name for name, _ in sheets]


def _column(letters: str) -> int:
    number = 0
    for letter in letters.upper():
        number = number * 26 + ord(letter) - 64
    return number


def _name(column: int, row: int) -> str:
    letters = ""
    while column:
        column, rest = divmod(column - 1, 26)
        letters = chr(65 + rest) + letters
    return f"{letters}{row}"


def _out_of_order(stream) -> str:
    """The first row of a worksheet part listed at or before the row before it (``"row 2
    listed before row 1"``, ``"row 1 listed twice"``), or cell listed at or before the
    column of the cell before it in its row (``"B1 listed before A1"``, ``"A1 listed
    twice"``); else ``""``.  A row or cell without ``r`` is the one after the one before it
    (ECMA-376 Part 1, 18.3.1.73, 18.3.1.4); a cell whose ``r`` cannot be read ends the
    check of its row."""
    last, row, readable = 0, 0, True
    for event, element in ElementTree.iterparse(stream, ("start", "end")):
        if event == "end":
            if element.tag == _MAIN + "row":
                element.clear()
            continue
        if element.tag == _MAIN + "row":
            r = element.get("r")
            number = int(r) if r and r.isdigit() else row + 1
            if number == row:
                return f"row {row} listed twice"
            if number < row:
                return f"row {row} listed before row {number}"
            row, last, readable = number, 0, True
        elif element.tag == _MAIN + "c" and readable:
            ref = element.get("r")
            if ref is None:
                column = last + 1
            else:
                match = _REF.match(ref)
                if match is None:
                    readable = False
                    continue
                column = _column(match.group(1))
            if column == last:
                return f"{_name(column, row)} listed twice"
            if column < last:
                return f"{_name(last, row)} listed before {_name(column, row)}"
            last = column
    return ""


def repair_causes(data: bytes) -> list[str]:
    """What in ``data`` makes Excel for Mac ask to repair it on open, as measured (Excel
    for Mac 16.106, 2026-10-27/28, made-up workbooks through this oracle; first taken for
    hangs, since with the screen locked the prompt cannot be seen -- :func:`screen_locked`);
    empty when none is found or the file cannot be read (Excel then decides):

    * two sheets whose names Excel treats as equal (:func:`same_name`);
    * rows out of order (row 2 listed before row 1), or a row's cells out of column
      order (``B1`` listed before ``A1``), or a row or cell listed twice.  ECMA-376 Part 1
      does not say rows and cells are in order; Excel needs them to be."""
    causes = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            sheets = _sheets(archive) or []
            seen: dict[str, str] = {}
            for name, _ in sheets:
                key = same_name(name)
                if key in seen:
                    causes.append(f"sheet names Excel treats as equal: {seen[key]!r} and "
                                  f"{name!r}")
                else:
                    seen[key] = name
            for name, part in sheets:
                if part not in archive.namelist():
                    continue
                with archive.open(part) as stream:
                    found = _out_of_order(stream)
                if found:
                    causes.append(f"cells out of order in sheet {name!r}: {found}")
    except (zipfile.BadZipFile, KeyError, ElementTree.ParseError):
        return causes
    return causes


def judge(out: str, err: str, *, saved: list[str] | None, expected: list[str] | None,
          causes: Sequence[str] = (), timed_out: bool = False,
          windows: int | None = None, locked: bool = False) -> tuple[str, str]:
    """The outcome and detail of a run that met no dialog, from what the script printed
    (``out``, ``err``), the sheet names of the file Excel saved (``saved``; ``None`` when
    there is none), the input's (``expected``), the input's :func:`repair_causes`, whether
    the script had to be killed (``timed_out``), Excel's windows then (``windows``) and
    whether the screen was locked (``locked``: a hang may then be a repair prompt that
    could not be seen, :func:`screen_locked`)."""
    out = out.strip()
    if timed_out or out.startswith(_TIMED_OUT):
        waited = (f"osascript did not return within {TIMEOUT:.0f} s" if timed_out
                  else "Excel did not answer (AppleEvent time-out)")
        shown = ("" if windows is None or locked else ", with no window open"
                 if windows == 0 else
                 f", with {windows} window{'s' if windows != 1 else ''} open")
        cause = ("; ".join(causes) + " (Excel asks to repair such a file)" if causes
                 else "no known cause in the file")
        if locked:
            cause += ("; the screen is locked, and System Events then sees no window of "
                      "Excel's, so a repair prompt the window server did not show either "
                      "cannot be told from a hang: unlock it and run again")
        return "hang", f"{waited}{shown}: {cause}"
    if out.startswith("wrong workbook:"):
        return "mismatch", ("Excel has no workbook of the input's name open (it has "
                            f"{out.partition(':')[2].strip()!r}); nothing saved")
    if out == "done":
        if saved is None:
            return "error", "Excel reported a save but the saved file cannot be read"
        if expected is not None and saved != expected:
            return "mismatch", f"Excel saved the sheets {saved}; the input has {expected}"
        return "ok", ""
    return "error", (err.strip() or out)[:300]


# -- interruptions --------------------------------------------------------------------------------

#: Sessions not yet recovered: :func:`_recover_live` quits their Excel at exit.
_LIVE: list["Session"] = []
#: What :func:`_interruptible` turns into Ctrl-C; :func:`_held` holds these and Ctrl-C.
_TERMINATIONS = tuple(getattr(signal, name) for name in ("SIGTERM", "SIGHUP")
                      if hasattr(signal, name))


def _main_thread() -> bool:
    return threading.current_thread() is threading.main_thread()


@contextlib.contextmanager
def _interruptible() -> Iterator[None]:
    """``SIGTERM`` and ``SIGHUP`` raise :class:`KeyboardInterrupt`, as Ctrl-C does, so a
    session's cleanup runs; a handler someone else installed is left alone."""
    if not _main_thread():
        yield
        return

    def stop(signum, frame):
        raise KeyboardInterrupt(f"signal {signum}")

    previous = {}
    for stopper in _TERMINATIONS:
        if signal.getsignal(stopper) is signal.SIG_DFL:
            previous[stopper] = signal.signal(stopper, stop)
    try:
        yield
    finally:
        for stopper, handler in previous.items():
            signal.signal(stopper, handler)


@contextlib.contextmanager
def _held() -> Iterator[None]:
    """Ctrl-C, ``SIGTERM`` and ``SIGHUP`` held while Excel is quit, and raised as
    :class:`KeyboardInterrupt` after: a second Ctrl-C does not leave Excel half-quit."""
    if not _main_thread():
        yield
        return
    caught: list[int] = []
    previous = {stopper: signal.signal(stopper, lambda signum, frame: caught.append(signum))
                for stopper in (signal.SIGINT, *_TERMINATIONS)}
    try:
        yield
    finally:
        for stopper, handler in previous.items():
            signal.signal(stopper, handler)
    if caught:
        raise KeyboardInterrupt(f"signal {caught[0]}")


def _recover_live() -> None:
    """Quit the Excel of a session that was never closed (an interpreter leaving without
    running its cleanup: a test runner stopped mid-fixture).  Called at exit."""
    for live in list(_LIVE):
        with contextlib.suppress(Exception, KeyboardInterrupt):
            live.recover()
        with contextlib.suppress(ValueError):
            _LIVE.remove(live)


class _Exit:
    """An object that lives as long as the module, so its finalizer runs at exit."""


_EXIT = _Exit()
#: ``weakref.finalize`` calls ``_recover_live`` at interpreter exit (its ``atexit``).
_AT_EXIT = weakref.finalize(_EXIT, _recover_live)


def _start_time(pid: int) -> str:
    """When process ``pid`` started, as ``ps`` says; ``""`` when it is not running."""
    found = subprocess.run(["ps", "-o", "lstart=", "-p", str(pid)], capture_output=True,
                           text=True)
    return found.stdout.strip()


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _leftover(running: set[int]) -> set[int] | None:
    """``running`` when every one of them is an Excel a killed oracle run started (named in
    :data:`STARTED` with its start time, that run's process gone) and has no workbook
    open; else ``None`` -- someone's Excel, never touched."""
    try:
        record = json.loads((ORACLE_DIR / STARTED).read_text())
        owner, started = int(record["owner"]), dict(record["excel"])
    except (OSError, ValueError, KeyError, TypeError):
        return None
    if owner != os.getpid() and _alive(owner):
        return None
    if any(started.get(str(pid)) != _start_time(pid) or not started.get(str(pid))
           for pid in running):
        return None
    try:
        if _osascript(_WORKBOOKS, timeout=10) != "0":
            return None
    except subprocess.TimeoutExpired:
        return None         # not answering: what it holds cannot be asked, so it is kept
    return set(running)


# -- the session ----------------------------------------------------------------------------------


@contextlib.contextmanager
def session() -> Iterator["Session"]:
    """Hold the machine-wide lock for a run of Excel work; :class:`ExcelBusy` if Excel or
    the lock is taken.  Excel is left not running on exit, an interrupted one too (only an
    Excel this session, or a killed run of the oracle, started is ever quit)."""
    if not available():
        raise ExcelBusy("Microsoft Excel or tools/excel_resave.applescript is not available")
    GROUP_CONTAINER.mkdir(parents=True, exist_ok=True)
    handle = open(LOCK, "a+")
    try:
        try:
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise ExcelBusy("another job holds the Excel lock") from None
        ORACLE_DIR.mkdir(exist_ok=True)
        for stale in ORACLE_DIR.glob("*-in.xlsx"):     # staged by a run killed outright
            stale.unlink(missing_ok=True)
        current = Session()
        running = excel_pids()
        if running:
            leftover = _leftover(running)
            if leftover is None:
                raise ExcelBusy("Excel is in use")
            current.started = leftover        # a killed run's own Excel: quit it now
            current.recover()
            if excel_pids():
                raise ExcelBusy("Excel is in use")
        _LIVE.append(current)
        try:
            with _interruptible():
                yield current
        finally:
            current.recover()
            with contextlib.suppress(ValueError):
                _LIVE.remove(current)
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
        #: Excel is held by a prompt nobody can answer: it is quit by signal.
        self.modal = False

    def resave(self, data: bytes, *, name: str = "book", preflight: bool = True) -> Outcome:
        """``data`` opened by Excel and saved again as a .xlsx; cached by content (``ok``
        and ``repair`` only).  A file holding a known cause of the repair prompt is not
        given to Excel (``preflight``) unless ``preflight`` is false."""
        causes = repair_causes(data)
        if causes and preflight:
            return Outcome("preflight", "; ".join(causes) + " (Excel asks to repair such a "
                           "file; not run: preflight=False runs it)")
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
            outcome = self._run(staged, saved, sheet_names(data), causes)
        finally:
            staged.unlink(missing_ok=True)
            self._clear_locks()
        if outcome.outcome in ("ok", "repair"):     # nothing else is kept: run it again
            record.write_text(json.dumps({"outcome": outcome.outcome,
                                          "detail": outcome.detail}))
        return outcome

    def _run(self, staged: Path, saved: Path, expected: list[str] | None = None,
             causes: Sequence[str] = ()) -> Outcome:
        start = time.monotonic()
        before = excel_pids()
        # Its own process group: Ctrl-C at the terminal reaches the oracle, which decides.
        process = subprocess.Popen(["osascript", str(SCRIPT), str(staged), str(saved)],
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                                   start_new_session=True)
        try:
            return self._watch(process, start, before, staged, saved, expected, causes)
        except BaseException:
            # Interrupted: whatever started is this session's, and a call still running is
            # stopped -- its AppleEvent may be queued, which recover() watches for.
            fresh = excel_pids() - before - self.started
            if fresh:
                self._note(fresh)
            if process.poll() is None:
                process.kill()
                self.orphaned = True
            with contextlib.suppress(Exception):
                process.communicate(timeout=5)
            raise

    def _watch(self, process: subprocess.Popen, start: float, before: set[int],
               staged: Path, saved: Path, expected: list[str] | None,
               causes: Sequence[str]) -> Outcome:
        outcome: str | None = None
        detail = ""
        locked = screen_locked()
        while time.monotonic() - start < TIMEOUT:
            fresh = excel_pids() - before - self.started
            if fresh:
                self._note(fresh)
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
            elif locked and outcome is None and _alert_shown():
                # It cannot be answered: the call is stopped, and Excel, held by the
                # prompt, is quit by signal (recover(); a polite quit would wait on it).
                outcome, detail = "repair", UNANSWERED
                process.kill()
                self.orphaned = self.modal = True
                break
            time.sleep(0.5)
        fresh = excel_pids() - before - self.started     # started since the last look
        if fresh:
            self._note(fresh)
        timed_out = False
        try:
            # After a click the open fails and the script returns by itself; it is killed
            # only when it is stuck.
            out, err = process.communicate(timeout=30 if outcome else max(1.0, TIMEOUT - (
                time.monotonic() - start)))
        except subprocess.TimeoutExpired:
            process.kill()
            self.orphaned = True
            out, err = process.communicate()
            timed_out = outcome is None
        seconds = time.monotonic() - start
        if outcome is None:
            hung = timed_out or out.strip().startswith(_TIMED_OUT)
            exists = saved.exists() and saved.stat().st_size > 0
            outcome, detail = judge(out, err, saved=sheet_names(saved.read_bytes())
                                    if exists else None, expected=expected, causes=causes,
                                    timed_out=timed_out,
                                    windows=_window_count() if hung else None,
                                    locked=hung and screen_locked())
            if outcome == "ok":
                return Outcome("ok", saved=saved, seconds=seconds)
        # Anything but a clean save: start the next file in a fresh Excel.
        self.recover()
        saved.unlink(missing_ok=True)
        return Outcome(outcome, detail, seconds=seconds)

    def _note(self, fresh: set[int]) -> None:
        """Add ``fresh`` to the Excel this session started, and keep :data:`STARTED` (the
        pids and start times, for a later session should this process be killed) as
        what of it still runs: removed when nothing does."""
        self.started |= fresh
        record = ORACLE_DIR / STARTED
        running = {pid: _start_time(pid) for pid in self.started & excel_pids()}
        running = {pid: when for pid, when in running.items() if when}
        if running:
            record.write_text(json.dumps({"owner": os.getpid(),
                                          "excel": {str(p): w for p, w in running.items()}}))
        else:
            record.unlink(missing_ok=True)

    def recover(self) -> None:
        """Quit the Excel this session started, and see that it stays quit.  After an
        ``osascript`` had to be killed, an Excel that starts within :data:`SETTLE` seconds
        is that call's queued AppleEvent -- this session's own -- and is quit too; otherwise
        an Excel that was not seen starting during a run is never touched.  Ctrl-C and
        termination signals wait until it is done."""
        with _held():
            deadline = time.monotonic() + (SETTLE if self.orphaned else 0.0)
            while True:
                ours = excel_pids() & self.started
                if ours:
                    self._quit(ours)
                    continue
                if self.orphaned and time.monotonic() < deadline:
                    fresh = excel_pids() - self.started
                    if fresh:
                        self._note(fresh)
                        deadline = time.monotonic() + SETTLE
                        continue
                    time.sleep(0.5)
                    continue
                break
            self.orphaned = self.modal = False
            self._note(set())
            self._clear_locks()

    def _quit(self, pids: set[int]) -> None:
        # "if running" first: a tell to an Excel that has just quit would start it again.
        try:
            if not self.modal:
                _osascript(f'if application "{PROCESS}" is running then tell application '
                           f'"{PROCESS}" to quit saving no', timeout=30)
        except subprocess.TimeoutExpired:
            pass        # blocked on a dialog: signalled below
        for signal_name in (("-TERM", "-KILL") if self.modal else (None, "-TERM", "-KILL")):
            if signal_name is not None:
                for pid in pids & excel_pids():
                    subprocess.run(["kill", signal_name, str(pid)], check=False)
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


#: The ``detail`` of a repair prompt seen through the window server.
UNANSWERED = ("Excel showed its repair prompt (an untitled window of a dialog's size, seen "
              "through the window server): the screen is locked, so it could not be read "
              "or answered, and Excel was quit")
_ALERTS = """ObjC.import('CoreGraphics');
var all = ObjC.deepUnwrap(ObjC.castRefToObject($.CGWindowListCopyWindowInfo(
    $.kCGWindowListOptionAll, $.kCGNullWindowID)));
all.filter(function (w) { return w.kCGWindowOwnerName === 'Microsoft Excel'; })
   .map(function (w) { return [w.kCGWindowName || '', w.kCGWindowLayer,
        w.kCGWindowBounds.Width, w.kCGWindowBounds.Height].join('|'); }).join('\\n');"""


def _alert_shown() -> bool:
    """Whether Excel has a window of a dialog's size and no title, as the window server
    lists them (``CGWindowListCopyWindowInfo`` through JavaScript for Automation: it does
    not ask Excel, and it works with the screen locked, where System Events lists no
    window).  Measured (Excel for Mac 16.106, 2026-10-28, screen locked): the repair
    prompt is such a window, 273 x 274-288 points; a workbook's window has its title."""
    try:
        done = subprocess.run(["osascript", "-l", "JavaScript", "-e", _ALERTS],
                              capture_output=True, text=True, timeout=5)
    except subprocess.TimeoutExpired:
        return False
    return any(_is_alert(line) for line in done.stdout.splitlines())


def _is_alert(line: str) -> bool:
    """A line of :data:`_ALERTS`'s output (title, layer, width, height) a dialog?"""
    try:
        title, layer, width, height = line.rsplit("|", 3)
        low, high = ALERT_SIZE
        return (not title and int(layer) == 0
                and low <= float(width) <= high and low <= float(height) <= high)
    except ValueError:
        return False


def screen_locked() -> bool:
    """Whether the Mac's screen is locked (``ioreg``'s ``CGSSessionScreenIsLocked``).
    Measured (Excel for Mac 16.106, 2026-10-28): while it is, Excel opens and saves
    workbooks as ever, but System Events lists no window of Excel's -- not an open
    workbook's, not the repair prompt's -- so the prompt cannot be answered; the window
    server still lists them (:func:`_alert_shown`)."""
    try:
        found = subprocess.run(["ioreg", "-n", "Root", "-d1", "-a"], capture_output=True,
                               text=True, timeout=10).stdout
    except (OSError, subprocess.TimeoutExpired):
        return False
    state = re.search(r"<key>CGSSessionScreenIsLocked</key>\s*<(true|false)/>", found)
    return state is not None and state.group(1) == "true"


def _window_count() -> int | None:
    """How many windows Excel shows, as System Events sees them (it does not ask Excel);
    ``None`` when that cannot be read (with the screen locked it reads none, measured)."""
    try:
        text = _osascript(_WINDOWS, timeout=5)
    except subprocess.TimeoutExpired:
        return None
    return int(text) if text.isdigit() else None


__all__ = ["ExcelBusy", "LOCK", "ORACLE_DIR", "OUTCOMES", "Outcome", "STARTED", "Session",
           "available", "excel_pids", "judge", "repair_causes", "same_name",
           "screen_locked", "session", "sheet_names"]

if __name__ == "__main__":      # pragma: no cover -- a manual check: python tests/oracle.py f
    import sys

    with session() as excel:
        for path in sys.argv[1:]:
            print(path, excel.resave(Path(path).read_bytes(), name=os.path.basename(path)))
