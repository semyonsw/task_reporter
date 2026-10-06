"""Task Reporter - files task reports into task_reports.xlsx.

Two surfaces run at the same time from a single process:

  * the browser UI       - type the report, press Ctrl+Enter
  * the terminal console - type the report at the `report>` prompt, press Enter

Closing either surface ends the whole session, so they always disappear
together.  If neither the browser nor a display can be reached, the terminal
console simply becomes the only surface instead of the program dying.

The UI is served over loopback and drawn by the system browser rather than by a
desktop toolkit.  That is a deliberate reliability choice, not a stylistic one -
see the "Browser UI" section for why the Qt window it replaced could only ever
be intermittent under WSLg.  The Qt window is still available behind --qt.
"""

try:
    from PySide6.QtWidgets import (
        QApplication,
        QMainWindow,
        QWidget,
        QVBoxLayout,
        QHBoxLayout,
        QLabel,
        QPlainTextEdit,
        QPushButton,
        QProgressBar,
        QFrame,
        QMenu,
        QMessageBox,
        QDialog,
        QSizePolicy,
        QScrollArea,
        QTableWidget,
        QTableWidgetItem,
        QHeaderView,
        QAbstractItemView,
    )
    from PySide6.QtCore import Qt, QTimer
    from PySide6.QtGui import QKeySequence, QShortcut, QTextCursor, QAction, QCursor

    GUI_AVAILABLE = True
except ModuleNotFoundError:
    GUI_AVAILABLE = False
    # Allow class declarations to be parsed when GUI libs are missing.
    QDialog = object
    QMainWindow = object

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
import argparse
import base64
import functools
import http.server
import json
import mimetypes
import os
import re
import secrets
import shutil
import signal
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from collections import deque
from datetime import datetime

# True when running as the packaged Windows .exe rather than as a script.
FROZEN = bool(getattr(sys, "frozen", False))


def app_home_path() -> str:
    """The per-user folder the app keeps its own things in (not the data).

    %LOCALAPPDATA%\\TaskReporter on Windows.  The WebView2 profile lives in
    here, and so does settings.json - which is why the data folder can be
    chosen in Settings without that choice being stored in the data folder.
    """
    root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(root, "TaskReporter")


def app_storage_path() -> str:
    """Where WebView2 keeps its profile, so localStorage survives a restart."""
    return os.path.join(app_home_path(), "webview")


SETTINGS_PATH = os.path.join(app_home_path(), "settings.json")
SETTINGS_DEFAULTS = {
    "dataDir": "",
    "theme": "system",
    "keepInTray": False,
    "hotkey": "Ctrl+Alt+T",
}


def load_settings() -> dict:
    """settings.json merged over the defaults.  Never raises."""
    settings = dict(SETTINGS_DEFAULTS)
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as handle:
            stored = json.load(handle)
        if isinstance(stored, dict):
            for key in SETTINGS_DEFAULTS:
                if key in stored:
                    settings[key] = stored[key]
    except (OSError, ValueError):
        pass
    if settings["theme"] not in ("light", "dark", "system"):
        settings["theme"] = "system"
    settings["keepInTray"] = bool(settings["keepInTray"])
    settings["dataDir"] = str(settings["dataDir"] or "").strip()
    settings["hotkey"] = str(settings["hotkey"] or "").strip()
    return settings


def save_settings(changes: dict) -> dict:
    """Merge `changes` into settings.json atomically and return the result."""
    settings = load_settings()
    for key, value in changes.items():
        if key in SETTINGS_DEFAULTS:
            settings[key] = value
    os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
    temp_path = SETTINGS_PATH + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(settings, handle, indent=2, ensure_ascii=False)
    os.replace(temp_path, SETTINGS_PATH)
    return load_settings()


# Where BASE_DIR came from, in words - shown in Settings, because "why is my
# workbook over there?" is the first question a moved data folder raises.
BASE_DIR_SOURCE = ""


def _resolve_base_dir() -> str:
    """Where the workbook and the task board live.

    As a script this is simply the folder holding this file.  Frozen into an
    .exe it cannot be: PyInstaller unpacks the bundle into a temporary folder
    that is deleted on exit, so `__file__` would put the workbook somewhere it
    disappears from.  The exe therefore looks for the project the same way
    task_reporter.bat does - the environment override first, then the folder
    chosen in Settings, then a folder that actually holds the data, then the
    path recorded when the exe was built.
    """
    global BASE_DIR_SOURCE

    override = (os.environ.get("TASK_REPORT_DIR") or "").strip()
    if override and os.path.isdir(override):
        BASE_DIR_SOURCE = "%TASK_REPORT_DIR%"
        return os.path.abspath(override)

    chosen = load_settings()["dataDir"]
    if chosen and os.path.isdir(chosen):
        BASE_DIR_SOURCE = "the folder chosen in Settings"
        return os.path.abspath(chosen)

    if not FROZEN:
        BASE_DIR_SOURCE = "the folder the script is in"
        return os.path.dirname(os.path.abspath(__file__))

    candidates = [os.path.dirname(os.path.abspath(sys.executable))]

    # Written into the bundle at build time - see windows/build.bat.
    recorded = os.path.join(getattr(sys, "_MEIPASS", ""), "project_home.txt")
    try:
        with open(recorded, "r", encoding="utf-8") as handle:
            home = handle.read().strip()
        if home and os.path.isdir(home):
            candidates.append(home)
    except OSError:
        pass

    for position, candidate in enumerate(candidates):
        if os.path.isfile(os.path.join(candidate, "task_reports.xlsx")) or (
            os.path.isfile(os.path.join(candidate, "task_board.json"))
        ):
            BASE_DIR_SOURCE = (
                "the exe folder" if position == 0 else "windows\\project_home.txt"
            )
            return candidate
    # Nothing found yet: the first run creates the workbook beside the exe.
    BASE_DIR_SOURCE = (
        "windows\\project_home.txt" if len(candidates) > 1 else "the exe folder"
    )
    return candidates[-1]


BASE_DIR = _resolve_base_dir()
EXCEL_FILE_NAME = "task_reports.xlsx"
EXCEL_FILE_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
# Reports land here when the workbook cannot be written (usually because it is
# open in Excel).  They are merged back in automatically on the next save.
PENDING_FILE_PATH = os.path.join(BASE_DIR, ".task_reports_pending.jsonl")

# The task board lives beside the workbook rather than inside it.  Ticking a
# box has to work while task_reports.xlsx is open in Excel, and a plain JSON
# file is the only store that can promise that.
TASKS_FILE_NAME = "task_board.json"
TASKS_FILE_PATH = os.path.join(BASE_DIR, TASKS_FILE_NAME)

# Files attached to a task: task_files/<task id>/<name>.  Removing one, or
# deleting its task, moves it into task_files/.trash/ rather than deleting it,
# which is what lets Undo bring it back.
TASK_FILES_DIR = os.path.join(BASE_DIR, "task_files")
TASK_FILES_TRASH = os.path.join(TASK_FILES_DIR, ".trash")
MAX_TASK_FILE_BYTES = 50 * 1024 * 1024

MAX_REPORT_LENGTH = 2000
MAX_TASK_LENGTH = 600
MAX_PROJECT_LENGTH = 60
TIMESTAMP_FORMAT = "%d/%m/%Y %H:%M:%S"
# Dates are stored ISO (sortable) and shown the way the task list writes them.
DATE_STORE_FORMAT = "%Y-%m-%d"
DATE_DISPLAY_FORMAT = "%d.%m.%Y"

RELAUNCH_ENV_FLAG = "TASK_REPORT_TK_RELAUNCH"
MOVS_PYTHON = "/root/miniconda3/envs/movs/bin/python"
MOVS_RELAUNCH_FLAG = "TASK_REPORT_MOVS_RELAUNCH"
QT_PROBE_TIMEOUT = 30


# ---------------------------------------------------------------------------
# Interpreter bootstrapping
# ---------------------------------------------------------------------------


def _python_has_pyside6(python_executable: str) -> bool:
    try:
        result = subprocess.run(
            [python_executable, "-c", "import PySide6"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        return result.returncode == 0
    except Exception:
        return False


def relaunch_with_gui_if_possible(script_path: str, args: list) -> bool:
    if GUI_AVAILABLE or os.environ.get(RELAUNCH_ENV_FLAG) == "1":
        return False

    candidates: list = []

    preferred_python = os.environ.get("TASK_REPORT_PYTHON_GUI")
    if preferred_python:
        candidates.append(preferred_python)

    for name in ("python3", "python"):
        resolved = shutil.which(name)
        if resolved:
            candidates.append(resolved)

    if os.path.exists(MOVS_PYTHON):
        candidates.append(MOVS_PYTHON)

    seen = set()
    unique_candidates = []
    for candidate in candidates:
        absolute = os.path.abspath(candidate)
        if absolute not in seen and absolute != os.path.abspath(sys.executable):
            seen.add(absolute)
            unique_candidates.append(absolute)

    for candidate in unique_candidates:
        if _python_has_pyside6(candidate):
            env = os.environ.copy()
            env[RELAUNCH_ENV_FLAG] = "1"
            os.execvpe(candidate, [candidate, script_path, *args], env)

    return False


def ensure_movs_python(script_path: str, args: list) -> None:
    """Hop into the `movs` interpreter, but only when that is an upgrade.

    Re-execing into an interpreter that lacks PySide6 used to be a silent way
    to lose the GUI, so the hop is skipped when the current interpreter can
    already draw the window or the target cannot.
    """
    if os.environ.get(MOVS_RELAUNCH_FLAG) == "1":
        return

    current_python = os.path.abspath(sys.executable)
    target_python = os.path.abspath(MOVS_PYTHON)

    if current_python == target_python:
        return

    if GUI_AVAILABLE:
        return

    if not os.path.exists(target_python) or not _python_has_pyside6(target_python):
        return

    env = os.environ.copy()
    env[MOVS_RELAUNCH_FLAG] = "1"
    os.execvpe(target_python, [target_python, script_path, *args], env)


# ---------------------------------------------------------------------------
# Qt platform probing
# ---------------------------------------------------------------------------

# Creating a QApplication only proves a plugin loaded. A compositor can accept
# the connection and still never present the surface - that is the WSLg
# "window is in the taskbar but nothing is on screen" failure. So the probe
# opens a real toplevel and insists on seeing it mapped AND painted.
#
# The window is frameless, tool-class (no taskbar button) and fully
# transparent, so it does not flash anything visible on the way past.
_QT_RENDER_PROBE_SNIPPET = r"""
import sys
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget

painted = {"count": 0}


class Probe(QWidget):
    def paintEvent(self, event):
        painted["count"] += 1
        super().paintEvent(event)


app = QApplication(sys.argv)
probe = Probe(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
probe.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
probe.setWindowOpacity(0.0)
probe.resize(240, 160)
QVBoxLayout(probe).addWidget(QLabel("probe"))
probe.show()

DEADLINE_MS = 4000
elapsed = {"ms": 0}
INTERVAL_MS = 100


def report(ok):
    handle = probe.windowHandle()
    sys.stdout.write(
        "QT_RENDER_%s platform=%s exposed=%s paints=%d size=%dx%d\n"
        % (
            "OK" if ok else "FAIL",
            app.platformName(),
            bool(handle and handle.isExposed()),
            painted["count"],
            probe.width(),
            probe.height(),
        )
    )
    sys.stdout.flush()
    app.quit()


def poll():
    handle = probe.windowHandle()
    exposed = bool(handle and handle.isExposed())
    if exposed and painted["count"] > 0 and probe.width() > 0 and probe.height() > 0:
        report(True)
        return
    elapsed["ms"] += INTERVAL_MS
    if elapsed["ms"] >= DEADLINE_MS:
        report(False)


timer = QTimer()
timer.timeout.connect(poll)
timer.start(INTERVAL_MS)
app.exec()
"""


# Platforms that put pixels on a real screen. The others Qt offers - offscreen,
# minimal, vnc, linuxfb, eglfs - will happily map and paint into something the
# user cannot see, which passes a render check and then hangs the event loop on
# a window nobody can find or close. So they are never chosen automatically.
VISUAL_QT_PLATFORMS = ("wayland", "wayland-egl", "xcb")


def _qt_platform_candidates(forced: str = None) -> list:
    """Platforms to try, best first.

    Native wayland is preferred because it is the path WSLg is built around.
    xcb (via XWayland) is the backstop: it survives some compositor states in
    which the wayland surface is created but never presented.
    """
    candidates = []

    # A deliberate choice - our own flag or our own env var - is honoured as
    # given, including non-visual platforms, because that is what it is for.
    deliberate = (
        forced or os.environ.get("TASK_REPORT_QT_PLATFORM") or ""
    ).strip()
    if deliberate:
        candidates.append(deliberate)

    # QT_QPA_PLATFORM can be left behind by unrelated tooling, so it only gets
    # a say when it names something the user could actually look at.
    foreign = (os.environ.get("QT_QPA_PLATFORM") or "").strip()
    if foreign and foreign in VISUAL_QT_PLATFORMS:
        candidates.append(foreign)

    if (os.environ.get("WAYLAND_DISPLAY") or "").strip():
        candidates.append("wayland")
    if (os.environ.get("DISPLAY") or "").strip():
        candidates.append("xcb")

    seen = set()
    ordered = []
    for candidate in candidates:
        if candidate and candidate not in seen:
            seen.add(candidate)
            ordered.append(candidate)
    return ordered


def _probe_one_platform(candidate: str):
    """Run the render probe for one platform.  Returns (ok, detail)."""
    env = os.environ.copy()
    env["QT_QPA_PLATFORM"] = candidate
    # Keep the probe's own diagnostics out of the captured output.
    env["QT_LOGGING_RULES"] = "qt.qpa.*=false"
    try:
        result = subprocess.run(
            [sys.executable, "-c", _QT_RENDER_PROBE_SNIPPET],
            env=env,
            capture_output=True,
            text=True,
            timeout=QT_PROBE_TIMEOUT,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return False, "probe timed out (compositor never answered)"
    except Exception as exc:
        return False, f"probe could not run: {exc}"

    stdout = result.stdout or ""

    # The verdict is printed and flushed before teardown, so a crash while
    # Qt unwinds does not invalidate a window that demonstrably rendered.
    for line in stdout.splitlines():
        if line.startswith("QT_RENDER_OK"):
            return True, line.strip()
        if line.startswith("QT_RENDER_FAIL"):
            return False, "window never rendered - " + line.split(" ", 1)[-1].strip()

    stderr_lines = [
        line.strip()
        for line in (result.stderr or "").splitlines()
        if line.strip() and "xcb-cursor0" not in line
    ]
    if stderr_lines:
        return False, stderr_lines[0]
    if result.returncode != 0:
        return False, f"probe exited with code {result.returncode}"
    return False, "probe produced no verdict"


def probe_qt_platform(forced: str = None):
    """Find a Qt platform that can actually put a window on screen.

    Two distinct failures are covered:

    1. No platform plugin loads at all.  Qt calls qFatal() for this, which
       aborts the *whole process* with SIGABRT rather than raising - so a
       try/except around QApplication() never reaches its fallback.  Probing in
       a throwaway subprocess makes the abort harmless.

    2. A plugin loads, a window is created, and the compositor never presents
       it.  That is the WSLg "taskbar button but no window" state, and an
       init-only check sails straight past it.  The probe therefore requires a
       window that is both exposed and painted.

    Returns (platform_name, attempts).  platform_name is None when no platform
    can render; attempts is a list of (candidate, ok, detail) for --doctor.
    """
    if not GUI_AVAILABLE:
        return None, [("-", False, "PySide6 is not importable in this interpreter")]

    candidates = _qt_platform_candidates(forced)
    if not candidates:
        return None, [("-", False, "neither DISPLAY nor WAYLAND_DISPLAY is set")]

    attempts = []
    for candidate in candidates:
        ok, detail = _probe_one_platform(candidate)
        attempts.append((candidate, ok, detail))
        if ok:
            return candidate, attempts

    return None, attempts


# ---------------------------------------------------------------------------
# Workbook access (shared by both surfaces, hence the lock)
# ---------------------------------------------------------------------------

_EXCEL_LOCK = threading.RLock()

# Reports are written with line breaks in them - one line per task, a blank
# line between projects.  Excel draws a multi-line cell as a single
# run-together line unless the cell says wrap_text, and cells written by
# openpyxl carry no alignment at all, so a filed report looked concatenated in
# the workbook even though the newlines were there the whole time.
REPORT_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
STAMP_ALIGNMENT = Alignment(vertical="top")


def _style_report_cells(ws):
    """Give every row the wrapping that makes a multi-line report readable."""
    for row in ws.iter_rows(min_row=1, max_col=2):
        row[0].alignment = STAMP_ALIGNMENT
        row[1].alignment = REPORT_ALIGNMENT


def excel_owner_file() -> str:
    """The owner file Excel keeps beside a workbook it has open: ~$name.xlsx."""
    folder, name = os.path.split(EXCEL_FILE_PATH)
    return os.path.join(folder, "~$" + name)


def workbook_is_open_in_excel() -> bool:
    """True while Excel holds the workbook open.

    This check matters far more than it looks.  The project normally lives on a
    Windows drive reached through WSL, and over that mount Excel's lock does not
    reach Python as a PermissionError - so a write *succeeds*, and then Excel
    saves its own older in-memory copy over the top minutes later and the row is
    simply gone.  Nothing in the app can see that happen after the fact, so the
    only safe move is to not write at all while Excel is in there: queue the
    report instead and merge it once the file is free, which is exactly what the
    pending queue already exists to do.
    """
    try:
        return os.path.exists(excel_owner_file())
    except OSError:
        return False


class ReportQueuedError(Exception):
    """The workbook was locked, so the report went to the pending queue."""

    def __init__(self, timestamp: str, original: Exception):
        super().__init__(str(original))
        self.timestamp = timestamp
        self.original = original


def _create_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Reports"
    ws.append(["Date-Time", "Task Report"])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 100
    _style_report_cells(ws)
    wb.save(EXCEL_FILE_PATH)


def _ensure_workbook():
    if not os.path.exists(EXCEL_FILE_PATH):
        _create_workbook()


# ---------------------------------------------------------------------------
# One day, one cell
# ---------------------------------------------------------------------------
#
# A day is a row.  Filing a report for a day that is already in the sheet grows
# that day's cell instead of starting another row: the new text goes a blank
# line under what is there, and the Date-Time moves on to the newer of the two
# stamps, so the row always says when the day's work was last added to.  Only a
# new day starts a new row.
#
# Every write goes through _write_one_report, which is what keeps the two
# routes into the sheet - a report filed now, and a report merged in from the
# pending queue later - from disagreeing about it.

# A blank line between one report and the next, matching the blank line
# compose_report_text already puts between two projects.
REPORT_JOIN = "\n\n"


def _row_date_key(value) -> str:
    """The day a Date-Time cell belongs to as YYYY-MM-DD, or "" if unreadable.

    Rows this app wrote are strings; rows typed straight into Excel come back
    as datetimes, and a cell nobody has touched since could be either.
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime(DATE_STORE_FORMAT)

    text = str(value).strip()
    if not text:
        return ""
    for fmt in (TIMESTAMP_FORMAT, "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime(DATE_STORE_FORMAT)
        except ValueError:
            continue
    # Anything else still begins with the date, so read just that.
    try:
        return datetime.strptime(text.split()[0], "%d/%m/%Y").strftime(
            DATE_STORE_FORMAT
        )
    except (ValueError, IndexError):
        return ""


def _parse_stamp(value):
    if isinstance(value, datetime):
        return value
    try:
        return datetime.strptime(str(value).strip(), TIMESTAMP_FORMAT)
    except (ValueError, TypeError):
        return None


def _later_stamp(existing, incoming: str) -> str:
    """The newer of two stamps for the same day.

    Normally the incoming one, since it was just written.  The comparison is
    here for the backlog case: a day's tasks ticked off at 09:00 and then more
    of them at 17:00 must not drag the row's clock time backwards.
    """
    was = _parse_stamp(existing)
    now = _parse_stamp(incoming)
    if was is None or now is None:
        return incoming
    return incoming if now >= was else was.strftime(TIMESTAMP_FORMAT)


def _find_day_row(ws, date_key: str) -> int:
    """The row this day already occupies, or 0 if it has none.

    Searched bottom-up, so a day left with several rows by an older version of
    this app grows the last of them rather than the first.
    """
    if not date_key:
        return 0
    for row in range(ws.max_row, 1, -1):
        if _row_date_key(ws.cell(row=row, column=1).value) == date_key:
            return row
    return 0


def _write_one_report(ws, timestamp: str, text: str) -> dict:
    """Put one report into the sheet, merging it into its day where there is one.

    Returns what was done, in enough detail to take it back out again - the
    filing Undo needs exactly that and nothing more.
    """
    row = _find_day_row(ws, _row_date_key(timestamp))
    if not row:
        ws.append([timestamp, text])
        return {"row": ws.max_row, "new": True, "result": text}

    previous_stamp = ws.cell(row=row, column=1).value
    previous_text = ws.cell(row=row, column=2).value
    existing = "" if previous_text is None else str(previous_text).rstrip()
    if not existing:
        merged = text
    elif len(existing) + len(REPORT_JOIN) + len(text) > EXCEL_MAX_CELL:
        # A second row for the day beats a report that will not fit in a cell.
        ws.append([timestamp, text])
        return {"row": ws.max_row, "new": True, "result": text}
    else:
        merged = existing + REPORT_JOIN + text

    ws.cell(
        row=row,
        column=1,
        value=_later_stamp(previous_stamp, timestamp),
    )
    ws.cell(row=row, column=2, value=merged)
    return {
        "row": row,
        "new": False,
        "result": merged,
        "previousStamp": _format_cell_timestamp(previous_stamp),
        "previousText": "" if previous_text is None else str(previous_text),
    }


def existing_report_days() -> dict:
    """{YYYY-MM-DD: characters already filed for that day}.

    What the filing preview needs to say whether a day will grow a row or
    start one, and to size the result against what a cell holds.  Queued
    reports count: they are going to merge into the same day themselves.
    """
    days = {}
    with _EXCEL_LOCK:
        if os.path.exists(EXCEL_FILE_PATH):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE_PATH, read_only=True)
                ws = wb.active
                first = True
                for row in ws.iter_rows(values_only=True):
                    if first:
                        first = False
                        continue
                    key = _row_date_key(row[0])
                    if not key:
                        continue
                    text = str(row[1]) if len(row) > 1 and row[1] is not None else ""
                    # Later rows win, because _find_day_row takes the last one.
                    days[key] = len(text.rstrip())
                wb.close()
            except Exception:
                days = {}

        for timestamp, report in _read_pending():
            key = _row_date_key(timestamp)
            if not key:
                continue
            if days.get(key):
                days[key] += len(REPORT_JOIN) + len(report)
            else:
                days[key] = len(report)
    return days


def _queue_pending(timestamp: str, text: str):
    with open(PENDING_FILE_PATH, "a", encoding="utf-8") as fh:
        fh.write(json.dumps({"timestamp": timestamp, "report": text}) + "\n")
        fh.flush()


def _read_pending() -> list:
    if not os.path.exists(PENDING_FILE_PATH):
        return []
    entries = []
    try:
        with open(PENDING_FILE_PATH, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                except ValueError:
                    continue
                report = str(data.get("report") or "").strip()
                if report:
                    entries.append((str(data.get("timestamp") or "").strip(), report))
    except OSError:
        return []
    return entries


def _drop_pending():
    try:
        os.remove(PENDING_FILE_PATH)
    except OSError:
        pass


def pending_report_count() -> int:
    with _EXCEL_LOCK:
        return len(_read_pending())


def flush_pending_reports() -> int:
    """Merge queued reports into the workbook.  Returns how many landed."""
    with _EXCEL_LOCK:
        entries = _read_pending()
        if not entries:
            return 0
        if workbook_is_open_in_excel():
            # Leave them queued; the next flush with Excel closed takes them.
            return 0
        try:
            _ensure_workbook()
            wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
            ws = wb.active
            for timestamp, report in entries:
                _write_one_report(ws, timestamp, report)
            _style_report_cells(ws)
            wb.save(EXCEL_FILE_PATH)
        except Exception:
            return 0
        _drop_pending()
        return len(entries)


def append_report_to_excel(
    report_text: str, timestamp: str = None, undo: list = None
) -> str:
    """File a report and return the timestamp that was written.

    A day already in the sheet is added to rather than repeated: see
    _write_one_report.  So this appends a row for a new day and grows the
    existing cell for a day already there, and the caller does not have to
    know which happened.

    `timestamp` is for reports that belong to a day other than today - the task
    board files a backlog under the date its tasks were listed against, not the
    date it happened to press the button.  Left out, it is now.

    Raises ReportQueuedError when the workbook cannot be written - the report
    is safely queued in that case rather than lost.

    `undo`, when given, receives a record of the change - but only when this
    write was the only thing that touched the sheet.  A write that also merged
    queued reports in is not offered for undo: unpicking one from the other is
    exactly the kind of cleverness that loses somebody's report.
    """
    if not report_text or not report_text.strip():
        raise ValueError("The report cannot be empty.")

    text = report_text.strip()
    timestamp = (timestamp or "").strip() or datetime.now().strftime(TIMESTAMP_FORMAT)

    with _EXCEL_LOCK:
        pending = _read_pending()
        if workbook_is_open_in_excel():
            # Writing now would appear to work and then be thrown away.
            _queue_pending(timestamp, text)
            raise ReportQueuedError(
                timestamp, OSError(f"{EXCEL_FILE_NAME} is open in Excel")
            )
        try:
            _ensure_workbook()
            wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
            ws = wb.active
            for queued_timestamp, queued_report in pending:
                _write_one_report(ws, queued_timestamp, queued_report)
            change = _write_one_report(ws, timestamp, text)
            _style_report_cells(ws)
            wb.save(EXCEL_FILE_PATH)
        except (PermissionError, OSError) as exc:
            _queue_pending(timestamp, text)
            raise ReportQueuedError(timestamp, exc) from exc

        if pending:
            _drop_pending()
        elif undo is not None:
            undo.append(change)

    return timestamp


def _format_cell_timestamp(value) -> str:
    """Render a Date-Time cell.

    Rows typed directly into Excel come back as real datetime objects, while
    rows written by this app are strings; normalise both to one format.
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime(TIMESTAMP_FORMAT)
    return str(value)


def load_reports_from_excel() -> list:
    with _EXCEL_LOCK:
        rows = []
        if os.path.exists(EXCEL_FILE_PATH):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE_PATH, read_only=True)
                ws = wb.active
                first = True
                for row in ws.iter_rows(values_only=True):
                    if first:
                        first = False
                        continue
                    dt_val = _format_cell_timestamp(row[0])
                    rpt_val = (
                        str(row[1]) if len(row) > 1 and row[1] is not None else ""
                    )
                    rows.append((dt_val, rpt_val))
                wb.close()
            except Exception:
                rows = []
        # Queued-but-not-yet-merged reports are real reports; show them too.
        rows.extend(_read_pending())
        return rows


def compact_excel():
    with _EXCEL_LOCK:
        if not os.path.exists(EXCEL_FILE_PATH):
            return
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        kept = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None:
                continue
            vals = [c for c in row if c is not None and str(c).strip()]
            if vals:
                kept.append(
                    (
                        _format_cell_timestamp(row[0]),
                        str(row[1]) if len(row) > 1 and row[1] is not None else "",
                    )
                )
        for r in range(ws.max_row, 1, -1):
            ws.delete_rows(r)
        for dt, rpt in kept:
            ws.append([dt, rpt])
        # The widths are deliberately left alone: this runs after every edit
        # and delete, and resetting them would undo whatever column sizing has
        # been set in Excel.
        _style_report_cells(ws)
        wb.save(EXCEL_FILE_PATH)


def _refuse_if_excel_has_it():
    """Rewriting the sheet under Excel loses the rewrite - say so instead."""
    if workbook_is_open_in_excel():
        raise PermissionError(
            f"{EXCEL_FILE_NAME} is open in Excel. Close it and try again - "
            "changing it now would be undone the next time Excel saves."
        )


def delete_report_from_excel(row_index: int):
    with _EXCEL_LOCK:
        _refuse_if_excel_has_it()
        flush_pending_reports()
        if not os.path.exists(EXCEL_FILE_PATH):
            return
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        excel_row = row_index + 2
        if excel_row > ws.max_row:
            wb.close()
            return
        ws.delete_rows(excel_row)
        wb.save(EXCEL_FILE_PATH)
        compact_excel()


def update_report_in_excel(row_index: int, new_datetime: str, new_text: str):
    with _EXCEL_LOCK:
        _refuse_if_excel_has_it()
        flush_pending_reports()
        if not os.path.exists(EXCEL_FILE_PATH):
            return
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        excel_row = row_index + 2
        if excel_row > ws.max_row:
            wb.close()
            return
        ws.cell(row=excel_row, column=1, value=new_datetime)
        ws.cell(row=excel_row, column=2, value=new_text.strip())
        _style_report_cells(ws)
        wb.save(EXCEL_FILE_PATH)
        compact_excel()


def load_sheet_reports() -> list:
    """The rows actually in the workbook, without the pending queue.

    The history panel shows queued reports separately - they cannot be edited
    or deleted, because they have no row yet - so it needs the two apart.
    """
    with _EXCEL_LOCK:
        rows = []
        if os.path.exists(EXCEL_FILE_PATH):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE_PATH, read_only=True)
                ws = wb.active
                first = True
                for row in ws.iter_rows(values_only=True):
                    if first:
                        first = False
                        continue
                    rows.append(
                        (
                            _format_cell_timestamp(row[0]),
                            str(row[1]) if len(row) > 1 and row[1] is not None else "",
                        )
                    )
                wb.close()
            except Exception:
                rows = []
        return rows


def list_pending_reports() -> list:
    with _EXCEL_LOCK:
        return _read_pending()


def restore_report_row(row_index: int, when: str, text: str):
    """Put a deleted row back where it was - the report-delete Undo."""
    text = str(text or "")
    if not text.strip():
        raise ValueError("Nothing to restore.")
    with _EXCEL_LOCK:
        _refuse_if_excel_has_it()
        flush_pending_reports()
        _ensure_workbook()
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        excel_row = max(2, min(int(row_index) + 2, ws.max_row + 1))
        if excel_row > ws.max_row:
            ws.append([when, text])
        else:
            ws.insert_rows(excel_row)
            ws.cell(row=excel_row, column=1, value=when)
            ws.cell(row=excel_row, column=2, value=text)
        _style_report_cells(ws)
        wb.save(EXCEL_FILE_PATH)


def undo_report_changes(changes: list):
    """Take back what _write_one_report did, newest first.

    Each change is checked against the sheet before it is reverted: if the
    cell no longer says what the write left in it, something else has touched
    the row since, and the undo is refused rather than guessing.
    """
    with _EXCEL_LOCK:
        _refuse_if_excel_has_it()
        if not os.path.exists(EXCEL_FILE_PATH):
            raise ValueError("The workbook is gone.")
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        for change in changes:
            row = int(change["row"])
            current = ws.cell(row=row, column=2).value
            if row > ws.max_row or str(current or "") != change["result"]:
                wb.close()
                raise ValueError(
                    "The workbook has changed since then, so the filing was "
                    "left as it is."
                )
        for change in sorted(changes, key=lambda item: item["row"], reverse=True):
            row = int(change["row"])
            if change["new"]:
                ws.delete_rows(row)
            else:
                ws.cell(row=row, column=1, value=change["previousStamp"])
                ws.cell(row=row, column=2, value=change["previousText"])
        _style_report_cells(ws)
        wb.save(EXCEL_FILE_PATH)


def plan_day_merge() -> list:
    """Which days in the sheet are still spread over more than one row.

    Nothing writes a second row for a day any more, so this only ever finds
    rows written before that was true - which is why merging them is a command
    to run once rather than something that happens on its own.
    """
    rows = load_reports_from_excel()
    order = []
    seen = {}
    for stamp, text in rows:
        key = _row_date_key(stamp)
        if not key:
            continue
        if key not in seen:
            seen[key] = []
            order.append(key)
        seen[key].append((stamp, text))

    return [
        {
            "date": key,
            "dateLabel": display_date(key),
            "rows": len(seen[key]),
            "chars": sum(len(t) for _, t in seen[key])
            + len(REPORT_JOIN) * (len(seen[key]) - 1),
        }
        for key in order
        if len(seen[key]) > 1
    ]


def merge_existing_days() -> dict:
    """Fold every day that has several rows into one row per day.

    The text is joined in the order the rows are already in, a blank line
    between each, and the day keeps its latest stamp.  A day whose rows will
    not fit in one cell is left as however many rows it needs.

    Rows whose Date-Time cannot be read as a date are left exactly where they
    are: there is no way to tell which day they belong to, and guessing would
    move somebody's report.

    This rewrites the sheet, so - like editing a report - it is refused while
    Excel has the workbook open, where the rewrite would be silently undone.
    """
    with _EXCEL_LOCK:
        _refuse_if_excel_has_it()
        flush_pending_reports()
        if not os.path.exists(EXCEL_FILE_PATH):
            return {"ok": False, "merged": 0, "rowsRemoved": 0}

        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active

        original = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None:
                continue
            stamp = row[0]
            text = str(row[1]) if len(row) > 1 and row[1] is not None else ""
            if not str(stamp or "").strip() and not text.strip():
                continue
            original.append((stamp, text.rstrip()))

        # One slot per day, in the order the days first appear, so the sheet
        # keeps the running order it already had.
        slots = []
        index = {}
        for stamp, text in original:
            key = _row_date_key(stamp)
            if not key:
                slots.append([None, [(stamp, text)]])
                continue
            if key in index:
                index[key][1].append((stamp, text))
            else:
                slot = [key, [(stamp, text)]]
                index[key] = slot
                slots.append(slot)

        merged_days = 0
        final = []
        for key, entries in slots:
            if key is None or len(entries) == 1:
                final.extend(
                    (_format_cell_timestamp(stamp), text) for stamp, text in entries
                )
                continue

            stamp = _format_cell_timestamp(entries[0][0])
            for other, _ in entries[1:]:
                stamp = _later_stamp(stamp, _format_cell_timestamp(other))

            # Filled a cell at a time, so an over-long day stays split rather
            # than losing the overflow.
            chunks = [""]
            for _, text in entries:
                if not text:
                    continue
                if not chunks[-1]:
                    chunks[-1] = text
                elif len(chunks[-1]) + len(REPORT_JOIN) + len(text) <= EXCEL_MAX_CELL:
                    chunks[-1] += REPORT_JOIN + text
                else:
                    chunks.append(text)

            if len(chunks) < len(entries):
                merged_days += 1
            final.extend((stamp, chunk) for chunk in chunks if chunk)

        removed = len(original) - len(final)
        if not merged_days:
            wb.close()
            return {"ok": True, "merged": 0, "rowsRemoved": 0}

        for row in range(ws.max_row, 1, -1):
            ws.delete_rows(row)
        for stamp, text in final:
            ws.append([stamp, text])
        # Widths left alone, for the same reason compact_excel leaves them.
        _style_report_cells(ws)
        wb.save(EXCEL_FILE_PATH)
        return {"ok": True, "merged": merged_days, "rowsRemoved": removed}


# ---------------------------------------------------------------------------
# The task board - what has to be done, and what has been done
# ---------------------------------------------------------------------------
#
# The board mirrors the way the task list is kept by hand: a day, the projects
# worked on that day in [square brackets], and the tasks under each one with a
# box to tick.  Ticking the box is the whole point - "file checked tasks" turns
# every ticked-but-not-yet-filed task into report rows, one row per day, with
# the tasks grouped under their project headers exactly as they were written.
#
# It is stored as JSON rather than as a second worksheet for one reason: the
# workbook cannot be written while Excel has it open, and ticking a box must
# never be the thing that fails.  Reports are the archive; the board is the
# working surface in front of it.

_BOARD_LOCK = threading.RLock()
BOARD_VERSION = 1
# Bumped on every write.  Both surfaces live in one process, so a counter is
# all that is needed for the browser to notice a change made in the terminal.
_BOARD_REVISION = 0
# Excel refuses a cell longer than this, so a composed report is capped just
# under it rather than being lost at save time.
EXCEL_MAX_CELL = 32000


def today_iso() -> str:
    return datetime.now().strftime(DATE_STORE_FORMAT)


def normalise_date(value) -> str:
    """Accept the ways a date gets typed and return one ISO date.

    `2026-08-21` comes from the date picker, `21.08.2026` from the task list,
    `21/08/2026` from the report timestamps.  Anything unreadable falls back to
    today rather than rejecting the task the user just wrote.
    """
    if isinstance(value, datetime):
        return value.strftime(DATE_STORE_FORMAT)
    text = str(value or "").strip()
    if not text:
        return today_iso()
    text = text.split()[0]
    for fmt in (DATE_STORE_FORMAT, DATE_DISPLAY_FORMAT, "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).strftime(DATE_STORE_FORMAT)
        except ValueError:
            continue
    return today_iso()


def display_date(date_iso: str) -> str:
    try:
        return datetime.strptime(date_iso, DATE_STORE_FORMAT).strftime(
            DATE_DISPLAY_FORMAT
        )
    except (ValueError, TypeError):
        return str(date_iso or "")


def _empty_board() -> dict:
    return {"version": BOARD_VERSION, "seq": 0, "tasks": [], "projects": []}


def _clean_project_names(raw) -> list:
    """Trim, cap and de-duplicate a list of project names, keeping its order."""
    names = []
    seen = set()
    for item in raw or []:
        name = str(item or "").strip()[:MAX_PROJECT_LENGTH]
        if name and name.lower() not in seen:
            seen.add(name.lower())
            names.append(name)
    return names


def _remember_project(board: dict, name: str):
    """Move a project name to the front of the remembered list.

    The list outlives the tasks that used the name, which is the point: the
    project dropdown has to keep offering "onex-academy" after that day's tasks
    have been filed and cleared off the board.
    """
    name = str(name or "").strip()[:MAX_PROJECT_LENGTH]
    if not name:
        return
    board["projects"] = [name] + [
        other for other in board["projects"] if other.lower() != name.lower()
    ]


def _clean_task(raw, seq_hint: int) -> dict:
    """Normalise one stored task, so a hand-edited file cannot break the UI."""
    if not isinstance(raw, dict):
        return None
    text = str(raw.get("text") or "").strip()
    if not text:
        return None
    filed_at = str(raw.get("filed_at") or "").strip() or None
    try:
        order = int(raw["order"]) if raw.get("order") is not None else None
    except (TypeError, ValueError):
        order = None
    return {
        "id": str(raw.get("id") or "").strip() or secrets.token_urlsafe(8),
        "seq": int(raw.get("seq") or seq_hint),
        # Where the task sits in its day once it has been dragged.  None means
        # "never moved", which sorts by seq - see _task_position.
        "order": order,
        "date": normalise_date(raw.get("date")),
        "project": str(raw.get("project") or "").strip()[:MAX_PROJECT_LENGTH],
        "text": text[:MAX_TASK_LENGTH],
        "done": bool(raw.get("done")),
        "created_at": str(raw.get("created_at") or "").strip(),
        "done_at": str(raw.get("done_at") or "").strip() or None,
        "filed_at": filed_at,
        # A board written before files existed simply has none.
        "files": _clean_file_list(raw.get("files")),
    }


def _clean_file_list(raw) -> list:
    files = []
    seen = set()
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        name = _safe_file_name(item.get("name"))
        if not name or name.lower() in seen:
            continue
        seen.add(name.lower())
        try:
            size = max(0, int(item.get("size") or 0))
        except (TypeError, ValueError):
            size = 0
        files.append(
            {"name": name, "size": size, "added_at": str(item.get("added_at") or "")}
        )
    return files


def _read_board() -> dict:
    """Read the board from disk.  Never raises - a broken file is set aside."""
    if not os.path.exists(TASKS_FILE_PATH):
        return _empty_board()
    try:
        with open(TASKS_FILE_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        # Losing the board silently would be worse than losing it loudly, so
        # the unreadable file is kept next to the new one.
        try:
            shutil.copyfile(TASKS_FILE_PATH, TASKS_FILE_PATH + ".bad")
        except OSError:
            pass
        return _empty_board()

    if not isinstance(data, dict):
        return _empty_board()

    tasks = []
    for position, raw in enumerate(data.get("tasks") or []):
        task = _clean_task(raw, position)
        if task is not None:
            tasks.append(task)

    seq = data.get("seq")
    try:
        seq = int(seq)
    except (TypeError, ValueError):
        seq = 0
    seq = max([seq] + [task["seq"] for task in tasks] + [0])

    if "projects" in data:
        projects = _clean_project_names(data.get("projects"))
    else:
        # A board written before names were remembered: seed the list from the
        # tasks on it, most recently written first.
        projects = _clean_project_names(
            task["project"]
            for task in sorted(tasks, key=lambda item: item["seq"], reverse=True)
        )

    return {
        "version": BOARD_VERSION,
        "seq": seq,
        "tasks": tasks,
        "projects": projects,
    }


def _write_board(board: dict):
    """Replace the board file atomically, so a crash cannot truncate it."""
    global _BOARD_REVISION
    board["version"] = BOARD_VERSION
    temp_path = TASKS_FILE_PATH + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as fh:
        json.dump(board, fh, indent=2, ensure_ascii=False)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(temp_path, TASKS_FILE_PATH)
    _BOARD_REVISION += 1


def board_revision() -> str:
    """A fingerprint that changes whenever the board does.

    The counter covers this process; the file's mtime and size cover the rest,
    so a task added from a second shell (./task-report -t "...") also makes an
    open browser page notice and reload.
    """
    with _BOARD_LOCK:
        try:
            info = os.stat(TASKS_FILE_PATH)
            disk = f"{info.st_mtime_ns}.{info.st_size}"
        except OSError:
            disk = "0.0"
        return f"{_BOARD_REVISION}.{disk}"


def _task_position(task: dict) -> tuple:
    """Where a task sorts inside its day.

    A dragged task has an `order`; one never moved falls back to its seq.
    Reordering a day numbers every task in it, so the two scales only meet
    for a task added after the drag - and seq is always the larger number, so
    a new task lands at the bottom, which is where it was typed.
    """
    order = task.get("order")
    return (task["seq"] if order is None else order, task["seq"])


def _sorted_tasks(tasks: list) -> list:
    """Oldest day first; within a day, the order shown on the board.

    This is the canonical order - it is the order the rows go into the
    workbook.  The board view turns the days round so today is on top.
    """
    return sorted(tasks, key=lambda task: (task["date"], _task_position(task)))


def load_tasks() -> list:
    with _BOARD_LOCK:
        return _sorted_tasks(_read_board()["tasks"])


def known_projects() -> list:
    """Project names to offer, most recently used first.

    This is what fills the project dropdown - the point is that the same
    project keeps the same spelling, because that spelling becomes the
    [bracketed] header in the filed report.
    """
    with _BOARD_LOCK:
        return list(_read_board()["projects"])


def forget_project(name: str) -> bool:
    """Stop offering a project name.  Tasks already using it are untouched.

    Typos are the reason this exists: a name typed once would otherwise sit in
    the dropdown forever.  Using the name again brings it back.
    """
    wanted = str(name or "").strip().lower()
    if not wanted:
        return False
    with _BOARD_LOCK:
        board = _read_board()
        kept = [item for item in board["projects"] if item.lower() != wanted]
        if len(kept) == len(board["projects"]):
            return False
        board["projects"] = kept
        _write_board(board)
        return True


def add_task(date, project: str, text: str) -> dict:
    text = str(text or "").strip()
    if not text:
        raise ValueError("The task cannot be empty.")
    if len(text) > MAX_TASK_LENGTH:
        raise ValueError(
            f"Task is too long ({len(text)} chars). The limit is {MAX_TASK_LENGTH}."
        )

    with _BOARD_LOCK:
        board = _read_board()
        board["seq"] += 1
        task = {
            "id": secrets.token_urlsafe(8),
            "seq": board["seq"],
            "date": normalise_date(date),
            "project": str(project or "").strip()[:MAX_PROJECT_LENGTH],
            "text": text,
            "done": False,
            "created_at": datetime.now().strftime(TIMESTAMP_FORMAT),
            "done_at": None,
            "filed_at": None,
            "order": None,
            "files": [],
        }
        board["tasks"].append(task)
        _remember_project(board, task["project"])
        _write_board(board)
        return task


def update_task(task_id: str, **changes) -> dict:
    """Change one task.  Only the keys passed in are touched."""
    if "text" in changes:
        text = str(changes["text"] or "").strip()
        if not text:
            raise ValueError("The task cannot be empty.")
        if len(text) > MAX_TASK_LENGTH:
            raise ValueError(
                f"Task is too long ({len(text)} chars). "
                f"The limit is {MAX_TASK_LENGTH}."
            )
        changes["text"] = text

    with _BOARD_LOCK:
        board = _read_board()
        for task in board["tasks"]:
            if task["id"] != task_id:
                continue
            if "text" in changes:
                task["text"] = changes["text"]
            if "project" in changes:
                task["project"] = (
                    str(changes["project"] or "").strip()[:MAX_PROJECT_LENGTH]
                )
                _remember_project(board, task["project"])
            if "date" in changes:
                new_date = normalise_date(changes["date"])
                if new_date != task["date"]:
                    # Moved to another day: it goes to the bottom of that day,
                    # not wherever its old position happens to fall.
                    others = [
                        _task_position(other)[0]
                        for other in board["tasks"]
                        if other["date"] == new_date and other is not task
                    ]
                    task["order"] = max(others) + 1 if others else None
                task["date"] = new_date
            if "done" in changes:
                done = bool(changes["done"])
                # Keep the original tick time when nothing actually changed.
                if done != task["done"]:
                    task["done_at"] = (
                        datetime.now().strftime(TIMESTAMP_FORMAT) if done else None
                    )
                task["done"] = done
            _write_board(board)
            return task
    raise KeyError("That task no longer exists.")


def delete_task(task_id: str) -> dict:
    """Take a task off the board.  Returns the task as it was, or None.

    Its files go to task_files/.trash/<id>/ rather than away, so Undo - which
    hands this same task back to restore_tasks - can bring them back.
    """
    with _BOARD_LOCK:
        board = _read_board()
        gone = [task for task in board["tasks"] if task["id"] == task_id]
        if not gone:
            return None
        board["tasks"] = [task for task in board["tasks"] if task["id"] != task_id]
        _write_board(board)
        _trash_task_folder(task_id)
        return gone[0]


# The last Clear Filed, kept for its Undo.  One batch is enough: the toast that
# offers the Undo is replaced by the next action anyway.
_LAST_CLEARED = []


def delete_filed_tasks() -> int:
    """Clear out everything already written into the workbook."""
    global _LAST_CLEARED
    with _BOARD_LOCK:
        board = _read_board()
        cleared = [task for task in board["tasks"] if task["filed_at"]]
        if cleared:
            board["tasks"] = [task for task in board["tasks"] if not task["filed_at"]]
            _write_board(board)
            for task in cleared:
                _trash_task_folder(task["id"])
        _LAST_CLEARED = cleared
        return len(cleared)


def restore_tasks(tasks: list) -> int:
    """Put deleted tasks back with their original id, seq and files.

    A task whose id is already on the board is skipped, so pressing Undo
    twice cannot duplicate anything.
    """
    with _BOARD_LOCK:
        board = _read_board()
        present = {task["id"] for task in board["tasks"]}
        restored = 0
        for position, raw in enumerate(tasks or []):
            task = _clean_task(raw, board["seq"] + 1 + position)
            if task is None or task["id"] in present:
                continue
            board["tasks"].append(task)
            present.add(task["id"])
            board["seq"] = max(board["seq"], task["seq"])
            _remember_project(board, task["project"])
            _untrash_task_folder(task["id"])
            restored += 1
        if restored:
            _write_board(board)
        return restored


def restore_cleared_tasks() -> int:
    global _LAST_CLEARED
    with _BOARD_LOCK:
        restored = restore_tasks(_LAST_CLEARED)
        _LAST_CLEARED = []
        return restored


def reorder_tasks(date, ids: list) -> int:
    """Number one day's tasks in the order given - the drag-to-reorder.

    Ids not on that day are ignored; tasks on that day the list leaves out
    keep their relative order and go after the ones it names.
    """
    date_iso = normalise_date(date)
    wanted = [str(item) for item in (ids or [])]
    with _BOARD_LOCK:
        board = _read_board()
        day = [task for task in board["tasks"] if task["date"] == date_iso]
        by_id = {task["id"]: task for task in day}
        named = [by_id[item] for item in dict.fromkeys(wanted) if item in by_id]
        rest = sorted(
            (task for task in day if task["id"] not in set(wanted)),
            key=_task_position,
        )
        for position, task in enumerate(named + rest):
            task["order"] = position
        if day:
            _write_board(board)
        return len(named)


# --------------------------------------------------------------- task files

_WINDOWS_RESERVED = {
    "con", "prn", "aux", "nul",
    *(f"com{n}" for n in range(1, 10)),
    *(f"lpt{n}" for n in range(1, 10)),
}


def _safe_file_name(name) -> str:
    """A file name that is safe on Windows and cannot climb out of its folder."""
    name = os.path.basename(str(name or "").replace("\\", "/")).strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    if not name or name in (".", ".."):
        return ""
    stem = name.split(".")[0].lower()
    if stem in _WINDOWS_RESERVED:
        name = "_" + name
    return name[:180]


def _safe_task_id(task_id) -> str:
    task_id = str(task_id or "").strip()
    # Ids are token_urlsafe - letters, digits, - and _ - and nothing else may
    # become a folder name.
    return task_id if re.fullmatch(r"[A-Za-z0-9_-]{1,64}", task_id) else ""


def task_files_folder(task_id: str) -> str:
    return os.path.join(TASK_FILES_DIR, _safe_task_id(task_id))


def task_file_path(task_id: str, name: str) -> str:
    """The file on disk, or "" when the id or name is not one of ours."""
    task_id = _safe_task_id(task_id)
    name = _safe_file_name(name)
    if not task_id or not name:
        return ""
    return os.path.join(TASK_FILES_DIR, task_id, name)


def _unique_name(folder: str, name: str, taken) -> str:
    """name, or "name (2).ext" and so on until it is free."""
    lowered = {item.lower() for item in taken}
    stem, ext = os.path.splitext(name)
    candidate = name
    counter = 2
    while candidate.lower() in lowered or os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{stem} ({counter}){ext}"
        counter += 1
    return candidate


def _trash_task_folder(task_id: str):
    source = task_files_folder(task_id)
    if not _safe_task_id(task_id) or not os.path.isdir(source):
        return
    target = os.path.join(TASK_FILES_TRASH, _safe_task_id(task_id))
    try:
        os.makedirs(TASK_FILES_TRASH, exist_ok=True)
        if os.path.exists(target):
            shutil.rmtree(target, ignore_errors=True)
        shutil.move(source, target)
    except OSError as exc:
        print(f"  [!] Could not move the files of task {task_id} to the trash: {exc}")


def _untrash_task_folder(task_id: str):
    source = os.path.join(TASK_FILES_TRASH, _safe_task_id(task_id))
    if not _safe_task_id(task_id) or not os.path.isdir(source):
        return
    target = task_files_folder(task_id)
    try:
        if os.path.isdir(target):
            # Files were added again in between: merge, keeping both.
            for name in os.listdir(source):
                final = _unique_name(target, name, [])
                shutil.move(os.path.join(source, name), os.path.join(target, final))
            shutil.rmtree(source, ignore_errors=True)
        else:
            os.makedirs(TASK_FILES_DIR, exist_ok=True)
            shutil.move(source, target)
    except OSError as exc:
        print(f"  [!] Could not bring back the files of task {task_id}: {exc}")


def add_task_file(task_id: str, name: str, data: bytes) -> dict:
    """Store a file with a task.  Returns its entry ({name, size, added_at})."""
    if len(data) > MAX_TASK_FILE_BYTES:
        raise ValueError(
            f"That file is {len(data) / 1048576:.1f} MB. Task files are limited "
            f"to {MAX_TASK_FILE_BYTES // 1048576} MB each."
        )
    clean = _safe_file_name(name)
    if not clean:
        raise ValueError("That file has no usable name.")
    with _BOARD_LOCK:
        board = _read_board()
        task = next((item for item in board["tasks"] if item["id"] == task_id), None)
        if task is None or not _safe_task_id(task_id):
            raise KeyError("That task no longer exists.")
        folder = task_files_folder(task_id)
        os.makedirs(folder, exist_ok=True)
        final = _unique_name(folder, clean, [entry["name"] for entry in task["files"]])
        temp_path = os.path.join(folder, "." + final + ".part")
        with open(temp_path, "wb") as handle:
            handle.write(data)
        os.replace(temp_path, os.path.join(folder, final))
        entry = {
            "name": final,
            "size": len(data),
            "added_at": datetime.now().strftime(TIMESTAMP_FORMAT),
        }
        task["files"].append(entry)
        _write_board(board)
        return entry


def remove_task_file(task_id: str, name: str) -> bool:
    """Take a file off a task, into task_files/.trash/<id>/."""
    with _BOARD_LOCK:
        board = _read_board()
        task = next((item for item in board["tasks"] if item["id"] == task_id), None)
        if task is None:
            raise KeyError("That task no longer exists.")
        kept = [entry for entry in task["files"] if entry["name"] != name]
        if len(kept) == len(task["files"]):
            return False
        task["files"] = kept
        _write_board(board)

        path = task_file_path(task_id, name)
        if path and os.path.exists(path):
            trash = os.path.join(TASK_FILES_TRASH, _safe_task_id(task_id))
            try:
                os.makedirs(trash, exist_ok=True)
                target = os.path.join(trash, _unique_name(trash, os.path.basename(path), []))
                shutil.move(path, target)
            except OSError as exc:
                print(f"  [!] Could not move {path} to the trash: {exc}")
        return True


def unfiled_checked_tasks(tasks: list = None) -> list:
    """The tasks that pressing "file checked tasks" would actually write."""
    if tasks is None:
        tasks = load_tasks()
    return [task for task in tasks if task["done"] and not task["filed_at"]]


def board_counts() -> dict:
    tasks = load_tasks()
    return {
        "total": len(tasks),
        "done": sum(1 for task in tasks if task["done"]),
        "open": sum(1 for task in tasks if not task["done"]),
        "filed": sum(1 for task in tasks if task["filed_at"]),
        "ready": len(unfiled_checked_tasks(tasks)),
    }


def group_day_projects(tasks: list) -> list:
    """One day's tasks as [(project, [tasks])].

    Un-bracketed tasks come first - a task with no project has no header to sit
    under - then one block per project, in the order the projects were first
    written that day.  Everything that renders a day goes through here, so the
    board and the report it files always agree on the order.
    """
    order = []
    grouped = {}
    for task in tasks:
        name = task["project"]
        if name not in grouped:
            grouped[name] = []
            order.append(name)
        grouped[name].append(task)

    # "" sorts to the front on purpose.
    first_seen = {name: position for position, name in enumerate(order)}
    order.sort(key=lambda name: (name != "", first_seen[name]))
    return [(name, grouped[name]) for name in order]


def group_by_day(tasks: list) -> list:
    """Tasks as [(date, [(project, [tasks])])], oldest day first."""
    order = []
    grouped = {}
    for task in tasks:
        if task["date"] not in grouped:
            grouped[task["date"]] = []
            order.append(task["date"])
        grouped[task["date"]].append(task)
    return [(date, group_day_projects(grouped[date])) for date in sorted(order)]


# One task, one line, marked so a run of them reads as a list rather than as a
# paragraph that happens to have line breaks in it.
TASK_BULLET = "• "
TASK_INDENT = " " * len(TASK_BULLET)


def _task_lines(text: str) -> list:
    """One task as a bulleted line, with any lines of its own lined up under it."""
    own = [line.strip() for line in str(text).strip().splitlines()]
    own = [line for line in own if line]
    if not own:
        return []
    return [TASK_BULLET + own[0]] + [TASK_INDENT + line for line in own[1:]]


def compose_report_text(tasks: list) -> str:
    """Render one day's tasks the way the task list writes them.

    A project heading, its tasks bulleted underneath it, and a blank line
    before the next project - so several projects in one day stay legible
    instead of running together.
    """
    blocks = []
    for name, group in group_day_projects(tasks):
        lines = ["[" + name + "]"] if name else []
        for task in group:
            lines.extend(_task_lines(task["text"]))
        if lines:
            blocks.append("\n".join(lines))
    return "\n\n".join(blocks).strip()


def _bucket_timestamp(date_iso: str) -> str:
    """The Date-Time a day's report is filed under.

    The day comes from the task list, the clock time from now - which for
    today's tasks is simply the current time, and for a backlog keeps the row
    on the day the work actually happened.
    """
    now = datetime.now()
    try:
        day = datetime.strptime(date_iso, DATE_STORE_FORMAT)
    except (ValueError, TypeError):
        return now.strftime(TIMESTAMP_FORMAT)
    return day.replace(
        hour=now.hour, minute=now.minute, second=now.second
    ).strftime(TIMESTAMP_FORMAT)


def preview_filing() -> list:
    """What filing would write, without writing it.

    One entry per day, oldest first, so the days land in the workbook in the
    order they happened.  An entry is not necessarily a new row: a day already
    in the sheet is added to its existing cell, which `appendsToExisting` says.
    """
    ready = unfiled_checked_tasks()
    if not ready:
        return []

    by_date = {}
    for task in ready:
        by_date.setdefault(task["date"], []).append(task)

    already = existing_report_days()

    groups = []
    for date_iso in sorted(by_date):
        day_tasks = sorted(by_date[date_iso], key=_task_position)
        text = compose_report_text(day_tasks)

        # A day already in the sheet grows its cell rather than adding a row -
        # unless the two together would not fit in one, which is the one case
        # _write_one_report starts a second row for.  The preview has to say
        # whichever of those is actually going to happen.
        existing_length = already.get(date_iso, 0)
        joined = existing_length + len(REPORT_JOIN) + len(text)
        appends = bool(existing_length) and joined <= EXCEL_MAX_CELL

        groups.append(
            {
                "date": date_iso,
                "dateLabel": display_date(date_iso),
                "timestamp": _bucket_timestamp(date_iso),
                "text": text,
                "length": len(text),
                "existingLength": existing_length,
                "totalLength": joined if appends else len(text),
                "appendsToExisting": appends,
                "taskCount": len(day_tasks),
                "taskIds": [task["id"] for task in day_tasks],
                "projects": [
                    name
                    for name in dict.fromkeys(task["project"] for task in day_tasks)
                    if name
                ],
            }
        )
    return groups


def file_checked_tasks(session=None, origin: str = "board") -> dict:
    """Write every ticked-but-unfiled task into the workbook.

    Returns what happened: the rows written, how many tasks were marked filed,
    and whether any row had to be queued because Excel held the workbook.  A
    queued row still counts as filed - it is in the pending file and merges
    itself in later, so re-filing it would duplicate it.
    """
    # Held across the whole operation - reading what to file, writing it, and
    # marking it filed have to be one step, or two surfaces filing at the same
    # instant could both write the same day.  _BOARD_LOCK is an RLock and is
    # always taken before _EXCEL_LOCK, never the other way round.
    with _BOARD_LOCK:
        return _file_checked_tasks_locked(session, origin)


def _file_checked_tasks_locked(session, origin: str) -> dict:
    groups = preview_filing()
    if not groups:
        return {"ok": False, "message": "No checked tasks are waiting to be filed."}

    written = []
    failed = []
    filed_ids = []
    undo_changes = []
    undoable = True

    for group in groups:
        text = group["text"]
        if len(text) > EXCEL_MAX_CELL:
            failed.append(
                {
                    "dateLabel": group["dateLabel"],
                    "message": (
                        f"{group['dateLabel']} is {len(text)} characters, which "
                        f"is over what one cell holds ({EXCEL_MAX_CELL}). Split "
                        "that day or file fewer tasks at once."
                    ),
                }
            )
            continue

        queued = False
        change = []
        try:
            timestamp = append_report_to_excel(
                text, timestamp=group["timestamp"], undo=change
            )
        except ReportQueuedError as exc:
            timestamp = exc.timestamp
            queued = True
        except Exception as exc:
            failed.append({"dateLabel": group["dateLabel"], "message": str(exc)})
            continue
        if change:
            undo_changes.extend(change)
        else:
            undoable = False

        if session is not None:
            session.record_save(origin, timestamp, queued=queued)

        written.append(
            {
                "dateLabel": group["dateLabel"],
                "timestamp": timestamp,
                "taskCount": group["taskCount"],
                "queued": queued,
            }
        )
        filed_ids.extend(group["taskIds"])

    global _LAST_FILING
    _LAST_FILING = None
    if filed_ids:
        stamp = datetime.now().strftime(TIMESTAMP_FORMAT)
        board = _read_board()
        wanted = set(filed_ids)
        for task in board["tasks"]:
            if task["id"] in wanted:
                task["filed_at"] = stamp
        _write_board(board)
        if undoable and undo_changes:
            _LAST_FILING = {
                "at": time.monotonic(),
                "stamp": stamp,
                "ids": list(filed_ids),
                "changes": undo_changes,
            }

    return {
        "ok": bool(written),
        "written": written,
        "failed": failed,
        "filedCount": len(filed_ids),
        "queued": any(row["queued"] for row in written),
        # Only a write that went straight into the sheet, alone, can be taken
        # back - see append_report_to_excel.
        "undoable": _LAST_FILING is not None,
    }


# The last filing, for its Undo.  Offered for a few seconds only: the toast is
# up for 6, and the server allows a little slack for a slow click.
_LAST_FILING = None
FILING_UNDO_SECONDS = 10.0


def undo_last_filing() -> dict:
    """Take the last filing back out of the workbook and un-mark its tasks."""
    global _LAST_FILING
    with _BOARD_LOCK:
        filing = _LAST_FILING
        if not filing or time.monotonic() - filing["at"] > FILING_UNDO_SECONDS:
            _LAST_FILING = None
            return {"ok": False, "message": "It is too late to undo that filing."}
        try:
            undo_report_changes(filing["changes"])
        except (PermissionError, ValueError) as exc:
            _LAST_FILING = None
            return {"ok": False, "message": str(exc)}
        board = _read_board()
        wanted = set(filing["ids"])
        count = 0
        for task in board["tasks"]:
            if task["id"] in wanted and task["filed_at"] == filing["stamp"]:
                task["filed_at"] = None
                count += 1
        _write_board(board)
        _LAST_FILING = None
        return {"ok": True, "count": count}


# ---------------------------------------------------------------------------
# Session coordination between the two surfaces
# ---------------------------------------------------------------------------


class ReporterSession:
    """Shared state linking the GUI window and the terminal console.

    Either surface can call request_shutdown(); the other notices and stops,
    which is what makes "close the window -> the terminal closes too" (and the
    reverse) work.
    """

    def __init__(self):
        self._shutdown = threading.Event()
        self._lock = threading.Lock()
        self._reason = None
        self._events = deque()
        self.saved_count = 0
        self.cli_active = False

    @property
    def reason(self):
        with self._lock:
            return self._reason

    def is_shutting_down(self) -> bool:
        return self._shutdown.is_set()

    def wait_for_shutdown(self, timeout=None) -> bool:
        return self._shutdown.wait(timeout)

    def request_shutdown(self, reason: str):
        with self._lock:
            if self._reason is None:
                self._reason = reason
        self._shutdown.set()

    def record_save(self, origin: str, timestamp: str, queued: bool = False):
        with self._lock:
            self.saved_count += 1
            self._events.append(
                {"origin": origin, "timestamp": timestamp, "queued": queued}
            )

    def drain_events(self) -> list:
        with self._lock:
            events = list(self._events)
            self._events.clear()
        return events


DARK_QSS = """
QWidget {
    background-color: #0F1117;
    color: #E2E8F0;
    font-family: "Segoe UI", "Roboto", "Helvetica Neue", sans-serif;
    font-size: 14px;
}

QFrame#headerCard {
    background-color: #161B27;
    border-radius: 14px;
    border: 1px solid #1E2640;
}

QFrame#separator {
    background-color: #1E2640;
    max-height: 1px;
    min-height: 1px;
}

QFrame#contentCard {
    background-color: #161B27;
    border-radius: 14px;
    border: 1px solid #1E2640;
}

QLabel#titleLabel {
    font-size: 28px;
    font-weight: 700;
    color: #E8F0FE;
    background: transparent;
}

QLabel#clockLabel {
    font-size: 13px;
    color: #5B6EA6;
    background: transparent;
}

QPushButton#helpBtn, QPushButton#historyBtn {
    background-color: #1E2640;
    color: #7B90D4;
    border: 1px solid #2D3860;
    border-radius: 14px;
    font-size: 14px;
    font-weight: 700;
    min-width:  28px;
    max-width:  28px;
    min-height: 28px;
    max-height: 28px;
    padding: 0px;
}
QPushButton#helpBtn:hover, QPushButton#historyBtn:hover {
    background-color: #2D3860;
    color: #A8BFFF;
}
QPushButton#helpBtn:pressed, QPushButton#historyBtn:pressed {
    background-color: #3B4A80;
}

QLabel#headingLabel {
    font-size: 20px;
    font-weight: 700;
    color: #E8F0FE;
    background: transparent;
}

QPlainTextEdit#editor {
    background-color: #0D1020;
    color: #CBD5E1;
    border: 2px solid #1E2640;
    border-radius: 10px;
    font-size: 15px;
    selection-background-color: #2563EB;
    selection-color: #FFFFFF;
    padding: 4px;
}
QPlainTextEdit#editor:focus {
    border: 2px solid #3B82F6;
}

QScrollBar:vertical {
    background: #0D1020;
    width: 10px;
    margin: 0;
    border-radius: 5px;
}
QScrollBar::handle:vertical {
    background: #2D3860;
    border-radius: 5px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover { background: #3B4A80; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }

QLabel#statusLabel {
    font-size: 13px;
    font-weight: 600;
    color: #3B82F6;
    background: transparent;
}

QProgressBar#progressBar {
    background-color: #1E2640;
    border: none;
    border-radius: 4px;
    min-height: 8px;
    max-height: 8px;
}
QProgressBar#progressBar::chunk {
    background-color: #3B82F6;
    border-radius: 4px;
}

QProgressBar#progressBar[danger="true"]::chunk {
    background-color: #EF4444;
    border-radius: 4px;
}

QLabel#counterLabel {
    font-size: 13px;
    color: #5B6EA6;
    background: transparent;
}
QLabel#counterLabel[danger="true"] { color: #EF4444; }

QPushButton#saveBtn {
    background-color: qlineargradient(
        x1:0, y1:0, x2:0, y2:1, stop:0 #3B82F6, stop:1 #2563EB
    );
    color: #FFFFFF;
    border: none;
    border-radius: 10px;
    font-size: 14px;
    font-weight: 700;
    min-width:  156px;
    max-width:  156px;
    min-height: 40px;
    max-height: 40px;
    padding: 0 20px;
}
QPushButton#saveBtn:hover {
    background-color: qlineargradient(
        x1:0, y1:0, x2:0, y2:1, stop:0 #60A5FA, stop:1 #3B82F6
    );
}
QPushButton#saveBtn:pressed  { background-color: #1D4ED8; }
QPushButton#saveBtn:disabled { background-color: #1E2640; color: #3D4F7A; }

QMenu {
    background-color: #1A1F33;
    border: 1px solid #2D3860;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item          { padding: 7px 22px; border-radius: 5px; }
QMenu::item:selected { background-color: #2563EB; color: #FFFFFF; }
QMenu::separator     { height: 1px; background: #2D3860; margin: 4px 8px; }

QDialog#helpDialog                { background-color: #161B27; border-radius: 12px; }
QDialog#helpDialog QLabel         { background: transparent; }
QDialog#historyDialog             { background-color: #161B27; border-radius: 12px; }
QDialog#historyDialog QLabel      { background: transparent; }

QPushButton#closeBtn {
    background-color: #1E2640;
    color: #7B90D4;
    border: 1px solid #2D3860;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    min-height: 34px;
    padding: 0 24px;
}
QPushButton#closeBtn:hover { background-color: #2D3860; color: #A8BFFF; }

QTableWidget#historyTable {
    background-color: #0D1020;
    alternate-background-color: #131728;
    color: #CBD5E1;
    border: 1px solid #1E2640;
    border-radius: 8px;
    gridline-color: #1E2640;
    font-size: 13px;
    selection-background-color: #2563EB;
    selection-color: #FFFFFF;
}
QTableWidget#historyTable QHeaderView::section {
    background-color: #161B27;
    color: #7B90D4;
    border: none;
    border-bottom: 2px solid #2D3860;
    font-size: 13px;
    font-weight: 700;
    padding: 8px 12px;
}
QTableWidget#historyTable QTableCornerButton::section {
    background-color: #161B27;
    border: none;
}

QLabel#historyCountLabel {
    font-size: 12px;
    color: #5B6EA6;
    background: transparent;
}

QPushButton#editBtn {
    background-color: #1E2640;
    color: #7B90D4;
    border: 1px solid #2D3860;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    min-height: 28px;
    padding: 0 14px;
}
QPushButton#editBtn:hover { background-color: #2D3860; color: #A8BFFF; }
QPushButton#editBtn:pressed { background-color: #3B4A80; }

QPushButton#deleteBtn {
    background-color: #2A1520;
    color: #F87171;
    border: 1px solid #5B2030;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    min-height: 28px;
    padding: 0 14px;
}
QPushButton#deleteBtn:hover { background-color: #3D1A2A; color: #FCA5A5; }
QPushButton#deleteBtn:pressed { background-color: #4A1F30; }

QDialog#editDialog { background-color: #161B27; border-radius: 12px; }
QDialog#editDialog QLabel { background: transparent; }

QLineEdit#editField, QPlainTextEdit#editTextField {
    background-color: #0D1020;
    color: #CBD5E1;
    border: 2px solid #1E2640;
    border-radius: 8px;
    font-size: 14px;
    padding: 6px 10px;
    selection-background-color: #2563EB;
    selection-color: #FFFFFF;
}
QLineEdit#editField:focus, QPlainTextEdit#editTextField:focus {
    border: 2px solid #3B82F6;
}

QPushButton#saveEditBtn {
    background-color: qlineargradient(
        x1:0, y1:0, x2:0, y2:1, stop:0 #3B82F6, stop:1 #2563EB
    );
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 700;
    min-height: 34px;
    padding: 0 24px;
}
QPushButton#saveEditBtn:hover {
    background-color: qlineargradient(
        x1:0, y1:0, x2:0, y2:1, stop:0 #60A5FA, stop:1 #3B82F6
    );
}
QPushButton#saveEditBtn:pressed { background-color: #1D4ED8; }
"""


class ShortcutsDialog(QDialog):
    SHORTCUTS = [
        ("Ctrl + Enter", "Save the report"),
        ("Ctrl + A", "Select all text"),
        ("Ctrl + Z", "Undo"),
        ("Ctrl + Y  /  Ctrl + Shift + Z", "Redo"),
        ("Ctrl + Backspace", "Delete previous word"),
        ("Ctrl + Delete", "Delete next word"),
        ("Ctrl + C", "Copy selection"),
        ("Ctrl + X", "Cut selection"),
        ("Ctrl + V", "Paste"),
        ("Right-click", "Open context menu"),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("helpDialog")
        self.setWindowTitle("Keyboard Shortcuts")
        self.setModal(True)
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 20)
        layout.setSpacing(16)

        title = QLabel("Keyboard Shortcuts")
        title.setStyleSheet("font-size: 17px; font-weight: 700; color: #E8F0FE;")
        layout.addWidget(title)

        rows_html = ""
        for i, (keys, desc) in enumerate(self.SHORTCUTS):
            bg = "#0F1117" if i % 2 == 0 else "#1A1F33"
            rows_html += (
                f'<tr style="background-color:{bg};">'
                f'<td style="padding:7px 14px 7px 10px;">'
                f'<span style="background-color:#1E2640; color:#A8BFFF;'
                f" font-family:monospace; font-size:12px; padding:2px 8px;"
                f' border-radius:4px; white-space:nowrap;">{keys}</span>'
                f"</td>"
                f'<td style="padding:7px 4px 7px 8px; color:#CBD5E1;'
                f' font-size:13px;">{desc}</td>'
                f"</tr>"
            )
        table_html = (
            '<table style="border-collapse:collapse; width:100%;">'
            + rows_html
            + "</table>"
        )
        table_label = QLabel(table_html)
        table_label.setTextFormat(Qt.TextFormat.RichText)
        table_label.setWordWrap(False)
        layout.addWidget(table_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setObjectName("closeBtn")
        close_btn.clicked.connect(self.accept)
        QShortcut(QKeySequence("Escape"), self, activated=self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)


class EditReportDialog(QDialog):
    def __init__(self, row_index: int, current_dt: str, current_text: str, parent=None):
        super().__init__(parent)
        self.setObjectName("editDialog")
        self.setWindowTitle("Edit Report")
        self.setModal(True)
        self.setMinimumWidth(520)
        self.resize(580, 400)
        self._row_index = row_index
        self._saved = False

        from PySide6.QtWidgets import QLineEdit

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 16)
        layout.setSpacing(12)

        title = QLabel("Edit Report")
        title.setStyleSheet("font-size: 17px; font-weight: 700; color: #E8F0FE;")
        layout.addWidget(title)

        dt_label = QLabel("Date-Time")
        dt_label.setStyleSheet("font-size: 12px; font-weight: 600; color: #5B6EA6;")
        layout.addWidget(dt_label)

        self.dt_field = QLineEdit(current_dt)
        self.dt_field.setObjectName("editField")
        layout.addWidget(self.dt_field)

        rpt_label = QLabel("Task Report")
        rpt_label.setStyleSheet("font-size: 12px; font-weight: 600; color: #5B6EA6;")
        layout.addWidget(rpt_label)

        self.text_field = QPlainTextEdit()
        self.text_field.setObjectName("editTextField")
        self.text_field.setPlainText(current_text)
        self.text_field.document().setDocumentMargin(8)
        layout.addWidget(self.text_field, stretch=1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("closeBtn")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        save_btn = QPushButton("Save Changes")
        save_btn.setObjectName("saveEditBtn")
        save_btn.clicked.connect(self._do_save)
        btn_row.addWidget(save_btn)

        layout.addLayout(btn_row)

        QShortcut(QKeySequence("Escape"), self, activated=self.reject)

    def _do_save(self):
        new_dt = self.dt_field.text().strip()
        new_text = self.text_field.toPlainText().strip()
        if not new_text:
            QMessageBox.warning(self, "Warning", "Report text cannot be empty.")
            return
        if len(new_text) > MAX_REPORT_LENGTH:
            QMessageBox.warning(
                self,
                "Warning",
                f"Report is too long ({len(new_text)} chars).\n"
                f"Limit is {MAX_REPORT_LENGTH} characters.",
            )
            return
        try:
            update_report_in_excel(self._row_index, new_dt, new_text)
            self._saved = True
            self.accept()
        except PermissionError:
            QMessageBox.critical(
                self,
                "Permission Error",
                f"Could not save to:\n{EXCEL_FILE_PATH}\n\n"
                "Is the Excel file currently open? Please close it and try again.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")

    @property
    def was_saved(self) -> bool:
        return self._saved


class ReportHistoryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("historyDialog")
        self.setWindowTitle("Report History")
        self.setModal(True)
        self.setMinimumSize(780, 480)
        self.resize(900, 580)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 16)
        layout.setSpacing(12)

        header_row = QHBoxLayout()
        title = QLabel("Previous Reports")
        title.setStyleSheet("font-size: 17px; font-weight: 700; color: #E8F0FE;")
        header_row.addWidget(title)
        header_row.addStretch()
        self.count_label = QLabel("")
        self.count_label.setObjectName("historyCountLabel")
        header_row.addWidget(self.count_label)
        layout.addLayout(header_row)

        self.table = QTableWidget()
        self.table.setObjectName("historyTable")
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["#", "Date-Time", "Task Report", "", ""])
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setWordWrap(True)

        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        h.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        h.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        h.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(1, 160)
        self.table.setColumnWidth(3, 70)
        self.table.setColumnWidth(4, 76)

        layout.addWidget(self.table, stretch=1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setObjectName("closeBtn")
        close_btn.clicked.connect(self.accept)
        QShortcut(QKeySequence("Escape"), self, activated=self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self._load_data()

    def _load_data(self):
        self._rows = load_reports_from_excel()
        display = list(self._rows)
        display.reverse()
        self.table.setRowCount(len(display))
        total = len(display)
        self.count_label.setText(f"{total} report{'s' if total != 1 else ''}")

        for i, (dt, rpt) in enumerate(display):
            original_index = total - 1 - i

            num_item = QTableWidgetItem(str(total - i))
            num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 0, num_item)

            dt_item = QTableWidgetItem(dt)
            dt_item.setTextAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            self.table.setItem(i, 1, dt_item)

            rpt_item = QTableWidgetItem(rpt)
            rpt_item.setToolTip(rpt)
            self.table.setItem(i, 2, rpt_item)

            edit_btn = QPushButton("Edit")
            edit_btn.setObjectName("editBtn")
            edit_btn.setProperty("_orig_idx", original_index)
            edit_btn.clicked.connect(self._on_edit_clicked)
            self.table.setCellWidget(i, 3, edit_btn)

            del_btn = QPushButton("Delete")
            del_btn.setObjectName("deleteBtn")
            del_btn.setProperty("_orig_idx", original_index)
            del_btn.clicked.connect(self._on_delete_clicked)
            self.table.setCellWidget(i, 4, del_btn)

        self.table.resizeRowsToContents()

    def _on_edit_clicked(self):
        btn = self.sender()
        if btn is None:
            return
        idx = btn.property("_orig_idx")
        if idx is None or idx < 0 or idx >= len(self._rows):
            return
        dt, rpt = self._rows[idx]
        dlg = EditReportDialog(idx, dt, rpt, parent=self)
        dlg.exec()
        if dlg.was_saved:
            self._load_data()

    def _on_delete_clicked(self):
        btn = self.sender()
        if btn is None:
            return
        idx = btn.property("_orig_idx")
        if idx is None or idx < 0 or idx >= len(self._rows):
            return
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete report #{idx + 1}?\n\nThis action cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            delete_report_from_excel(idx)
            self._load_data()
        except PermissionError:
            QMessageBox.critical(
                self,
                "Permission Error",
                f"Could not save to:\n{EXCEL_FILE_PATH}\n\n"
                "Is the Excel file currently open? Please close it and try again.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")


class TaskReporterApp(QMainWindow):
    def __init__(self, session: "ReporterSession" = None):
        super().__init__()
        self._session = session
        self.setWindowTitle("Task Reporter")
        self.resize(1000, 720)
        self.setMinimumSize(920, 680)

        self._check_or_create_excel()
        self._build_ui()

        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._tick_clock)
        self._clock_timer.start(1000)
        self._tick_clock()

        # Picks up reports filed from the terminal console so both surfaces
        # stay in agreement about what has been saved.
        self._sync_timer = QTimer(self)
        self._sync_timer.timeout.connect(self._drain_session_events)
        self._sync_timer.start(400)

        self._bind_shortcuts()

        self.editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self._show_context_menu)

        self._update_counter()
        self._center_window()
        self.editor.setFocus()

    def _drain_session_events(self):
        if self._session is None:
            return
        for event in self._session.drain_events():
            if event.get("origin") == "window":
                continue
            suffix = " (queued - workbook is locked)" if event.get("queued") else ""
            self.status_label.setText(
                f"Filed from the terminal at {event.get('timestamp', '')}{suffix}"
            )
            self.status_label.setStyleSheet(
                "font-size: 13px; font-weight: 600;"
                " color: #22C55E; background: transparent;"
            )

    def closeEvent(self, event):
        # Closing the window ends the whole session, terminal console included.
        if self._session is not None:
            self._session.request_shutdown("window closed")
        event.accept()

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)

        header_card = QFrame()
        header_card.setObjectName("headerCard")
        header_card.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(22, 16, 22, 16)
        header_layout.setSpacing(12)

        self.title_label = QLabel("Task Reporter")
        self.title_label.setObjectName("titleLabel")

        self.clock_label = QLabel("")
        self.clock_label.setObjectName("clockLabel")
        self.clock_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )

        history_btn = QPushButton("\u2630")
        history_btn.setObjectName("historyBtn")
        history_btn.setToolTip("View previous reports")
        history_btn.clicked.connect(self._open_history_dialog)

        help_btn = QPushButton("?")
        help_btn.setObjectName("helpBtn")
        help_btn.setToolTip("View keyboard shortcuts")
        help_btn.clicked.connect(self._open_help_dialog)

        header_layout.addWidget(self.title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.clock_label)
        header_layout.addWidget(history_btn)
        header_layout.addWidget(help_btn)

        outer.addWidget(header_card)

        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.Shape.HLine)
        outer.addWidget(sep)

        content_card = QFrame()
        content_card.setObjectName("contentCard")
        content_layout = QVBoxLayout(content_card)
        content_layout.setContentsMargins(24, 20, 24, 20)
        content_layout.setSpacing(12)

        heading = QLabel("What did you accomplish?")
        heading.setObjectName("headingLabel")
        content_layout.addWidget(heading)

        self.editor = QPlainTextEdit()
        self.editor.setObjectName("editor")
        self.editor.setPlaceholderText("Describe your work clearly and concisely\u2026")
        self.editor.document().setDocumentMargin(12)
        self.editor.setUndoRedoEnabled(True)
        self.editor.textChanged.connect(self._update_counter)
        content_layout.addWidget(self.editor, stretch=1)

        footer = QHBoxLayout()
        footer.setSpacing(12)
        footer.setContentsMargins(0, 4, 0, 0)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setMinimumWidth(160)
        footer.addWidget(self.status_label, stretch=1)

        footer.addStretch(1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setRange(0, MAX_REPORT_LENGTH)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setProperty("danger", False)
        footer.addWidget(self.progress_bar)

        self.counter_label = QLabel(f"0 / {MAX_REPORT_LENGTH}")
        self.counter_label.setObjectName("counterLabel")
        self.counter_label.setProperty("danger", False)
        self.counter_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.counter_label.setMinimumWidth(80)
        footer.addWidget(self.counter_label)

        self.save_btn = QPushButton("Save Report")
        self.save_btn.setObjectName("saveBtn")
        self.save_btn.clicked.connect(self.save_report)
        footer.addWidget(self.save_btn)

        content_layout.addLayout(footer)

        outer.addWidget(content_card, stretch=1)

    def _bind_shortcuts(self):
        ctx = Qt.ShortcutContext.WindowShortcut
        QShortcut(
            QKeySequence("Ctrl+Return"), self, context=ctx, activated=self.save_report
        )
        QShortcut(
            QKeySequence("Ctrl+A"), self, context=ctx, activated=self.editor.selectAll
        )
        QShortcut(
            QKeySequence("Ctrl+Backspace"),
            self,
            context=ctx,
            activated=self._delete_previous_word,
        )
        QShortcut(
            QKeySequence("Ctrl+Delete"),
            self,
            context=ctx,
            activated=self._delete_next_word,
        )
        QShortcut(QKeySequence("Ctrl+Z"), self, context=ctx, activated=self._undo)
        QShortcut(QKeySequence("Ctrl+Y"), self, context=ctx, activated=self._redo)
        QShortcut(QKeySequence("Ctrl+Shift+Z"), self, context=ctx, activated=self._redo)

    def _delete_previous_word(self):
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()
        else:
            cursor.movePosition(
                QTextCursor.MoveOperation.PreviousWord,
                QTextCursor.MoveMode.KeepAnchor,
            )
            cursor.removeSelectedText()
        self.editor.setTextCursor(cursor)
        self._update_counter()

    def _delete_next_word(self):
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()
        else:
            cursor.movePosition(
                QTextCursor.MoveOperation.NextWord,
                QTextCursor.MoveMode.KeepAnchor,
            )
            cursor.removeSelectedText()
        self.editor.setTextCursor(cursor)
        self._update_counter()

    def _undo(self):
        self.editor.undo()
        self._update_counter()

    def _redo(self):
        self.editor.redo()
        self._update_counter()

    def _show_context_menu(self, pos):
        menu = QMenu(self)

        a_undo = QAction("Undo", self)
        a_undo.triggered.connect(self._undo)
        a_undo.setEnabled(self.editor.document().isUndoAvailable())

        a_redo = QAction("Redo", self)
        a_redo.triggered.connect(self._redo)
        a_redo.setEnabled(self.editor.document().isRedoAvailable())

        menu.addAction(a_undo)
        menu.addAction(a_redo)
        menu.addSeparator()

        a_cut = QAction("Cut", self)
        a_cut.triggered.connect(self.editor.cut)
        a_copy = QAction("Copy", self)
        a_copy.triggered.connect(self.editor.copy)
        a_paste = QAction("Paste", self)
        a_paste.triggered.connect(self.editor.paste)
        menu.addAction(a_cut)
        menu.addAction(a_copy)
        menu.addAction(a_paste)
        menu.addSeparator()

        a_sel = QAction("Select All", self)
        a_sel.triggered.connect(self.editor.selectAll)
        menu.addAction(a_sel)

        menu.exec(self.editor.mapToGlobal(pos))

    def _tick_clock(self):
        self.clock_label.setText(datetime.now().strftime("%d/%m/%Y  %H:%M:%S"))

    def _update_counter(self):
        text = self.editor.toPlainText()
        length = len(text)

        self.counter_label.setText(f"{length} / {MAX_REPORT_LENGTH}")
        self.progress_bar.setValue(min(length, MAX_REPORT_LENGTH))

        over = length > MAX_REPORT_LENGTH
        self._set_danger_prop(self.progress_bar, over)
        self._set_danger_prop(self.counter_label, over)

        if over:
            self._set_status(
                f"Report is too long  ({length - MAX_REPORT_LENGTH} chars over limit)",
                danger=True,
            )
        else:
            self._set_status("Ready", danger=False)

    def _set_danger_prop(self, widget, danger: bool):
        widget.setProperty("danger", danger)
        widget.style().unpolish(widget)
        widget.style().polish(widget)
        widget.update()

    def _set_status(self, message: str, danger: bool):
        self.status_label.setText(message)
        colour = "#EF4444" if danger else "#3B82F6"
        self.status_label.setStyleSheet(
            f"font-size: 13px; font-weight: 600;"
            f" color: {colour}; background: transparent;"
        )

    def save_report(self):
        report_text = self.editor.toPlainText().strip()

        if not report_text:
            self._set_status("Report cannot be empty", danger=True)
            QMessageBox.warning(self, "Warning", "The report cannot be empty!")
            return

        if len(report_text) > MAX_REPORT_LENGTH:
            self._set_status(
                f"Please keep report under {MAX_REPORT_LENGTH} characters",
                danger=True,
            )
            QMessageBox.warning(
                self,
                "Warning",
                f"Report is too long ({len(report_text)} chars).\n"
                f"Limit is {MAX_REPORT_LENGTH} characters.",
            )
            return

        try:
            timestamp = append_report_to_excel(report_text)
            self.editor.clear()
            self._update_counter()
            if self._session is not None:
                self._session.record_save("window", timestamp)
            self.status_label.setText(f"Saved at {timestamp} \u2713")
            self.status_label.setStyleSheet(
                "font-size: 13px; font-weight: 600;"
                " color: #22C55E; background: transparent;"
            )
            QMessageBox.information(self, "Success", "Task saved successfully!")

        except ReportQueuedError as exc:
            # Nothing is lost: the report sits in the pending queue and is
            # merged into the workbook on the next successful save.
            self.editor.clear()
            self._update_counter()
            if self._session is not None:
                self._session.record_save("window", exc.timestamp, queued=True)
            self.status_label.setText(f"Queued at {exc.timestamp} \u2713")
            self.status_label.setStyleSheet(
                "font-size: 13px; font-weight: 600;"
                " color: #F59E0B; background: transparent;"
            )
            QMessageBox.information(
                self,
                "Saved to the pending queue",
                f"The workbook could not be written:\n{EXCEL_FILE_PATH}\n\n"
                "It is probably open in Excel. Your report was kept in\n"
                f"{PENDING_FILE_PATH}\n\n"
                "and will be merged in automatically once the file is free.",
            )

        except Exception as e:
            self._set_status("Unexpected error while saving", danger=True)
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")

    def _center_window(self):
        """Fit the window to the screen it is on, then centre it.

        Two ways this used to strand the window off-screen: a minimum size
        larger than the display, and centring on a multi-monitor virtual desktop
        whose primary screen is not where the window actually is.
        """
        # Prefer the screen the pointer is on: on this machine the *primary*
        # screen is reported at offset (1920, 724), so centring on it blindly
        # puts the window on the other monitor - invisible if that monitor is
        # off or is a stale entry left behind by a display change.
        screen = None
        try:
            screen = QApplication.screenAt(QCursor.pos())
        except Exception:
            screen = None
        if screen is None:
            handle = self.windowHandle()
            screen = (handle.screen() if handle is not None else None) or (
                QApplication.primaryScreen()
            )
        if screen is None:
            return
        available = screen.availableGeometry()

        # Never demand more room than the screen has.
        max_w = max(480, available.width() - 80)
        max_h = max(360, available.height() - 80)
        self.setMinimumSize(min(920, max_w), min(680, max_h))
        self.resize(min(self.width(), max_w), min(self.height(), max_h))

        if QApplication.platformName().startswith("wayland"):
            # Wayland clients cannot place themselves; the compositor decides.
            return

        frame = self.frameGeometry()
        frame.moveCenter(available.center())
        # Clamp so the title bar can never end up outside the visible area.
        x = max(available.left(), min(frame.left(), available.right() - frame.width()))
        y = max(available.top(), min(frame.top(), available.bottom() - frame.height()))
        self.move(x, y)

    def _check_or_create_excel(self):
        if os.path.exists(EXCEL_FILE_PATH):
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Reports"
            ws.append(["Date-Time", "Task Report"])
            ws.column_dimensions["A"].width = 25
            ws.column_dimensions["B"].width = 100
            wb.save(EXCEL_FILE_PATH)
        except Exception as e:
            QMessageBox.critical(
                None, "Error", f"Could not create the Excel file:\n{e}"
            )

    def _open_help_dialog(self):
        dlg = ShortcutsDialog(self)
        dlg.exec()

    def _open_history_dialog(self):
        dlg = ReportHistoryDialog(self)
        dlg.exec()


# ---------------------------------------------------------------------------
# Browser UI - the default surface
# ---------------------------------------------------------------------------

# Why a browser and not a desktop toolkit.
#
# The Qt window above renders through WSLg, and WSLg is the one part of this
# stack that cannot be relied on.  Three failures show up as "the UI just did
# not work this time":
#
#   * a plugin fails to load, and Qt answers with qFatal() - an abort, not an
#     exception, so it cannot be caught in-process;
#   * a plugin loads, a surface is created, and the compositor never presents
#     it (the "taskbar button but no window" state);
#   * a stale multi-monitor layout places the window on a screen that is off.
#
# None of that is fixable from inside the app.  The render probe narrows the
# odds and cannot close them: it probes a frameless transparent tool window,
# which does not always take the same presentation path as a real decorated
# toplevel, and WSLg can degrade in the gap between probe and window anyway.
# That is the intermittency - a green probe is not a promise.
#
# Serving the UI over loopback and opening it in the *system* browser removes
# the whole failure class.  WSL2 forwards Windows localhost into the VM, so the
# page is fetched by a native Windows browser, drawn by a native Windows
# process, and placed by the Windows window manager.  No X11, no Wayland, no
# compositor, no GPU path, and nothing outside the standard library.  On a
# normal Linux or macOS desktop the same code opens the same page in the
# default browser, so there is one surface to maintain rather than two.

WEB_HOST = "127.0.0.1"
# A stable port keeps the URL predictable between runs; the ephemeral fallback
# means a busy port can never be the reason the UI does not come up.
WEB_PREFERRED_PORTS = tuple(range(8770, 8780))
# How long after the last page says goodbye to wait before ending the session.
#
# This is deliberately generous, because `pagehide` is a hint and not a verdict:
# browsers fire it for back/forward cache, tab freezing, prerender swaps and
# process moves as well as for a real close, and Chrome was observed firing it
# on a visible page seconds after load.  Acting on it directly meant the UI
# could shut itself down while the user was still looking at it - the exact
# fault this rewrite exists to remove.  So a goodbye only starts a countdown,
# and any ping from any page cancels it: a page that is coming back always
# comes back well inside this window, and a page that is really gone never does.
WEB_BYE_GRACE_SECONDS = 15.0
# Browser-only mode has no other surface to notice a dead browser, so a page
# that stops checking in eventually ends the session.  Background tabs are
# throttled to roughly one timer per minute, hence the generous margin.
WEB_IDLE_TIMEOUT_SECONDS = 300.0
# How long to wait for the browser to actually load the page before trying the
# alternate host spelling and then falling back to printing the URL.
WEB_FIRST_CLIENT_TIMEOUT = 9.0
WEB_MAX_BODY_BYTES = 256 * 1024

WEB_PAGE_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Task Reporter</title>
<link rel="icon" href="/favicon.ico">
<script>
"use strict";
window.CFG = %%CONFIG%%;
// Before first paint, so the window never flashes the wrong theme.
(function () {
  var root = document.documentElement;
  var pref = null;
  try { pref = localStorage.getItem("taskReporterTheme"); } catch (err) { /* private mode */ }
  if (pref !== "light" && pref !== "dark" && pref !== "system") pref = CFG.theme || "system";
  var dark = pref === "dark" ||
    (pref === "system" && window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches);
  root.setAttribute("data-theme-pref", pref);
  root.setAttribute("data-theme", dark ? "dark" : "light");
  if (CFG.app) root.classList.add("is-app");
  if (CFG.mode === "quick") root.classList.add("is-quick");
})();
</script>
<style>
/* Modernist: flat, square, 2px rules, one red.  Tokens copied from the design
   system's styles.css; the dark set is the DARK map in TaskReporterApp.dc.html. */
:root {
  --color-bg: #f3f2f2;
  --color-surface: #eae9e9;
  --color-text: #201e1d;
  --color-accent: #ec3013;
  --color-divider: color-mix(in srgb, #201e1d 40%, transparent);

  --color-neutral-100: #f8f4f4;
  --color-neutral-200: #eae7e7;
  --color-neutral-300: #d7d3d3;
  --color-neutral-400: #bab6b6;
  --color-neutral-500: #9b9797;
  --color-neutral-600: #7d7979;
  --color-neutral-700: #605d5d;
  --color-neutral-800: #444141;
  --color-neutral-900: #2d2b2b;

  --color-accent-100: #fff2ef;
  --color-accent-200: #ffe0d9;
  --color-accent-300: #ffc4b8;
  --color-accent-400: #ff9783;
  --color-accent-500: #ff563c;
  --color-accent-600: #dd2b0f;
  --color-accent-700: #ae1800;
  --color-accent-800: #7c1405;
  --color-accent-900: #4d170e;

  /* Archivo when it is installed; Segoe UI otherwise.  The page runs offline
     in WebView2, so nothing is fetched. */
  --font-heading: "Archivo", "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
  --font-heading-weight: 800;
  --font-body: "Archivo", "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
  --font-mono: ui-monospace, "Cascadia Mono", Consolas, monospace;

  --shadow-sm: 0 1px 2px color-mix(in srgb, #2d2b2b 14%, transparent);
  --shadow-md: 0 3px 10px color-mix(in srgb, #2d2b2b 16%, transparent);
  --shadow-lg: 0 12px 32px color-mix(in srgb, #2d2b2b 22%, transparent);

  --muted: color-mix(in srgb, var(--color-text) 62%, transparent);
  --faint: color-mix(in srgb, var(--color-text) 45%, transparent);
  --hover: color-mix(in srgb, var(--color-text) 8%, transparent);
  --press: color-mix(in srgb, var(--color-text) 14%, transparent);
  color-scheme: light;
}
:root[data-theme="dark"] {
  --color-bg: #201e1d;
  --color-surface: #2d2b2b;
  --color-text: #f3f2f2;
  --color-divider: color-mix(in srgb, #f3f2f2 35%, transparent);
  --color-neutral-100: #444141;
  --color-neutral-800: #eae7e7;
  --color-neutral-900: #000000;
  --color-accent-100: #4d170e;
  --color-accent-200: #7c1405;
  --color-accent-700: #ff9783;
  --color-accent-800: #ffc4b8;
  --shadow-md: 0 3px 10px color-mix(in srgb, #000000 45%, transparent);
  --shadow-lg: 0 12px 32px color-mix(in srgb, #000000 60%, transparent);
  color-scheme: dark;
}

*, *::before, *::after { box-sizing: border-box; border-radius: 0; }
html, body { height: 100%; }
body {
  margin: 0; overflow: hidden;
  background: var(--color-bg); color: var(--color-text);
  font-family: var(--font-body); font-size: 14px; line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}
button, input, select, textarea { font: inherit; color: inherit; }
button { cursor: pointer; }
h3, h4, h6 { font-family: var(--font-heading); font-weight: var(--font-heading-weight); margin: 0; line-height: 1.12; letter-spacing: -0.015em; }
h3 { font-size: 25px; }
h4 { font-size: 20px; }
h6 { font-size: 13px; letter-spacing: 0.08em; text-transform: uppercase; }
code { font-family: var(--font-mono); font-size: 12px; }
:focus { outline: none; }
:focus-visible { outline: 2px solid var(--color-accent); outline-offset: 2px; }
::selection { background: color-mix(in srgb, var(--color-accent) 30%, transparent); }
::placeholder { color: var(--faint); }
.hidden { display: none !important; }
svg { display: block; flex: none; }

/* — buttons — labels sit flush left, even in wide buttons — */
.btn {
  display: inline-flex; align-items: center; justify-content: flex-start; gap: 8px;
  font-family: var(--font-heading); font-weight: var(--font-heading-weight);
  font-size: 14px; line-height: 1.2; color: var(--color-text);
  background: transparent; border: 1px solid transparent;
  padding: 8px 14px; min-height: 36px; white-space: nowrap;
}
.btn:disabled { opacity: 0.45; cursor: not-allowed; }
.btn-primary { background: var(--color-accent); color: #ffffff; }
.btn-primary:hover:not(:disabled) { background: var(--color-accent-600); }
.btn-primary:active:not(:disabled) { background: var(--color-accent-700); }
.btn-secondary { border-color: var(--color-divider); }
.btn-secondary:hover:not(:disabled) { background: var(--hover); }
.btn-secondary:active:not(:disabled) { background: var(--press); }
.btn-ghost { color: var(--color-accent-700); padding-inline: 6px; }
.btn-ghost:hover:not(:disabled) { background: color-mix(in srgb, var(--color-accent) 10%, transparent); }
.btn-ghost:active:not(:disabled) { background: color-mix(in srgb, var(--color-accent) 18%, transparent); }
.btn-icon { width: 36px; height: 36px; padding: 0; justify-content: center; }
.btn-icon:hover { background: var(--hover); }
.btn-icon:active { background: var(--press); }
.btn-icon.is-on { background: var(--color-accent); color: #ffffff; }
.btn-tall { height: 40px; }
.mini {
  flex: none; width: 26px; height: 26px; padding: 0;
  display: grid; place-items: center;
  border: 1px solid var(--color-divider); background: transparent; color: var(--color-text);
}
.mini:hover { background: var(--hover); }
.mini.danger:hover { background: var(--color-accent-100); color: var(--color-accent-700); border-color: var(--color-accent); }

/* — fields — */
.input {
  width: 100%; min-height: 36px; padding: 6px 10px;
  font-size: 14px; color: var(--color-text); caret-color: var(--color-accent);
  background: var(--color-surface); border: 1px solid var(--color-divider);
}
.input:hover { border-color: color-mix(in srgb, var(--color-text) 45%, transparent); }
.input:focus, .input:focus-visible { border-color: var(--color-accent); outline: none; }
textarea.input { resize: none; line-height: 1.55; display: block; }
select.input { cursor: pointer; }
.field { display: flex; flex-direction: column; gap: 5px; min-width: 0; }
.field > label, .field-label { font-size: 12px; color: color-mix(in srgb, var(--color-text) 70%, transparent); }
.tag { display: inline-flex; align-items: center; font-size: 11px; letter-spacing: 0.02em; padding: 3px 10px; white-space: nowrap; }
.tag-small { font-size: 10px; padding: 1px 7px; letter-spacing: 0.06em; text-transform: uppercase; }
.tag-accent { background: var(--color-accent-100); color: var(--color-accent-800); }
.tag-neutral { background: var(--color-neutral-100); color: var(--color-neutral-800); }
.seg { display: inline-flex; border: 1px solid var(--color-divider); align-self: flex-start; }
.seg-opt { display: inline-flex; align-items: center; padding: 7px 14px; font-size: 13px; cursor: pointer; position: relative; }
.seg-opt input { position: absolute; opacity: 0; width: 0; height: 0; pointer-events: none; }
.seg-opt + .seg-opt { border-left: 1px solid var(--color-divider); }
.seg-opt:has(input:checked) { background: var(--color-accent); color: #ffffff; }
.seg-opt:not(:has(input:checked)):hover { background: var(--hover); }
.seg-opt:has(input:focus-visible) { outline: 2px solid var(--color-accent); outline-offset: -2px; }
.check { width: 16px; height: 16px; margin: 0; accent-color: var(--color-accent); flex: none; }
.hint { margin: 0; font-size: 12px; color: var(--muted); line-height: 1.6; }
.hint b { color: var(--color-text); font-weight: 600; }
kbd {
  display: inline-block; padding: 2px 6px; background: var(--color-surface);
  font-family: var(--font-mono); font-size: 12px; white-space: nowrap;
}

/* — the window — */
.shell { height: 100%; display: flex; flex-direction: column; min-height: 0; position: relative; }

/* Title bar - only in the frameless desktop window. */
.titlebar {
  display: none; flex: none; height: 40px; align-items: stretch;
  background: var(--color-bg); border-bottom: 2px solid var(--color-divider);
  user-select: none; -webkit-user-select: none;
}
.is-app .titlebar { display: flex; }
.titlebar-drag { flex: 1; min-width: 0; display: flex; align-items: center; gap: 10px; padding: 0 14px; }
/* WebView2 hands these to Windows as the caption: dragging, Aero Snap,
   double-click to maximize and the system menu all come from Windows. */
.is-app .titlebar-drag, .is-app .quick-bar .drag { app-region: drag; -webkit-app-region: drag; }
.is-app .winbtn { app-region: no-drag; -webkit-app-region: no-drag; }
.titlebar-drag img { width: 16px; height: 16px; pointer-events: none; }
.titlebar-name { font-size: 12px; font-weight: 600; white-space: nowrap; pointer-events: none; }
.titlebar-sub { font-size: 12px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; pointer-events: none; }
.winbtn {
  width: 46px; border: none; background: transparent; color: var(--color-text);
  display: grid; place-items: center; padding: 0;
}
.winbtn:hover { background: var(--hover); }
.winbtn:active { background: var(--press); }
.winbtn.close:hover { background: var(--color-accent); color: #ffffff; }
.winbtn.close:active { background: var(--color-accent-700); color: #ffffff; }
.winbtn .restore-icon { display: none; }
.is-maximized .winbtn .max-icon { display: none; }
.is-maximized .winbtn .restore-icon { display: block; }
/* Resize grips round a frameless window.  Hidden when maximized. */
.grip { display: none; position: fixed; z-index: 100; }
.is-app.native-frame .grip { display: block; }
.is-maximized .grip { display: none !important; }
.grip-n { top: 0; left: 8px; right: 8px; height: 6px; cursor: n-resize; }
.grip-s { bottom: 0; left: 8px; right: 8px; height: 6px; cursor: s-resize; }
.grip-w { left: 0; top: 8px; bottom: 8px; width: 6px; cursor: w-resize; }
.grip-e { right: 0; top: 8px; bottom: 8px; width: 6px; cursor: e-resize; }
.grip-nw { top: 0; left: 0; width: 8px; height: 8px; cursor: nw-resize; }
.grip-ne { top: 0; right: 0; width: 8px; height: 8px; cursor: ne-resize; }
.grip-sw { bottom: 0; left: 0; width: 8px; height: 8px; cursor: sw-resize; }
.grip-se { bottom: 0; right: 0; width: 8px; height: 8px; cursor: se-resize; }
.is-app body { border: 1px solid var(--color-divider); }
.is-app.is-maximized body { border: none; }

/* Header: Report / Board tabs, the clock, the tool buttons. */
.header {
  flex: none; height: 60px; display: flex; align-items: center; gap: 16px;
  padding: 0 20px; border-bottom: 2px solid var(--color-divider);
}
.tabs { display: flex; border: 1px solid var(--color-divider); flex: none; }
.tab {
  display: flex; align-items: center; gap: 8px; height: 36px; padding: 0 18px; border: none;
  font-family: var(--font-heading); font-weight: 800; font-size: 14px;
  background: transparent; color: var(--color-text);
}
.tab + .tab { border-left: 1px solid var(--color-divider); }
.tab:hover { background: var(--hover); }
.tab.is-on { background: var(--color-accent); color: #ffffff; }
.tab-badge {
  min-width: 20px; padding: 0 6px; line-height: 20px; font-size: 11px; text-align: center;
  background: var(--color-accent); color: #ffffff;
}
.tab.is-on .tab-badge { background: #ffffff; color: var(--color-accent-700); }
.tab-badge:empty { display: none; }
.header-gap { flex: 1; }
.clock { font-size: 13px; font-variant-numeric: tabular-nums; color: var(--muted); white-space: nowrap; }
.tools { display: flex; gap: 4px; }
.theme-sun { display: none; }
[data-theme="dark"] .theme-sun { display: block; }
[data-theme="dark"] .theme-moon { display: none; }

/* Excel has the workbook: a full-width strip under the header. */
.banner {
  flex: none; display: flex; align-items: center; gap: 12px; padding: 10px 20px;
  background: var(--color-accent-100); color: var(--color-accent-800);
  border-bottom: 2px solid var(--color-accent);
}
.banner-text { flex: 1; font-size: 13px; }
.banner .btn-ghost { color: var(--color-accent-800); font-size: 13px; }

.main { flex: 1; min-height: 0; display: flex; }
.view { flex: 1; min-width: 0; min-height: 0; display: flex; flex-direction: column; gap: 14px; padding: 24px; }

/* Report */
#editor { flex: 1; min-height: 160px; padding: 14px 16px; font-size: 15px; line-height: 1.6; }
.footer {
  display: flex; align-items: center; gap: 12px; padding-top: 14px;
  border-top: 2px solid var(--color-divider); flex-wrap: wrap; flex: none;
}
.status { flex: 1; min-width: 120px; font-size: 13px; font-weight: 600; color: var(--muted); }
.status.success, .status.danger, .status.warn { color: var(--color-accent-700); }
.progress { width: 140px; height: 6px; background: color-mix(in srgb, var(--color-text) 12%, transparent); flex: none; }
.progress-fill { height: 100%; width: 0; background: var(--color-accent); transition: width .12s linear; }
.progress-fill.danger { background: var(--color-accent-700); }
.counter { min-width: 80px; text-align: right; font-size: 13px; font-variant-numeric: tabular-nums; color: var(--muted); }
.counter.danger { color: var(--color-accent-700); font-weight: 600; }
.wide-btn { min-width: 156px; height: 40px; }

/* Board */
.board-head { display: flex; align-items: baseline; gap: 12px; flex: none; }
.board-sub { margin-left: auto; font-size: 13px; color: var(--muted); }
.board-sub b { color: var(--color-accent-700); }
.composer { display: flex; gap: 8px; align-items: flex-start; flex-wrap: wrap; flex: none; }
.c-date { flex: none; width: 150px; height: 40px; font-variant-numeric: tabular-nums; }
.c-proj { flex: none; width: 190px; }
.c-proj .input { height: 40px; }
#taskText { flex: 1; min-width: 220px; min-height: 40px; padding: 9px 12px; line-height: 1.45; overflow-y: auto; }
.c-add { flex: none; height: 40px; min-width: 124px; }

.combo { position: relative; display: flex; }
.combo .input { padding-right: 32px; }
.combo-caret {
  position: absolute; top: 0; right: 0; width: 32px; height: 100%;
  border: none; background: transparent; color: var(--color-text); display: grid; place-items: center;
}
.combo-panel {
  position: absolute; top: calc(100% + 4px); left: 0; z-index: 30;
  min-width: 100%; width: max-content; max-width: 380px; max-height: 260px; overflow-y: auto;
  display: none; flex-direction: column; padding: 4px;
  background: var(--color-bg); border: 2px solid var(--color-text); box-shadow: var(--shadow-md);
}
.combo-panel.open { display: flex; }
.combo-option { display: flex; align-items: center; }
.combo-option:hover, .combo-option.is-active { background: var(--hover); }
.combo-name {
  flex: 1; min-width: 0; text-align: left; padding: 8px 10px; border: none; background: transparent;
  font-family: var(--font-mono); font-size: 13px; line-height: 1.4; white-space: normal; overflow-wrap: anywhere;
}
.combo-forget {
  flex: none; width: 24px; height: 24px; margin: 0 4px; padding: 0; border: none;
  background: transparent; color: transparent; display: grid; place-items: center;
}
.combo-option:hover .combo-forget, .combo-option.is-active .combo-forget { color: var(--faint); }
.combo-forget:hover { background: var(--color-accent-100); color: var(--color-accent-700) !important; }
.combo-sep { height: 1px; margin: 4px 6px; background: var(--color-divider); flex: none; }
.combo-hint { padding: 8px 10px; font-size: 12px; color: color-mix(in srgb, var(--color-text) 70%, transparent); }
.combo-hint b { color: var(--color-accent-700); }

.chips { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; flex: none; }
.chips-label { font-size: 11px; letter-spacing: 0.08em; text-transform: uppercase; color: var(--muted); margin-right: 4px; }
.chip {
  display: flex; align-items: center; gap: 6px; height: 28px; padding: 0 10px; font-size: 12px;
  border: 1px solid var(--color-divider); background: transparent; color: var(--color-text);
}
.chip:hover { background: var(--hover); }
.chip.is-on { background: var(--color-text); border-color: var(--color-text); color: var(--color-bg); }
.chip-count { font-variant-numeric: tabular-nums; opacity: 0.7; }

.board-scroll {
  flex: 1; min-height: 0; overflow: auto;
  border-top: 2px solid var(--color-divider); border-bottom: 2px solid var(--color-divider);
}
.day { padding-bottom: 8px; border-bottom: 1px solid var(--color-divider); }
.day:last-child { border-bottom: none; }
.day-head {
  position: sticky; top: 0; z-index: 2; display: flex; align-items: center; gap: 12px;
  padding: 12px 0 8px; background: var(--color-bg);
}
.day-date { font-family: var(--font-heading); font-weight: 800; font-size: 17px; font-variant-numeric: tabular-nums; }
.day-when { font-size: 12px; color: var(--muted); }
.day-count { margin-left: auto; font-size: 12px; font-variant-numeric: tabular-nums; color: var(--muted); }
.day-count.all-done { color: var(--color-accent-700); }
.day-copy { font-size: 12px; gap: 6px; min-height: 28px; padding: 4px 8px; }
.proj { padding: 2px 0 6px; }
.proj-head { padding: 4px 0 4px 28px; font-family: var(--font-mono); font-size: 13px; font-weight: 700; }
.proj-head.no-project { color: var(--muted); font-style: italic; font-family: var(--font-body); font-weight: 600; }
.proj-head.drop-into { background: var(--color-accent-100); color: var(--color-accent-800); outline: 2px dashed var(--color-accent); outline-offset: -2px; }

.task { display: flex; align-items: flex-start; gap: 10px; padding: 7px 8px 7px 0; position: relative; }
.task:hover, .task:focus-within { background: color-mix(in srgb, var(--color-text) 5%, transparent); }
.task:focus-visible { outline: 2px solid var(--color-accent); outline-offset: -2px; }
.task.dragging { opacity: 0.4; }
.task.drop-before::before, .task.drop-after::after {
  content: ""; position: absolute; left: 0; right: 0; height: 2px; background: var(--color-accent);
}
.task.drop-before::before { top: -1px; }
.task.drop-after::after { bottom: -1px; }
.grip-handle { flex: none; width: 18px; display: grid; place-items: center; padding-top: 3px; cursor: grab; color: var(--faint); }
.grip-handle:hover { color: var(--color-text); }
.box {
  flex: none; width: 18px; height: 18px; margin-top: 2px; padding: 0;
  display: grid; place-items: center; color: #ffffff;
  background: transparent; border: 2px solid color-mix(in srgb, var(--color-text) 55%, transparent);
}
.box:hover:not(:disabled) { border-color: var(--color-accent); }
.box svg { visibility: hidden; }
.task.is-done .box { background: var(--color-accent); border-color: var(--color-accent); }
.task.is-done .box svg { visibility: visible; }
.task.is-filed .box { background: color-mix(in srgb, var(--color-text) 30%, transparent); border-color: transparent; cursor: default; }
.task-text {
  flex: 1; min-width: 0; padding: 0; border: none; background: transparent; text-align: left;
  font-size: 14px; line-height: 1.5; white-space: pre-wrap; overflow-wrap: anywhere; color: var(--color-text);
}
.task-text:hover { text-decoration: underline dotted var(--faint); }
.task.is-done .task-text { color: var(--muted); }
.task.is-filed .task-text { color: var(--faint); text-decoration: line-through; }
.clip {
  flex: none; display: flex; align-items: center; gap: 4px; height: 22px; padding: 0 7px;
  border: 1px solid var(--color-divider); background: var(--color-surface); font-size: 11px;
}
.clip:hover { border-color: var(--color-accent); color: var(--color-accent-700); }
.task-meta { flex: none; padding-top: 2px; font-size: 11px; font-variant-numeric: tabular-nums; color: var(--faint); white-space: nowrap; }
.board-empty { padding: 48px 8px; line-height: 1.7; white-space: pre-line; color: var(--muted); }
.board-footer { display: flex; align-items: center; gap: 12px; flex: none; flex-wrap: wrap; }
.toggle { display: flex; align-items: center; gap: 8px; font-size: 13px; cursor: pointer; user-select: none; }
.file-btn { min-width: 190px; height: 40px; }

/* Settings */
.settings { overflow: auto; }
.settings-inner { max-width: 860px; display: flex; flex-direction: column; }
.settings-head { display: flex; align-items: baseline; gap: 12px; padding-bottom: 16px; }
.settings-grid { display: grid; grid-template-columns: 200px minmax(0, 1fr); border-top: 2px solid var(--color-divider); }
.settings-grid > h6 { padding: 20px 0; }
.settings-grid > .s-body { padding: 20px 0; display: flex; flex-direction: column; gap: 14px; }
.settings-grid > .s-sep { border-top: 2px solid var(--color-divider); }
.path-row { display: flex; gap: 8px; }
.path-row .input { flex: 1; font-family: var(--font-mono); font-size: 13px; }
.cards { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.file-card { display: flex; flex-direction: column; gap: 4px; padding: 12px; background: var(--color-surface); min-width: 0; }
.file-card-name { font-weight: 600; font-family: var(--font-mono); font-size: 13px; overflow-wrap: anywhere; }
.file-card-note { font-size: 12px; color: var(--muted); }
.file-card .btn-ghost { align-self: flex-start; font-size: 13px; margin-top: 4px; }
.inline-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }

/* Overlays */
.backdrop {
  position: fixed; left: 0; right: 0; bottom: 0; top: 0; z-index: 20;
  display: none; place-items: center; padding: 24px;
  background: color-mix(in srgb, var(--color-neutral-900) 50%, transparent);
}
.is-app .backdrop, .is-app .panel-scrim { top: 40px; }
.backdrop.open { display: grid; }
.modal {
  width: 100%; max-height: 100%; overflow: auto; display: flex; flex-direction: column; gap: 14px;
  padding: 20px 24px; background: var(--color-bg); border: 2px solid var(--color-text); box-shadow: var(--shadow-lg);
}
.modal-sm { max-width: 440px; }
.modal-md { max-width: 520px; }
.modal-edit { max-width: 680px; }
.modal-lg { max-width: 860px; height: 100%; }
.modal-head { display: flex; align-items: center; gap: 12px; flex: none; }
.modal-head .btn-icon { margin-left: auto; }
.dialog-title { font-family: var(--font-heading); font-weight: var(--font-heading-weight); font-size: 20px; }
.dialog-body { font-size: 14px; opacity: 0.85; }
.modal-count { font-size: 12px; color: var(--muted); }
.modal-actions { display: flex; align-items: center; gap: 8px; flex: none; }
.modal-actions .keyhint { margin-right: auto; font-size: 12px; color: var(--muted); }
.edit-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
#tText, #editText { min-height: 120px; resize: vertical; }
.edit-error { min-height: 16px; font-size: 12px; font-weight: 600; color: var(--color-accent-700); }
#confirmBackdrop { z-index: 25; background: color-mix(in srgb, var(--color-neutral-900) 35%, transparent); }
#confirmBackdrop .modal { border-color: var(--color-accent); }
.confirm-actions .btn-ghost { margin-left: auto; color: var(--color-accent-700); }

/* Task files */
.files { display: flex; flex-direction: column; gap: 10px; padding-top: 14px; border-top: 2px solid var(--color-divider); }
.files-head { display: flex; align-items: center; gap: 10px; }
.files-head .btn { margin-left: auto; font-size: 13px; }
.file-row { display: flex; align-items: center; gap: 12px; padding: 8px; background: var(--color-surface); }
.file-row.is-new { outline: 1px dashed var(--color-accent); outline-offset: -1px; }
.thumb {
  flex: none; width: 56px; height: 42px; display: grid; place-items: center; overflow: hidden;
  background: var(--color-bg); border: 1px solid var(--color-divider); color: var(--muted);
}
.thumb img { width: 100%; height: 100%; object-fit: cover; display: block; }
.thumb-ext { font-size: 10px; font-weight: 800; letter-spacing: 0.04em; text-transform: uppercase; color: var(--color-text); }
.file-info { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.file-name { font-size: 13px; font-weight: 600; overflow-wrap: anywhere; }
.file-meta { font-size: 11px; color: var(--muted); font-variant-numeric: tabular-nums; }
.file-row .btn-ghost { font-size: 12px; min-height: 28px; padding: 4px 6px; }
.drop {
  display: flex; align-items: center; gap: 10px; padding: 14px 16px; font-size: 13px; cursor: pointer;
  border: 2px dashed var(--color-divider); color: color-mix(in srgb, var(--color-text) 70%, transparent);
}
.drop.over { border-color: var(--color-accent); background: var(--color-accent-100); }

/* Filing preview */
.preview-wrap { flex: 1; min-height: 0; overflow: auto; border-top: 2px solid var(--color-divider); border-bottom: 2px solid var(--color-divider); }
.pv-group { padding: 14px 0; border-bottom: 1px solid var(--color-divider); }
.pv-group:last-child { border-bottom: none; }
.pv-head { display: flex; align-items: baseline; gap: 10px; padding-bottom: 8px; }
.pv-when { font-weight: 800; font-variant-numeric: tabular-nums; }
.pv-note { margin-left: auto; font-size: 11px; font-variant-numeric: tabular-nums; color: var(--muted); }
.pv-warn { color: var(--color-accent-700); font-weight: 600; }
.pv-merge, .pv-filter { margin-bottom: 8px; padding: 8px 12px; font-size: 12px; background: var(--color-surface); }
.pv-filter { margin: 0; }
.pv-body { margin: 0; padding: 12px 14px; background: var(--color-surface); font-family: inherit; font-size: 13px; line-height: 1.6; white-space: pre-wrap; overflow-wrap: anywhere; }
.pv-locked { flex: none; padding: 10px 14px; font-size: 12px; background: var(--color-accent-100); color: var(--color-accent-800); }

/* Keyboard shortcuts */
.keys { width: 100%; border-collapse: collapse; font-size: 13px; }
.keys td { padding: 6px 0; border-bottom: 1px solid var(--color-divider); vertical-align: top; }
.keys td:first-child { width: 1%; white-space: nowrap; padding-right: 12px; }

/* Previous reports: a side panel */
.panel-scrim {
  position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 10;
  background: color-mix(in srgb, var(--color-neutral-900) 25%, transparent);
  opacity: 0; pointer-events: none; transition: opacity .18s ease;
}
.panel {
  position: fixed; top: 0; right: 0; bottom: 0; z-index: 11; width: 560px; max-width: 100%;
  display: flex; flex-direction: column; background: var(--color-bg);
  border-left: 2px solid var(--color-text); box-shadow: var(--shadow-lg);
  transform: translateX(105%); transition: transform .2s ease; visibility: hidden;
}
.is-app .panel { top: 40px; }
.panel-open .panel-scrim { opacity: 1; pointer-events: auto; }
.panel-open .panel { transform: none; visibility: visible; }
.panel-head { display: flex; align-items: center; gap: 12px; padding: 16px 20px 12px; }
.panel-head .btn-icon { margin-left: auto; }
.panel-tools { display: flex; gap: 8px; padding: 0 20px 14px; border-bottom: 2px solid var(--color-divider); }
.search { position: relative; flex: 1; }
.search svg { position: absolute; left: 10px; top: 10px; opacity: 0.6; pointer-events: none; }
.search .input { padding-left: 34px; }
#historyProject { width: 170px; flex: none; }
.panel-list { flex: 1; min-height: 0; overflow: auto; }
.queued-strip { padding: 8px 20px; font-size: 11px; letter-spacing: 0.08em; text-transform: uppercase; background: var(--color-accent-100); color: var(--color-accent-800); }
.h-row {
  display: grid; grid-template-columns: 32px minmax(0, 1fr) auto; gap: 10px;
  padding: 14px 20px; border-bottom: 1px solid var(--color-divider);
}
.h-row:hover { background: color-mix(in srgb, var(--color-text) 4%, transparent); }
.h-num { font-size: 12px; font-variant-numeric: tabular-nums; color: var(--muted); padding-top: 1px; }
.h-main { display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.h-when { display: flex; align-items: center; gap: 8px; font-size: 12px; font-weight: 600; font-variant-numeric: tabular-nums; }
.h-text { font-size: 13px; line-height: 1.6; white-space: pre-wrap; overflow-wrap: anywhere; }
.h-text mark { background: var(--color-accent-200); color: inherit; }
.h-acts { display: flex; gap: 4px; align-items: flex-start; }
.h-acts .mini { width: 28px; height: 28px; }
.empty { padding: 32px 20px; color: var(--muted); }

/* Toast */
.toast {
  position: fixed; left: 20px; bottom: 20px; z-index: 40; display: flex; align-items: center; gap: 16px;
  min-width: 320px; max-width: min(560px, calc(100% - 40px)); padding: 12px 12px 12px 16px;
  background: var(--color-text); color: var(--color-bg); box-shadow: var(--shadow-lg);
}
.toast-msg { flex: 1; font-size: 13px; overflow-wrap: anywhere; }
.toast-undo {
  border: none; background: transparent; color: var(--color-accent-400);
  font-family: var(--font-heading); font-weight: 800; font-size: 13px; padding: 4px 8px;
}
[data-theme="dark"] .toast-undo { color: var(--color-accent-600); }
.toast-undo:hover { background: color-mix(in srgb, var(--color-bg) 12%, transparent); }
.toast-x { border: none; background: transparent; color: var(--color-bg); width: 24px; height: 24px; display: grid; place-items: center; padding: 0; }

/* Quick add (the tray window) */
.quick { display: none; height: 100%; flex-direction: column; border: 2px solid var(--color-text); background: var(--color-bg); }
.is-quick .quick { display: flex; }
.is-quick .shell, .is-quick .grip { display: none !important; }
.is-quick body { border: none; }
.quick-bar { flex: none; display: flex; align-items: center; gap: 10px; height: 40px; padding: 0 6px 0 14px; border-bottom: 2px solid var(--color-divider); user-select: none; }
.quick-bar .drag { flex: 1; display: flex; align-items: center; gap: 10px; height: 100%; min-width: 0; }
.quick-bar img { width: 16px; height: 16px; pointer-events: none; }
.quick-bar span { font-size: 12px; font-weight: 600; pointer-events: none; }
.quick-bar .winbtn { width: 32px; height: 32px; }
.quick-body { flex: 1; min-height: 0; overflow: auto; display: flex; flex-direction: column; gap: 10px; padding: 16px; }
.quick-row { display: grid; grid-template-columns: 140px 1fr; gap: 8px; }
#qText { min-height: 84px; flex: 1; }
.quick-files { display: flex; flex-wrap: wrap; gap: 6px; }
.quick-file { display: flex; align-items: center; gap: 6px; padding: 2px 4px 2px 8px; font-size: 12px; background: var(--color-surface); max-width: 100%; }
.quick-file span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.quick-file button { border: none; background: transparent; width: 20px; height: 20px; display: grid; place-items: center; padding: 0; }
.quick-file button:hover { background: var(--color-accent-100); color: var(--color-accent-700); }
.quick-attach { display: flex; align-items: center; gap: 10px; font-size: 12px; color: var(--muted); }
.quick-attach.over { color: var(--color-accent-700); }
.quick-foot { flex: none; display: flex; align-items: center; gap: 8px; padding-top: 10px; border-top: 2px solid var(--color-divider); }
.quick-foot .hint { flex: 1; }
.quick-msg { min-height: 16px; font-size: 12px; font-weight: 600; color: var(--color-accent-700); }
.is-quick .backdrop { top: 0; }

/* The session-over curtain */
#gone {
  position: fixed; inset: 0; z-index: 60; display: none; align-items: center; justify-content: center;
  text-align: center; white-space: pre-line; padding: 24px; background: var(--color-bg); color: var(--muted); font-size: 15px;
}
#gone.open { display: flex; }

@media (max-width: 900px) {
  .clock { display: none; }
  .task-meta { display: none; }
  .cards { grid-template-columns: 1fr; }
  .settings-grid { grid-template-columns: 1fr; }
  .settings-grid > h6 { padding-bottom: 0; }
  .settings-grid > .s-body.s-sep { border-top: none; }
}
@media (max-width: 640px) {
  .header { padding: 0 12px; gap: 8px; }
  .view { padding: 16px; }
  .c-date, .c-proj { flex: 1 1 140px; width: auto; }
  #taskText { flex: 1 0 100%; min-width: 0; }
  .c-add { flex: 1 0 100%; }
  .edit-row { grid-template-columns: 1fr; }
  .file-row { flex-wrap: wrap; }
}
</style>
</head>
<body>
<div class="grip grip-n" data-edge="n"></div><div class="grip grip-s" data-edge="s"></div>
<div class="grip grip-w" data-edge="w"></div><div class="grip grip-e" data-edge="e"></div>
<div class="grip grip-nw" data-edge="nw"></div><div class="grip grip-ne" data-edge="ne"></div>
<div class="grip grip-sw" data-edge="sw"></div><div class="grip grip-se" data-edge="se"></div>

<div class="shell" id="shell">
  <div class="titlebar" id="titlebar">
    <div class="titlebar-drag pywebview-drag-region" id="dragRegion" title="">
      <img src="/favicon.ico" alt="">
      <span class="titlebar-name">Task Reporter</span>
      <span class="titlebar-sub" id="winSub"></span>
    </div>
    <button class="winbtn" id="winMin" aria-label="Minimize" title="Minimize"><i data-icon="minus" data-size="16" data-stroke="1.5"></i></button>
    <button class="winbtn" id="winMax" aria-label="Maximize" title="Maximize"><span class="max-icon"><i data-icon="square" data-size="13" data-stroke="1.75"></i></span><span class="restore-icon"><i data-icon="restore" data-size="13" data-stroke="1.75"></i></span></button>
    <button class="winbtn close" id="winClose" aria-label="Close" title="Close"><i data-icon="x" data-size="16" data-stroke="1.5"></i></button>
  </div>

  <header class="header">
    <div class="tabs" role="tablist">
      <button class="tab is-on" id="tabReport" role="tab" title="Write a report (Ctrl+B)">Report</button>
      <button class="tab" id="tabBoard" role="tab" title="Track tasks (Ctrl+B)">Board<span class="tab-badge" id="tabBadge"></span></button>
    </div>
    <div class="header-gap"></div>
    <span class="clock" id="clock"></span>
    <div class="tools">
      <button class="btn btn-icon" id="revealBtn" aria-label="Show task_reports.xlsx in Explorer"><i data-icon="folder"></i></button>
      <button class="btn btn-icon" id="historyBtn" title="Previous reports (Ctrl+H)" aria-label="Previous reports"><i data-icon="history"></i></button>
      <button class="btn btn-icon" id="themeBtn" title="Switch light / dark" aria-label="Switch light / dark"><span class="theme-moon"><i data-icon="moon"></i></span><span class="theme-sun"><i data-icon="sun"></i></span></button>
      <button class="btn btn-icon" id="helpBtn" title="Keyboard shortcuts" aria-label="Keyboard shortcuts"><i data-icon="help"></i></button>
      <button class="btn btn-icon" id="settingsBtn" title="Settings" aria-label="Settings"><i data-icon="settings"></i></button>
    </div>
  </header>

  <div class="banner hidden" id="excelBanner" role="status">
    <i data-icon="alert" data-stroke="2"></i>
    <span class="banner-text" id="excelBannerText"></span>
    <button class="btn btn-ghost" id="viewQueuedBtn">View queued</button>
  </div>

  <div class="main">
    <main class="view" id="reportView">
      <h3>What did you accomplish?</h3>
      <textarea id="editor" class="input" placeholder="Describe your work clearly and concisely&#8230;" autofocus></textarea>
      <span class="hint"><b>Enter</b> new line · <b>Ctrl + Enter</b> save · drafts are kept if the window closes</span>
      <div class="footer">
        <span class="status" id="status">Ready</span>
        <div class="progress"><div class="progress-fill" id="progressFill"></div></div>
        <span class="counter" id="counter"></span>
        <button class="btn btn-primary wide-btn" id="saveBtn">Save Report</button>
      </div>
    </main>

    <section class="view hidden" id="boardView">
      <div class="board-head">
        <h3>Task Board</h3>
        <span class="board-sub" id="boardSummary"></span>
      </div>

      <form class="composer" id="composer" autocomplete="off">
        <input type="date" id="taskDate" class="input c-date" title="The day this task belongs to">
        <div class="combo c-proj">
          <input type="text" id="taskProject" class="input" placeholder="project" spellcheck="false"
                 maxlength="%%MAXPROJECT%%" role="combobox" aria-expanded="false"
                 aria-autocomplete="list" aria-controls="taskProjectPanel"
                 title="Project name - becomes the [bracketed] header in the filed report">
          <button type="button" class="combo-caret" id="taskProjectCaret" tabindex="-1" aria-label="Show existing projects"><i data-icon="chevron" data-size="16" data-stroke="2"></i></button>
          <div class="combo-panel" id="taskProjectPanel" role="listbox"></div>
        </div>
        <textarea id="taskText" class="input" rows="1" maxlength="%%MAXTASK%%"
                  placeholder="What needs doing&#8230;   Shift + Enter for a new line"></textarea>
        <button type="submit" class="btn btn-primary c-add"><i data-icon="plus" data-size="16" data-stroke="2.5"></i>Add Task</button>
      </form>

      <div class="chips" id="projectChips"></div>

      <div class="board-scroll" id="boardScroll"></div>

      <div class="board-footer">
        <span class="status" id="boardStatus">Ready</span>
        <label class="toggle" title="Show tasks already written into the workbook">
          <input type="checkbox" class="check" id="showFiled">Show filed
        </label>
        <button class="btn btn-secondary btn-tall" id="clearFiledBtn" title="Remove filed tasks from the board">Clear Filed</button>
        <button class="btn btn-primary file-btn" id="fileBtn">File Checked Tasks</button>
      </div>
    </section>

    <section class="view settings hidden" id="settingsView">
      <div class="settings-inner">
        <div class="settings-head">
          <h3>Settings</h3>
          <button class="btn btn-ghost" id="settingsBack" style="margin-left: auto; font-size: 13px;">Back</button>
        </div>
        <div class="settings-grid">
          <h6>Files</h6>
          <div class="s-body">
            <div class="field">
              <label for="sDataDir">Data folder — task_reports.xlsx, task_board.json and task files live here</label>
              <div class="path-row">
                <input class="input" id="sDataDir" readonly>
                <button class="btn btn-secondary" id="sChangeDir">Change…</button>
                <button class="btn btn-secondary" id="sOpenDir">Open in Explorer</button>
              </div>
              <div class="hint" id="sDirSource"></div>
            </div>
            <div class="cards">
              <div class="file-card">
                <span class="file-card-name">task_reports.xlsx</span>
                <span class="file-card-note" id="sWorkbookNote"></span>
                <button class="btn btn-ghost" data-reveal="workbook">Show in Explorer</button>
              </div>
              <div class="file-card">
                <span class="file-card-name">task_board.json</span>
                <span class="file-card-note" id="sBoardNote"></span>
                <button class="btn btn-ghost" data-reveal="board">Show in Explorer</button>
              </div>
              <div class="file-card">
                <span class="file-card-name">.task_reporter_app.log</span>
                <span class="file-card-note" id="sLogNote"></span>
                <button class="btn btn-ghost" data-open="log">Open log</button>
              </div>
            </div>
          </div>
          <h6 class="s-sep">Appearance</h6>
          <div class="s-body s-sep">
            <span class="field-label">Theme</span>
            <div class="seg" id="themeSeg">
              <label class="seg-opt"><input type="radio" name="themePick" value="light">Light</label>
              <label class="seg-opt"><input type="radio" name="themePick" value="dark">Dark</label>
              <label class="seg-opt"><input type="radio" name="themePick" value="system">System</label>
            </div>
          </div>
          <h6 class="s-sep">System tray</h6>
          <div class="s-body s-sep">
            <label class="toggle" style="font-size: 14px;"><input type="checkbox" class="check" id="sKeepInTray">Keep Task Reporter in the system tray when the window is closed</label>
            <div class="field" style="max-width: 420px;">
              <label for="sHotkey">Quick-add shortcut (works anywhere in Windows)</label>
              <div class="path-row">
                <input class="input" id="sHotkey" spellcheck="false" placeholder="Ctrl+Alt+T">
                <button class="btn btn-secondary" id="sHotkeySave">Set</button>
              </div>
            </div>
            <span class="hint" id="sTrayNote">Left-click the tray icon to quick-add a task. Right-click: Open, Quick add task, Previous reports, Quit.</span>
          </div>
          <h6 class="s-sep">Excel</h6>
          <div class="s-body s-sep">
            <div class="inline-row">
              <span class="tag tag-neutral" id="sExcelTag">WORKBOOK FREE</span>
              <span class="hint" id="sQueueNote">Nothing queued</span>
            </div>
          </div>
        </div>
      </div>
    </section>
  </div>
</div>

<!-- Previous reports: slides in from the right -->
<div class="panel-scrim" id="historyScrim"></div>
<aside class="panel" id="historyPanel" aria-label="Previous reports">
  <div class="panel-head">
    <h4>Previous Reports</h4>
    <span class="modal-count" id="historyCount"></span>
    <button class="btn btn-icon" id="historyClose" aria-label="Close"><i data-icon="x" data-stroke="2"></i></button>
  </div>
  <div class="panel-tools">
    <div class="search">
      <i data-icon="search" data-size="16" data-stroke="2"></i>
      <input class="input" id="historySearch" placeholder="Search text or date (07/09)…" spellcheck="false">
    </div>
    <select class="input" id="historyProject" aria-label="Project"></select>
  </div>
  <div class="panel-list" id="historyBody"></div>
</aside>

<div class="backdrop" id="previewBackdrop">
  <div class="modal modal-lg">
    <div class="modal-head">
      <span class="dialog-title">File Checked Tasks</span>
      <span class="modal-count" id="previewCount"></span>
    </div>
    <p class="hint">One report row per day, with the checked tasks under their project headers. The tasks stay on the board afterwards, struck through, so nothing gets filed twice.</p>
    <p class="pv-filter hidden" id="previewFilter"></p>
    <div class="pv-locked hidden" id="previewLocked"></div>
    <div class="preview-wrap" id="previewWrap"></div>
    <div class="modal-actions" style="justify-content: flex-end;">
      <button class="btn btn-secondary" data-close>Cancel</button>
      <button class="btn btn-primary" id="previewConfirm" style="min-width: 170px;">Write to Workbook</button>
    </div>
  </div>
</div>

<div class="backdrop" id="taskEditBackdrop">
  <div class="modal modal-edit">
    <div class="modal-head">
      <span class="dialog-title">Edit Task</span>
      <span class="tag tag-accent tag-small hidden" id="tDirty">Unsaved</span>
      <button class="btn btn-icon" data-close aria-label="Close"><i data-icon="x" data-stroke="2"></i></button>
    </div>
    <div class="edit-row">
      <div class="field"><label for="tDate">Date</label><input type="date" class="input" id="tDate"></div>
      <div class="field">
        <label for="tProject">Project</label>
        <div class="combo">
          <input type="text" class="input" id="tProject" placeholder="(no project)" spellcheck="false" autocomplete="off"
                 maxlength="%%MAXPROJECT%%" role="combobox" aria-expanded="false" aria-autocomplete="list" aria-controls="tProjectPanel">
          <button type="button" class="combo-caret" id="tProjectCaret" tabindex="-1" aria-label="Show existing projects"><i data-icon="chevron" data-size="16" data-stroke="2"></i></button>
          <div class="combo-panel" id="tProjectPanel" role="listbox"></div>
        </div>
      </div>
    </div>
    <div class="field">
      <label for="tText">Task</label>
      <textarea class="input" id="tText" maxlength="%%MAXTASK%%"></textarea>
    </div>
    <div class="files" id="tFiles">
      <div class="files-head">
        <h6>Task files</h6>
        <span class="modal-count" id="tFilesCount"></span>
        <input type="file" id="tFilesInput" multiple hidden>
        <button class="btn btn-secondary" id="tFilesAdd"><i data-icon="upload" data-size="15" data-stroke="2"></i>Add files…</button>
      </div>
      <div id="tFilesList" style="display: flex; flex-direction: column; gap: 8px;"></div>
      <div class="drop" id="tDrop" tabindex="0" role="button">
        <i data-icon="paperclip" data-stroke="2"></i>
        <span>Drop images, CSVs or any file here — or paste an image with <b>Ctrl + V</b>. Copies are kept in <code>task_files\</code> next to the workbook.</span>
      </div>
    </div>
    <div class="edit-error" id="tError"></div>
    <div class="modal-actions">
      <span class="keyhint"><b>Ctrl + Enter</b> save · <b>Esc</b> close</span>
      <button class="btn btn-secondary" data-close>Cancel</button>
      <button class="btn btn-primary" id="tSave" style="min-width: 140px;">Save Changes</button>
    </div>
  </div>
</div>

<div class="backdrop" id="editBackdrop">
  <div class="modal modal-edit">
    <div class="modal-head">
      <span class="dialog-title">Edit Report</span>
      <span class="tag tag-accent tag-small hidden" id="editDirty">Unsaved</span>
      <button class="btn btn-icon" data-close aria-label="Close"><i data-icon="x" data-stroke="2"></i></button>
    </div>
    <div class="field"><label for="editWhen">Date-Time</label><input type="text" class="input" id="editWhen" autocomplete="off" spellcheck="false" style="font-variant-numeric: tabular-nums;"></div>
    <div class="field"><label for="editText">Task Report</label><textarea class="input" id="editText" style="min-height: 200px;"></textarea></div>
    <div class="edit-error" id="editError"></div>
    <div class="modal-actions">
      <span class="keyhint"><b>Ctrl + Enter</b> save · <b>Esc</b> close</span>
      <button class="btn btn-secondary" data-close>Cancel</button>
      <button class="btn btn-primary" id="editSave" style="min-width: 140px;">Save Changes</button>
    </div>
  </div>
</div>

<div class="backdrop" id="helpBackdrop">
  <div class="modal modal-md">
    <span class="dialog-title">Keyboard Shortcuts</span>
    <table class="keys">
      <tr><td><kbd>Ctrl + Enter</kbd></td><td>Save the report · in a dialog: save it · on the board: file the checked tasks</td></tr>
      <tr><td><kbd>Enter</kbd></td><td>In the report: new line</td></tr>
      <tr><td><kbd>Enter</kbd></td><td>In the add-task box: add the task</td></tr>
      <tr><td><kbd>Shift + Enter</kbd></td><td>In the add-task box: new line in the task</td></tr>
      <tr><td><kbd>Ctrl + B</kbd></td><td>Switch between Report and Board</td></tr>
      <tr><td><kbd>Ctrl + H</kbd></td><td>Previous reports</td></tr>
      <tr><td><kbd>&#8593;</kbd> <kbd>&#8595;</kbd> <kbd>Enter</kbd></td><td>In the project box: pick from the list, or type a new name and press Enter</td></tr>
      <tr><td><kbd>Alt + &#8593;</kbd> / <kbd>Alt + &#8595;</kbd></td><td>On a focused task: move it up or down</td></tr>
      <tr><td><kbd>Ctrl + V</kbd></td><td>In a task's files: paste an image</td></tr>
      <tr><td><kbd>Ctrl + Z</kbd></td><td>Undo the last delete, file or clear (outside a text box)</td></tr>
      <tr><td><kbd id="helpHotkey">Ctrl + Alt + T</kbd></td><td>Quick-add a task from anywhere (tray)</td></tr>
      <tr><td><kbd>Esc</kbd></td><td>Close a dialog or the history panel — asks first if there are unsaved changes</td></tr>
    </table>
    <p class="hint">The terminal console, when there is one, is live at the same time - a report filed there shows up here.</p>
    <div class="modal-actions" style="justify-content: flex-end;"><button class="btn btn-secondary" data-close>Close</button></div>
  </div>
</div>

<div class="backdrop" id="confirmBackdrop">
  <div class="modal modal-sm" role="alertdialog" aria-labelledby="confirmTitle" aria-describedby="confirmBody">
    <span class="dialog-title" id="confirmTitle">Save changes?</span>
    <span class="dialog-body" id="confirmBody"></span>
    <div class="modal-actions confirm-actions" style="margin-top: 4px;">
      <button class="btn btn-primary" id="confirmSave" style="min-width: 140px;">Save changes</button>
      <button class="btn btn-secondary" id="confirmKeep">Keep editing</button>
      <button class="btn btn-ghost" id="confirmDiscard">Discard</button>
    </div>
  </div>
</div>

<div class="toast hidden" id="toast" role="status" aria-live="polite">
  <span class="toast-msg" id="toastMsg"></span>
  <button class="toast-undo hidden" id="toastUndo">Undo</button>
  <button class="toast-x" id="toastClose" aria-label="Dismiss"><i data-icon="x" data-size="14" data-stroke="2"></i></button>
</div>

<!-- Quick add: the same page in the small tray window -->
<div class="quick" id="quickView">
  <div class="quick-bar">
    <div class="drag pywebview-drag-region" id="quickDrag">
      <img src="/favicon.ico" alt="">
      <span>Quick add task</span>
    </div>
    <button class="winbtn close" id="qClose" aria-label="Close"><i data-icon="x" data-size="14" data-stroke="1.75"></i></button>
  </div>
  <div class="quick-body" id="quickBody">
    <div class="quick-row">
      <input class="input" type="date" id="qDate">
      <div class="combo">
        <input class="input" id="qProject" placeholder="project" spellcheck="false" autocomplete="off" maxlength="%%MAXPROJECT%%"
               role="combobox" aria-expanded="false" aria-autocomplete="list" aria-controls="qProjectPanel">
        <button type="button" class="combo-caret" id="qProjectCaret" tabindex="-1" aria-label="Show existing projects"><i data-icon="chevron" data-size="16" data-stroke="2"></i></button>
        <div class="combo-panel" id="qProjectPanel" role="listbox"></div>
      </div>
    </div>
    <textarea class="input" id="qText" maxlength="%%MAXTASK%%" placeholder="What needs doing…   Shift + Enter for a new line"></textarea>
    <div class="quick-files" id="qFiles"></div>
    <div class="quick-attach" id="qAttach">
      <i data-icon="paperclip" data-size="14" data-stroke="2"></i>
      <span style="flex: 1;">Paste or drop a file to attach it</span>
    </div>
    <div class="quick-msg" id="qMsg"></div>
    <div class="quick-foot">
      <span class="hint"><b>Enter</b> add · <b>Esc</b> close</span>
      <button class="btn btn-secondary" id="qOpenApp">Open app</button>
      <button class="btn btn-primary" id="qAdd" style="min-width: 110px;">Add Task</button>
    </div>
  </div>
</div>

<div id="gone"><div id="goneText"></div></div>

<script>
"use strict";
const CFG = window.CFG;
const IS_APP = Boolean(CFG.app);
const IS_QUICK = CFG.mode === "quick";

const $ = (id) => document.getElementById(id);
const editor = $("editor");
const statusEl = $("status");
const counterEl = $("counter");
const fillEl = $("progressFill");
const saveBtn = $("saveBtn");

function el(tag, cls, text) {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  if (text !== undefined) node.textContent = text;
  return node;
}

/* -------------------------------------------------------------------- icons */

// Lucide, drawn inline: the page runs offline, so nothing is fetched.
const ICONS = {
  folder: '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
  history: '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l4 2"/>',
  moon: '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/>',
  help: '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><path d="M12 17h.01"/>',
  settings: '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
  pencil: '<path d="M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z"/>',
  trash: '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>',
  copy: '<rect width="14" height="14" x="8" y="8"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
  grip: '<circle cx="9" cy="5" r="1.6" fill="currentColor" stroke="none"/><circle cx="9" cy="12" r="1.6" fill="currentColor" stroke="none"/><circle cx="9" cy="19" r="1.6" fill="currentColor" stroke="none"/><circle cx="15" cy="5" r="1.6" fill="currentColor" stroke="none"/><circle cx="15" cy="12" r="1.6" fill="currentColor" stroke="none"/><circle cx="15" cy="19" r="1.6" fill="currentColor" stroke="none"/>',
  paperclip: '<path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/>',
  x: '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
  check: '<path d="M20 6 9 17l-5-5"/>',
  plus: '<path d="M5 12h14"/><path d="M12 5v14"/>',
  minus: '<path d="M5 12h14"/>',
  square: '<rect width="18" height="18" x="3" y="3"/>',
  restore: '<rect width="14" height="14" x="3" y="7"/><path d="M7 7V3h14v14h-4"/>',
  chevron: '<path d="m6 9 6 6 6-6"/>',
  search: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
  alert: '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
  upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>',
  image: '<rect width="18" height="18" x="3" y="3"/><circle cx="9" cy="9" r="2"/><path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/>',
};

function icon(name, size, stroke) {
  const holder = document.createElement("span");
  // Only ever the fixed strings above - never anything a user typed.
  holder.innerHTML =
    '<svg width="' + (size || 18) + '" height="' + (size || 18) + '" viewBox="0 0 24 24" fill="none" ' +
    'stroke="currentColor" stroke-width="' + (stroke || 1.75) + '" stroke-linecap="round" ' +
    'stroke-linejoin="round" aria-hidden="true">' + (ICONS[name] || "") + "</svg>";
  return holder.firstChild;
}

function drawIcons(root) {
  root.querySelectorAll("i[data-icon]").forEach((slot) => {
    slot.replaceWith(icon(slot.dataset.icon, Number(slot.dataset.size) || 18, Number(slot.dataset.stroke) || 1.75));
  });
}
drawIcons(document);

/* ---------------------------------------------------------------- transport */

// Every request carries the session token.  The server refuses anything
// without it, which matters on WSL: the port is reachable from Windows, so
// "bound to loopback" is not on its own a closed door.
function withToken(path) {
  return path + (path.includes("?") ? "&" : "?") + "t=" + encodeURIComponent(CFG.token);
}

async function api(path, body) {
  const res = await fetch(withToken(path), {
    method: body === undefined ? "GET" : "POST",
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
    cache: "no-store",
  });
  if (!res.ok) throw new Error("HTTP " + res.status);
  return res.json();
}

// A file goes up as itself, not as base64 inside JSON.
async function uploadTaskFile(taskId, file, name) {
  const res = await fetch(
    withToken("/api/tasks/files/add?id=" + encodeURIComponent(taskId) + "&name=" + encodeURIComponent(name || file.name)),
    { method: "POST", headers: { "Content-Type": "application/octet-stream" }, body: file, cache: "no-store" }
  );
  if (!res.ok) throw new Error("HTTP " + res.status);
  return res.json();
}

function fileUrl(taskId, name) {
  return withToken("/api/tasks/files/raw?id=" + encodeURIComponent(taskId) + "&name=" + encodeURIComponent(name));
}

// The desktop window's own controls - minimise, drag, the folder picker - are
// Python methods pywebview exposes.  In a plain browser they simply are not
// there, and every caller copes with null.
function native() {
  return window.pywebview && window.pywebview.api ? window.pywebview.api : null;
}

async function callNative(name, ...args) {
  const bridge = native();
  if (!bridge || typeof bridge[name] !== "function") return null;
  try { return await bridge[name](...args); } catch (err) { return null; }
}

/* ------------------------------------------------------------------- status */

let statusTimer = null;

function setStatus(text, kind) {
  if (statusTimer) { clearTimeout(statusTimer); statusTimer = null; }
  statusEl.textContent = text;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

// A transient message that decays back to whatever the counter thinks.
function flashStatus(text, kind) {
  setStatus(text, kind);
  statusTimer = setTimeout(() => { statusTimer = null; updateCounter(); }, 6000);
}

let boardStatusTimer = null;

function setBoardStatus(text, kind) {
  if (boardStatusTimer) { clearTimeout(boardStatusTimer); boardStatusTimer = null; }
  const node = $("boardStatus");
  node.textContent = text;
  node.className = "status" + (kind ? " " + kind : "");
}

function flashBoard(text, kind) {
  setBoardStatus(text, kind);
  boardStatusTimer = setTimeout(() => { boardStatusTimer = null; setBoardStatus("Ready", null); }, 6000);
}

// Whichever view is on screen owns the status line.
function flashHere(text, kind) {
  if (currentView === "board") flashBoard(text, kind);
  else if (currentView === "report") flashStatus(text, kind);
  else showToast(text);
}

/* -------------------------------------------------------------------- toast */

// One toast at a time, bottom left.  An undo is only ever offered for the
// last thing done, which is also what Ctrl+Z takes back.
let toastTimer = null;
let toastUndo = null;

function showToast(message, undo) {
  if (toastTimer) clearTimeout(toastTimer);
  toastUndo = undo || null;
  $("toastMsg").textContent = message;
  $("toastUndo").classList.toggle("hidden", !toastUndo);
  $("toast").classList.remove("hidden");
  toastTimer = setTimeout(hideToast, 6000);
}

function hideToast() {
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = null;
  toastUndo = null;
  $("toast").classList.add("hidden");
}

async function runToastUndo() {
  const undo = toastUndo;
  hideToast();
  if (undo) await undo();
}

$("toastUndo").addEventListener("click", runToastUndo);
$("toastClose").addEventListener("click", hideToast);

function shorten(text, length) {
  const flat = String(text || "").replace(/\s+/g, " ").trim();
  return flat.length > length ? flat.slice(0, length - 1) + "…" : flat;
}

/* ---------------------------------------------------------------- clipboard */

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (err) {
    // The old way still works where the async clipboard is not allowed.
    const area = el("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    let ok = false;
    try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
    area.remove();
    return ok;
  }
}

// Images go onto the clipboard as PNG, the one format every paste target
// (an AI chat, Paint, Word) accepts.
async function copyImageBlob(blob) {
  if (!window.ClipboardItem || !navigator.clipboard || !navigator.clipboard.write) return false;
  let png = blob;
  if (blob.type !== "image/png") {
    const bitmap = await createImageBitmap(blob);
    const canvas = document.createElement("canvas");
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext("2d").drawImage(bitmap, 0, 0);
    png = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  }
  try {
    await navigator.clipboard.write([new ClipboardItem({ "image/png": png })]);
    return true;
  } catch (err) {
    return false;
  }
}

/* ------------------------------------------------------------------ dialogs */

// Each closable layer, top-most last.  The history panel is in here too, so
// Esc always closes whatever is in front.
const openStack = [];
// Dialogs that hold edits register how to tell, how to save and how to drop
// them.  Closing one of those while it is dirty asks first.
const guarded = {};

function openModal(id) {
  $(id).classList.add("open");
  if (!openStack.includes(id)) openStack.push(id);
}

function closeModal(id) {
  if (id === "historyPanel") { closeHistory(); return; }
  $(id).classList.remove("open");
  const at = openStack.indexOf(id);
  if (at !== -1) openStack.splice(at, 1);
  if (guarded[id]) { guarded[id].onClosed(); reportDirty(); }
  if (!openStack.length) focusView();
}

function topLayer() {
  return openStack[openStack.length - 1];
}

// The polite close: a dirty dialog asks "Save changes?" instead of going.
async function requestClose(id) {
  const guard = guarded[id];
  if (!guard || !guard.isDirty()) { closeModal(id); return true; }
  const answer = await askSaveChanges(guard.confirmText);
  if (answer === "save") return guard.save();
  if (answer === "discard") { guard.discard(); closeModal(id); return true; }
  return false;
}

document.querySelectorAll("[data-close]").forEach((btn) => {
  btn.addEventListener("click", () => {
    const backdrop = btn.closest(".backdrop");
    if (backdrop) requestClose(backdrop.id);
  });
});

document.querySelectorAll(".backdrop").forEach((backdrop) => {
  // mousedown and click both on the backdrop itself: a drag that starts in a
  // text box and ends outside the dialog must not count as "clicked outside".
  let downOnBackdrop = false;
  backdrop.addEventListener("mousedown", (event) => { downOnBackdrop = event.target === backdrop; });
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop && downOnBackdrop && backdrop.id !== "confirmBackdrop") {
      requestClose(backdrop.id);
    }
    downOnBackdrop = false;
  });
});

/* ------------------------------------------------- the "Save changes?" prompt */

let confirmResolve = null;

function askSaveChanges(body) {
  if (confirmResolve) return Promise.resolve("keep");
  $("confirmBody").textContent = body || "You changed this. If you close now, those changes are lost.";
  openModal("confirmBackdrop");
  // Keep editing is the default: Enter or Esc never loses anything.
  $("confirmKeep").focus();
  return new Promise((resolve) => { confirmResolve = resolve; });
}

function answerConfirm(answer) {
  const resolve = confirmResolve;
  confirmResolve = null;
  $("confirmBackdrop").classList.remove("open");
  const at = openStack.indexOf("confirmBackdrop");
  if (at !== -1) openStack.splice(at, 1);
  if (resolve) resolve(answer);
}

$("confirmSave").addEventListener("click", () => answerConfirm("save"));
$("confirmKeep").addEventListener("click", () => answerConfirm("keep"));
$("confirmDiscard").addEventListener("click", () => answerConfirm("discard"));

function anyDirty() {
  return Object.keys(guarded).some((id) => openStack.includes(id) && guarded[id].isDirty());
}

// The window's close button lives in Python; it needs to know whether
// closing now would lose an edit.
let lastDirtyReport = null;
function reportDirty() {
  // Only the main window has edits worth asking about; the quick-add window
  // shares the session and must not overwrite what the main one said.
  if (IS_QUICK) return;
  const dirty = anyDirty();
  if (dirty === lastDirtyReport) return;
  lastDirtyReport = dirty;
  api("/api/page-state", { dirty: dirty }).catch(() => { lastDirtyReport = null; });
}

/* -------------------------------------------------------------------- drafts */

// Everything typed is kept in localStorage as it is typed, so a crash, a
// closed window or a reload never costs more than the last 300 ms.
const DRAFT_KEY = IS_QUICK ? "taskReporterQuickDraft" : "taskReporterDrafts";
let drafts = {};
try { drafts = JSON.parse(localStorage.getItem(DRAFT_KEY) || "{}") || {}; } catch (err) { drafts = {}; }
let draftTimer = null;

function writeDrafts() {
  try { localStorage.setItem(DRAFT_KEY, JSON.stringify(drafts)); } catch (err) { /* private mode */ }
}

function setDraft(key, value) {
  if (value === null || value === undefined || value === "") delete drafts[key];
  else drafts[key] = value;
  if (draftTimer) clearTimeout(draftTimer);
  draftTimer = setTimeout(() => { draftTimer = null; writeDrafts(); }, 300);
}

function clearDraft(key) {
  delete drafts[key];
  if (draftTimer) clearTimeout(draftTimer);
  draftTimer = null;
  writeDrafts();
}

window.addEventListener("pagehide", () => { if (draftTimer) { clearTimeout(draftTimer); writeDrafts(); } });

/* ------------------------------------------------------------------- report */

function updateCounter() {
  const length = editor.value.length;
  const over = length > CFG.maxLength;
  counterEl.textContent = length + " / " + CFG.maxLength;
  counterEl.className = "counter" + (over ? " danger" : "");
  fillEl.style.width = Math.min(100, (length / CFG.maxLength) * 100) + "%";
  fillEl.className = "progress-fill" + (over ? " danger" : "");
  saveBtn.disabled = saving || !editor.value.trim() || over;
  if (statusTimer) return;
  if (over) setStatus("Report is too long  (" + (length - CFG.maxLength) + " chars over limit)", "danger");
  else setStatus("Ready", null);
}

let saving = false;

async function saveReport() {
  if (saving) return;
  const text = editor.value.trim();
  if (!text) { flashStatus("Report cannot be empty", "danger"); editor.focus(); return; }
  if (text.length > CFG.maxLength) {
    flashStatus("Please keep report under " + CFG.maxLength + " characters", "danger");
    editor.focus();
    return;
  }

  saving = true;
  saveBtn.disabled = true;
  setStatus("Saving…", null);
  try {
    const out = await api("/api/save", { text });
    if (out.ok) {
      editor.value = "";
      clearDraft("report");
      flashStatus(
        out.queued
          ? "Queued at " + out.timestamp + " - Excel has the workbook open"
          : "Saved at " + out.timestamp + " ✓",
        out.queued ? "warn" : "success"
      );
      if (out.queued) refreshBanner(true, (lastPing.pendingCount || 0) + 1);
    } else {
      flashStatus(out.message || "Could not save the report", "danger");
    }
  } catch (err) {
    // The report is still in the box, so nothing is lost by retrying.
    flashStatus("Lost contact with the reporter - your text is still here", "danger");
  } finally {
    saving = false;
    updateCounter();
    editor.focus();
  }
}

/* ------------------------------------------------------- previous reports */

let historyRows = [];
let historyQueued = [];
let historyOpen = false;

function openHistory() {
  if (historyOpen) { $("historySearch").focus(); return; }
  historyOpen = true;
  document.body.classList.add("panel-open");
  if (!openStack.includes("historyPanel")) openStack.push("historyPanel");
  loadHistory();
  setTimeout(() => $("historySearch").focus(), 30);
}

function closeHistory() {
  historyOpen = false;
  document.body.classList.remove("panel-open");
  const at = openStack.indexOf("historyPanel");
  if (at !== -1) openStack.splice(at, 1);
  if (!openStack.length) focusView();
}

async function loadHistory() {
  try {
    const out = await api("/api/list");
    historyRows = out.reports || [];
    historyQueued = out.queued || [];
  } catch (err) {
    const body = $("historyBody");
    body.innerHTML = "";
    body.appendChild(el("div", "empty", "Could not read the workbook."));
    return;
  }
  fillHistoryProjects();
  renderHistory();
}

function reportProjects(text) {
  const found = [];
  const pattern = /^\[([^\]\n]+)\]\s*$/gm;
  let match;
  while ((match = pattern.exec(String(text || "")))) found.push(match[1].trim());
  return found;
}

function fillHistoryProjects() {
  const select = $("historyProject");
  const keep = select.value || "all";
  const names = [];
  const seen = new Set();
  const add = (name) => {
    const key = name.toLowerCase();
    if (name && !seen.has(key)) { seen.add(key); names.push(name); }
  };
  boardProjects.forEach(add);
  historyRows.concat(historyQueued).forEach((row) => reportProjects(row.text).forEach(add));
  select.innerHTML = "";
  const all = el("option", null, "All projects");
  all.value = "all";
  select.appendChild(all);
  for (const name of names) {
    const option = el("option", null, name);
    option.value = name;
    select.appendChild(option);
  }
  select.value = names.includes(keep) ? keep : "all";
}

function historyMatches(row) {
  const project = $("historyProject").value;
  if (project && project !== "all" && !String(row.text).includes("[" + project + "]")) return false;
  const needle = $("historySearch").value.trim().toLowerCase();
  if (!needle) return true;
  return String(row.text).toLowerCase().includes(needle) || String(row.datetime).toLowerCase().includes(needle);
}

// The search term, marked in the text it was found in.  Built from text
// nodes, never innerHTML: report text is whatever anyone typed.
function highlighted(text, needle) {
  const node = el("div", "h-text");
  const source = String(text || "");
  if (!needle) { node.textContent = source; return node; }
  const lower = source.toLowerCase();
  let from = 0;
  let at;
  while ((at = lower.indexOf(needle, from)) !== -1) {
    node.appendChild(document.createTextNode(source.slice(from, at)));
    node.appendChild(el("mark", null, source.slice(at, at + needle.length)));
    from = at + needle.length;
  }
  node.appendChild(document.createTextNode(source.slice(from)));
  return node;
}

function miniButton(name, title, danger, onClick) {
  const button = el("button", "mini" + (danger ? " danger" : ""));
  button.type = "button";
  button.title = title;
  button.setAttribute("aria-label", title);
  button.appendChild(icon(name, 13, 2));
  button.addEventListener("click", onClick);
  return button;
}

function renderHistory() {
  const body = $("historyBody");
  const needle = $("historySearch").value.trim().toLowerCase();
  body.innerHTML = "";

  const queued = historyQueued.filter(historyMatches);
  const rows = historyRows.filter(historyMatches);
  const total = historyRows.length + historyQueued.length;
  $("historyCount").textContent = (rows.length + queued.length) + " of " + total + " rows";

  if (queued.length) {
    body.appendChild(el("div", "queued-strip", "Queued — not in the workbook yet"));
    for (const item of queued) {
      const row = el("div", "h-row");
      row.appendChild(el("span", "h-num", ""));
      const main = el("div", "h-main");
      const when = el("span", "h-when", item.datetime);
      when.appendChild(el("span", "tag tag-accent tag-small", "Queued"));
      main.appendChild(when);
      main.appendChild(highlighted(item.text, needle));
      row.appendChild(main);
      const acts = el("div", "h-acts");
      acts.appendChild(miniButton("copy", "Copy report text", false, () => copyReport(item)));
      row.appendChild(acts);
      body.appendChild(row);
    }
  }

  if (!rows.length && !queued.length) {
    body.appendChild(el("div", "empty", total ? "No reports match." : "No reports yet."));
    return;
  }

  // Newest first, but numbered by their real position in the workbook.
  for (let i = rows.length - 1; i >= 0; i--) {
    const item = rows[i];
    const row = el("div", "h-row");
    row.appendChild(el("span", "h-num", String(item.index + 1)));
    const main = el("div", "h-main");
    main.appendChild(el("span", "h-when", item.datetime));
    main.appendChild(highlighted(item.text, needle));
    row.appendChild(main);
    const acts = el("div", "h-acts");
    acts.appendChild(miniButton("copy", "Copy report text", false, () => copyReport(item)));
    acts.appendChild(miniButton("pencil", "Edit", false, () => openEdit(item.index)));
    acts.appendChild(miniButton("trash", "Delete", true, () => deleteReport(item.index)));
    row.appendChild(acts);
    body.appendChild(row);
  }
}

async function copyReport(item) {
  const ok = await copyText(item.text);
  showToast(ok ? "Report " + String(item.datetime).slice(0, 10) + " copied" : "Could not reach the clipboard");
}

async function deleteReport(index) {
  const item = historyRows.find((row) => row.index === index);
  if (!item) return;
  let out;
  try {
    out = await api("/api/delete", { index });
  } catch (err) {
    showToast("Lost contact with the reporter. Nothing was deleted.");
    return;
  }
  if (!out.ok) { showToast(out.message || "Could not delete the report."); return; }
  await loadHistory();
  // The row is kept here, so Undo can put it back exactly where it was.
  const kept = { index: item.index, datetime: item.datetime, text: item.text };
  showToast("Report " + String(item.datetime).slice(0, 10) + " deleted", async () => {
    try {
      const back = await api("/api/restore", kept);
      if (!back.ok) { showToast(back.message || "Could not restore the report."); return; }
      showToast("Report " + String(kept.datetime).slice(0, 10) + " restored");
    } catch (err) {
      showToast("Lost contact with the reporter. The report was not restored.");
    }
    if (historyOpen) loadHistory();
  });
}

$("historyBtn").addEventListener("click", () => (historyOpen ? closeHistory() : openHistory()));
$("historyClose").addEventListener("click", closeHistory);
$("historyScrim").addEventListener("click", closeHistory);
$("historySearch").addEventListener("input", renderHistory);
$("historyProject").addEventListener("change", renderHistory);
$("viewQueuedBtn").addEventListener("click", openHistory);

/* -------------------------------------------------------------- report edit */

let editIndex = null;
let editOriginal = null;

function editSnapshot() {
  return JSON.stringify({ when: $("editWhen").value, text: $("editText").value });
}

function openEdit(index, draft) {
  const item = historyRows.find((row) => row.index === index);
  if (!item) return false;
  editIndex = index;
  $("editWhen").value = item.datetime;
  $("editText").value = item.text;
  editOriginal = editSnapshot();
  if (draft) {
    $("editWhen").value = draft.when;
    $("editText").value = draft.text;
  }
  $("editError").textContent = "";
  openModal("editBackdrop");
  editChanged();
  $("editText").focus();
  return true;
}

function editChanged() {
  const dirty = guarded.editBackdrop.isDirty();
  $("editDirty").classList.toggle("hidden", !dirty);
  setDraft("edit", dirty && editIndex !== null
    ? { kind: "report", index: editIndex, origWhen: JSON.parse(editOriginal).when,
        origText: JSON.parse(editOriginal).text, when: $("editWhen").value, text: $("editText").value }
    : null);
  reportDirty();
}

async function saveEdit() {
  if (editIndex === null) return false;
  const text = $("editText").value.trim();
  const when = $("editWhen").value.trim();
  const error = $("editError");

  if (!text) { error.textContent = "Report text cannot be empty."; return false; }
  if (text.length > CFG.maxLength) {
    error.textContent = "Report is too long (" + text.length + " chars). Limit is " + CFG.maxLength + ".";
    return false;
  }

  error.textContent = "";
  try {
    const out = await api("/api/update", { index: editIndex, datetime: when, text });
    if (!out.ok) { error.textContent = out.message || "Could not save the changes."; return false; }
  } catch (err) {
    error.textContent = "Lost contact with the reporter. Nothing was changed.";
    return false;
  }
  editOriginal = editSnapshot();
  closeModal("editBackdrop");
  if (historyOpen) await loadHistory();
  showToast("Report saved to the workbook");
  return true;
}

guarded.editBackdrop = {
  confirmText: "You changed this report. If you close now, those changes are lost.",
  isDirty: () => editIndex !== null && editOriginal !== null && editSnapshot() !== editOriginal,
  save: saveEdit,
  discard: () => { editOriginal = editSnapshot(); },
  onClosed: () => { editIndex = null; editOriginal = null; $("editDirty").classList.add("hidden"); clearDraft("edit"); },
};

$("editWhen").addEventListener("input", editChanged);
$("editText").addEventListener("input", editChanged);
$("editSave").addEventListener("click", saveEdit);

/* --------------------------------------------------------- session heartbeat */

// The page tells the server it is alive, and says goodbye on the way out, so
// that closing the tab ends the session the way closing a window used to.
// Missed pings alone never end it: browsers throttle timers in background
// tabs, and a false "the browser is gone" would take the terminal with it.
let ended = false;
let lastPing = { workbookOpen: false, pendingCount: 0 };

function endSession(reason) {
  if (ended) return;
  ended = true;
  saveBtn.disabled = true;
  editor.readOnly = true;
  $("goneText").textContent =
    "Task Reporter session ended" + (reason ? " - " + reason : "") + ".\nYou can close this window.";
  $("gone").classList.add("open");
}

function refreshBanner(open, pending) {
  const banner = $("excelBanner");
  const show = Boolean(open) || pending > 0;
  banner.classList.toggle("hidden", !show || IS_QUICK);
  const reports = pending + (pending === 1 ? " report is" : " reports are");
  const text = $("excelBannerText");
  text.innerHTML = "";
  if (open) {
    text.appendChild(el("b", null, "task_reports.xlsx is open in Excel. "));
    text.appendChild(document.createTextNode(
      pending
        ? reports + " queued and will be written in as soon as Excel closes it."
        : "Anything saved now is queued and written in as soon as Excel closes it."
    ));
  } else {
    text.appendChild(el("b", null, reports + " queued. "));
    text.appendChild(document.createTextNode("They will be written into task_reports.xlsx in a moment."));
  }
  $("viewQueuedBtn").classList.toggle("hidden", !pending);
  // Settings shows the same thing as a tag.
  const tag = $("sExcelTag");
  tag.textContent = open ? "OPEN IN EXCEL" : "WORKBOOK FREE";
  tag.className = "tag " + (open ? "tag-accent" : "tag-neutral");
  $("sQueueNote").textContent = pending
    ? reports + " waiting in .task_reports_pending.jsonl"
    : "Nothing queued";
}

async function ping() {
  if (ended) return;
  try {
    const out = await api("/api/ping", { client: CFG.clientId });
    if (out.shuttingDown) { endSession(out.reason); return; }
    for (const event of out.events || []) {
      flashStatus(
        "Filed from the terminal at " + event.timestamp + (event.queued ? " (queued - workbook is locked)" : ""),
        event.queued ? "warn" : "success"
      );
    }
    const pending = out.pendingCount !== undefined ? out.pendingCount : out.queued || 0;
    if (out.workbookOpen !== lastPing.workbookOpen || pending !== lastPing.pendingCount) {
      const dropped = lastPing.pendingCount > 0 && pending === 0;
      lastPing = { workbookOpen: Boolean(out.workbookOpen), pendingCount: pending };
      refreshBanner(lastPing.workbookOpen, pending);
      if (dropped && historyOpen) loadHistory();
    }
    for (const command of out.commands || []) runCommand(command);
    // The terminal and the tray window can add, tick and file tasks too.  The
    // revision changes on every write, so this is also what fills the Board
    // badge in on start-up without the board having been opened.
    if (out.boardRevision !== undefined && out.boardRevision !== boardRevision) loadBoard();
  } catch (err) {
    // A single miss means nothing - the server may just be busy saving.
  }
}

// Things Python asks the page to do: the tray menu, mostly.
function runCommand(command) {
  if (command === "history") { if (!IS_QUICK) openHistory(); }
  else if (command === "board" || command === "report") { if (!IS_QUICK) setView(command); }
  else if (command === "requestClose") requestWindowClose();
  else if (command === "quickShow") quickShown();
}

window.addEventListener("pagehide", (event) => {
  // persisted means the page is going into the back/forward cache and will be
  // reused - it has not been closed, so it must not say goodbye.
  if (ended || event.persisted) return;
  const url = withToken("/api/bye?client=" + encodeURIComponent(CFG.clientId));
  // sendBeacon survives teardown; fetch(keepalive) is the fallback.
  if (!(navigator.sendBeacon && navigator.sendBeacon(url))) {
    fetch(url, { method: "POST", keepalive: true }).catch(() => {});
  }
});

// A plain browser tab gets the browser's own "leave site?" check.
window.addEventListener("beforeunload", (event) => {
  if (!IS_APP && anyDirty()) { event.preventDefault(); event.returnValue = ""; }
});

window.addEventListener("pageshow", () => { ping(); });
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") ping();
});

/* --------------------------------------------------------------- task board */

// The board is the same shape as the hand-kept task list it replaces: a day,
// the projects worked on that day, and the tasks under each one with a box to
// tick.  Ticking is what matters - "File Checked Tasks" turns every ticked task
// into report rows, one row per day, grouped under [project] headers.

const boardScroll = $("boardScroll");
const tabBadge = $("tabBadge");

let boardTasks = [];
let boardCounts = { total: 0, done: 0, open: 0, filed: 0, ready: 0 };
// What the server last said the board looked like.  null means "never read",
// so the first ping pulls the board in even if the Report view is showing.
let boardRevision = null;
let boardLoaded = false;
let showFiledTasks = false;
let projectFilter = "all";
let previewGroups = [];
let currentView = "report";
try { projectFilter = localStorage.getItem("taskReporterFilter") || "all"; } catch (err) { /* private mode */ }

/* --------------------------------------------------------------------- dates */

const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function pad2(n) { return String(n).padStart(2, "0"); }

function isoToday() {
  const now = new Date();
  return now.getFullYear() + "-" + pad2(now.getMonth() + 1) + "-" + pad2(now.getDate());
}

function isoToDisplay(iso) {
  const parts = String(iso || "").split("-");
  return parts.length === 3 ? parts[2] + "." + parts[1] + "." + parts[0] : String(iso || "");
}

function stampNow() {
  const now = new Date();
  return pad2(now.getDate()) + "/" + pad2(now.getMonth() + 1) + "/" + now.getFullYear() + " " +
    pad2(now.getHours()) + ":" + pad2(now.getMinutes()) + ":" + pad2(now.getSeconds());
}

// "Today", "Yesterday", or the weekday and how far back it was.  Built from
// local Y/M/D parts rather than Date.parse, which reads a bare ISO date as UTC
// midnight and can land on the wrong day west of Greenwich.
function relativeDay(iso) {
  const parts = String(iso || "").split("-").map(Number);
  if (parts.length !== 3 || parts.some((n) => !isFinite(n))) return "";
  const day = new Date(parts[0], parts[1] - 1, parts[2]);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diff = Math.round((day - today) / 86400000);
  if (diff === 0) return "Today";
  if (diff === -1) return "Yesterday";
  if (diff === 1) return "Tomorrow";
  const name = WEEKDAYS[day.getDay()];
  if (diff < 0) return name + " · " + (-diff) + " days ago";
  return name + " · in " + diff + " days";
}

// "21/08/2026 15:04:12" -> "15:04"
function clockOf(stamp) {
  const bits = String(stamp || "").split(" ");
  return bits.length > 1 ? bits[1].slice(0, 5) : "";
}

/* --------------------------------------------------------------- board load */

async function loadBoard() {
  let out;
  try {
    out = await api("/api/tasks");
  } catch (err) {
    if (!IS_QUICK) {
      boardScroll.innerHTML = "";
      boardScroll.appendChild(el("div", "board-empty", "Could not read the task board."));
    }
    return;
  }
  boardTasks = (out.tasks || []).map((task) => Object.assign({ files: [] }, task));
  boardCounts = out.counts || boardCounts;
  boardRevision = out.revision;
  setProjectOptions(out.projects || []);
  if (!IS_QUICK) renderBoard();
  if (!boardLoaded) { boardLoaded = true; restoreEditDraft(); }
}

/* --------------------------------------------------------- project dropdown */

// A real dropdown rather than a <datalist>: browsers only reveal datalist
// suggestions once you have typed a matching prefix, which is no use when the
// whole point is seeing which projects already exist.  Typing a name that is
// not in the list and pressing Enter is how a new project gets made - there is
// no separate "create project" step.

let boardProjects = [];
const projectCombos = [];

function attachProjectCombo(input, panel, caret, onCommit) {
  // `typed` separates "the user is narrowing the list" from "this box already
  // held a value" - opening the dropdown on a task that is already tagged
  // [Logistics] has to show every project, not just that one.
  const combo = { input: input, open: false, active: -1, items: [], typed: false };

  function setOpen(state) {
    combo.open = state;
    panel.classList.toggle("open", state);
    input.setAttribute("aria-expanded", state ? "true" : "false");
  }

  function close() { combo.active = -1; setOpen(false); }

  function commit(name) {
    input.value = name;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    close();
    if (onCommit) onCommit();
  }

  function render() {
    const typed = input.value.trim();
    const needle = combo.typed ? typed.toLowerCase() : "";
    combo.items = boardProjects.filter((name) => !needle || name.toLowerCase().includes(needle));
    panel.innerHTML = "";

    combo.items.forEach((name, index) => {
      const row = el("div", "combo-option" + (index === combo.active ? " is-active" : ""));
      row.setAttribute("role", "option");
      row.setAttribute("aria-selected", index === combo.active ? "true" : "false");
      const label = el("button", "combo-name", name);
      label.type = "button";
      label.tabIndex = -1;
      // mousedown, and preventDefault: the input never loses focus, so its
      // blur handler cannot close the panel out from under the click.
      label.addEventListener("mousedown", (event) => { event.preventDefault(); commit(name); });
      row.appendChild(label);
      const forget = el("button", "combo-forget");
      forget.appendChild(icon("x", 12, 2.5));
      forget.type = "button";
      forget.tabIndex = -1;
      forget.title = "Forget this name (tasks already using it keep it)";
      forget.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
        forgetProject(name);
      });
      row.appendChild(forget);
      panel.appendChild(row);
    });

    const exact = boardProjects.some((name) => name.toLowerCase() === typed.toLowerCase());
    const hint = el("div", "combo-hint");
    if (!boardProjects.length) {
      hint.append("No projects yet - type a name and press ", el("b", null, "Enter"), ".");
    } else if (typed && !exact && combo.typed) {
      hint.append("Press ", el("b", null, "Enter"), " to create ", el("b", null, "[" + typed + "]"));
    } else {
      hint.append("Pick with ", el("b", null, "↑ ↓"), " and ", el("b", null, "Enter"), ", or type a new name.");
    }
    if (combo.items.length) panel.appendChild(el("div", "combo-sep"));
    panel.appendChild(hint);
  }

  function open() { combo.active = -1; combo.typed = false; render(); setOpen(true); }

  function move(step) {
    if (!combo.open) { open(); return; }
    const count = combo.items.length;
    if (!count) return;
    // active === -1 is the "keep what I typed" stop, so there are count + 1 of
    // them and the arrows wrap through it rather than sticking at the ends.
    const span = count + 1;
    let slot = combo.active + 1 + step;
    slot = ((slot % span) + span) % span;
    combo.active = slot - 1;
    render();
    const active = panel.querySelector(".combo-option.is-active");
    if (active) active.scrollIntoView({ block: "nearest" });
  }

  input.addEventListener("focus", open);
  input.addEventListener("mousedown", () => { if (!combo.open) open(); });
  input.addEventListener("input", (event) => {
    if (!event.isTrusted) return;
    combo.typed = true;
    combo.active = -1;
    if (!combo.open) setOpen(true);
    render();
  });
  caret.addEventListener("mousedown", (event) => {
    event.preventDefault();
    if (combo.open) { close(); return; }
    input.focus();
    if (!combo.open) open();
  });
  input.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown") { event.preventDefault(); move(1); return; }
    if (event.key === "ArrowUp") { event.preventDefault(); move(-1); return; }
    if (event.key === "Escape") {
      // Only swallow Escape while the panel is up, so it still closes a dialog.
      if (!combo.open) return;
      event.preventDefault();
      event.stopPropagation();
      close();
      return;
    }
    if (event.key === "Enter" && !event.ctrlKey && !event.metaKey) {
      // Never submits the composer: Enter here means "this is the project".
      event.preventDefault();
      if (combo.open && combo.active >= 0) commit(combo.items[combo.active]);
      else { close(); if (onCommit) onCommit(); }
      return;
    }
    if (event.key === "Tab") close();
  });
  input.addEventListener("blur", close);

  combo.close = close;
  combo.refresh = () => { if (combo.open) render(); };
  projectCombos.push(combo);
  return combo;
}

function setProjectOptions(projects) {
  boardProjects = projects || [];
  for (const combo of projectCombos) combo.refresh();
}

async function forgetProject(name) {
  try {
    const out = await api("/api/projects/forget", { name: name });
    if (!out.ok) { flashBoard(out.message || "Could not forget that name", "danger"); return; }
  } catch (err) {
    flashBoard("Lost contact with the reporter - nothing was changed", "danger");
    return;
  }
  // Updated in place rather than by reloading the board, so an open panel does
  // not blink shut while it is being tidied.
  setProjectOptions(boardProjects.filter((item) => item !== name));
  flashBoard("“" + name + "” removed from the project list", "success");
}

// A click anywhere else closes an open panel.
document.addEventListener("mousedown", (event) => {
  for (const combo of projectCombos) {
    if (!combo.open) continue;
    const wrap = combo.input.closest(".combo");
    if (wrap && !wrap.contains(event.target)) combo.close();
  }
});

/* ------------------------------------------------------------- board render */

// Days newest first; inside a day, projects in the order they first appear,
// with un-projected tasks ahead of them - the same order the filed report is
// composed in, so the board reads like a preview of it.  The server already
// sends tasks in their day order, dragged positions included.
function groupDay(tasks) {
  const projects = [];
  const byProject = new Map();
  for (const task of tasks) {
    let bucket = byProject.get(task.project);
    if (!bucket) {
      bucket = { project: task.project, tasks: [] };
      byProject.set(task.project, bucket);
      projects.push(bucket);
    }
    bucket.tasks.push(task);
  }
  // Array#sort is stable, so everything else keeps its first-seen order.
  projects.sort((a, b) => (a.project === "" ? -1 : 0) - (b.project === "" ? -1 : 0));
  return projects;
}

function groupBoard(tasks) {
  const byDate = new Map();
  for (const task of tasks) {
    if (!byDate.has(task.date)) byDate.set(task.date, []);
    byDate.get(task.date).push(task);
  }
  return Array.from(byDate.keys())
    .sort((a, b) => (a < b ? 1 : a > b ? -1 : 0))
    .map((date) => ({ date: date, tasks: byDate.get(date), projects: groupDay(byDate.get(date)) }));
}

// One day as report text, exactly as file_checked_tasks writes it: a
// [project] header, "• " bullets with any extra lines indented under them,
// and a blank line between projects.
function composeDay(tasks) {
  const blocks = [];
  for (const bucket of groupDay(tasks)) {
    const lines = bucket.project ? ["[" + bucket.project + "]"] : [];
    for (const task of bucket.tasks) {
      const own = String(task.text).trim().split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
      if (!own.length) continue;
      lines.push("• " + own[0]);
      for (const line of own.slice(1)) lines.push("  " + line);
    }
    if (lines.length) blocks.push(lines.join("\n"));
  }
  return blocks.join("\n\n").trim();
}

function baseTasks() {
  return showFiledTasks ? boardTasks : boardTasks.filter((task) => !task.filed_at);
}

function renderChips() {
  const base = baseTasks();
  const box = $("projectChips");
  box.innerHTML = "";
  const counts = new Map();
  for (const task of base) counts.set(task.project, (counts.get(task.project) || 0) + 1);
  if (projectFilter !== "all" && !counts.has(projectFilter)) projectFilter = "all";
  if (counts.size < 2 && projectFilter === "all") { box.classList.add("hidden"); return; }
  box.classList.remove("hidden");
  box.appendChild(el("span", "chips-label", "Project"));

  // Most recently used first, the same order as the dropdown.
  const names = boardProjects.filter((name) => counts.has(name));
  for (const name of counts.keys()) if (!names.includes(name)) names.push(name);
  const chips = [{ key: "all", label: "All", count: base.length }].concat(
    names.map((name) => ({ key: name, label: name || "(no project)", count: counts.get(name) }))
  );
  for (const chip of chips) {
    const button = el("button", "chip" + (projectFilter === chip.key ? " is-on" : ""), chip.label);
    button.type = "button";
    button.appendChild(el("span", "chip-count", String(chip.count)));
    button.addEventListener("click", () => {
      projectFilter = chip.key;
      try { localStorage.setItem("taskReporterFilter", projectFilter); } catch (err) { /* private mode */ }
      renderBoard();
    });
    box.appendChild(button);
  }
}

function renderBoard() {
  const ready = boardCounts.ready || 0;
  tabBadge.textContent = ready ? String(ready) : "";

  const summary = $("boardSummary");
  summary.innerHTML = "";
  if (boardCounts.total) {
    summary.append(boardCounts.open + " open · ", el("b", null, ready + " ready to file"), " · " + boardCounts.filed + " filed");
  }
  $("fileBtn").disabled = !ready;
  $("clearFiledBtn").disabled = !boardCounts.filed;

  renderChips();
  const visible = baseTasks().filter((task) => projectFilter === "all" || task.project === projectFilter);

  // Keep the scroll position and the focused row across a re-render.
  const scrollTop = boardScroll.scrollTop;
  const focusedId = document.activeElement && document.activeElement.dataset
    ? document.activeElement.dataset.taskId : null;
  boardScroll.innerHTML = "";

  if (!visible.length) {
    let text;
    if (!boardTasks.length) text = "No tasks yet.\nAdd one above - the day and the [project] are what the filed report is grouped by.";
    else if (projectFilter !== "all") text = "Nothing on the board for [" + projectFilter + "].";
    else text = "Everything on the board is filed.\nTick “Show filed” to see it.";
    boardScroll.appendChild(el("div", "board-empty", text));
    return;
  }

  for (const day of groupBoard(visible)) {
    const dayNode = el("section", "day");
    const head = el("div", "day-head");
    head.appendChild(el("span", "day-date", isoToDisplay(day.date)));
    head.appendChild(el("span", "day-when", relativeDay(day.date)));
    const done = day.tasks.filter((task) => task.done).length;
    head.appendChild(el("span", "day-count" + (done === day.tasks.length ? " all-done" : ""),
      done + " of " + day.tasks.length + " done"));
    const copy = el("button", "btn btn-ghost day-copy");
    copy.type = "button";
    copy.title = "Copy this day as report text";
    copy.append(icon("copy", 14, 2), "Copy day");
    copy.addEventListener("click", async () => {
      const ok = await copyText(composeDay(day.tasks));
      showToast(ok ? "Copied " + isoToDisplay(day.date) + " to the clipboard" : "Could not reach the clipboard");
    });
    head.appendChild(copy);
    dayNode.appendChild(head);

    for (const bucket of day.projects) {
      const group = el("div", "proj");
      const label = bucket.project
        ? el("div", "proj-head", "[" + bucket.project + "]")
        : el("div", "proj-head no-project", "(no project)");
      label.dataset.date = day.date;
      label.dataset.project = bucket.project;
      wireProjectDrop(label);
      group.appendChild(label);
      for (const task of bucket.tasks) group.appendChild(taskRow(task));
      dayNode.appendChild(group);
    }
    boardScroll.appendChild(dayNode);
  }

  boardScroll.scrollTop = scrollTop;
  if (focusedId) {
    const again = boardScroll.querySelector('.task[data-task-id="' + CSS.escape(focusedId) + '"]');
    if (again) again.focus();
  }
}

function taskRow(task) {
  const filed = Boolean(task.filed_at);
  const row = el("div", "task" + (task.done ? " is-done" : "") + (filed ? " is-filed" : ""));
  row.tabIndex = 0;
  row.dataset.taskId = task.id;
  row.dataset.date = task.date;

  const grip = el("span", "grip-handle");
  grip.title = "Drag to reorder (Alt + ↑ / ↓)";
  grip.appendChild(icon("grip", 14, 1));
  grip.addEventListener("mousedown", () => { row.draggable = true; });
  row.appendChild(grip);

  const box = el("button", "box");
  box.type = "button";
  box.appendChild(icon("check", 12, 3.5));
  box.disabled = filed;
  box.setAttribute("role", "checkbox");
  box.setAttribute("aria-checked", task.done ? "true" : "false");
  box.title = filed
    ? "Already written into the workbook on " + task.filed_at
    : task.done ? "Done - will be filed on the next “File Checked Tasks”" : "Tick when this is done";
  box.addEventListener("click", () => toggleTask(task, !task.done));
  row.appendChild(box);

  // A button: it is genuinely the way into the edit dialog, and textContent
  // keeps arbitrary task text out of the parser.
  const text = el("button", "task-text", task.text);
  text.type = "button";
  text.title = "Click to edit";
  text.addEventListener("click", () => openTaskEdit(task.id));
  row.appendChild(text);

  const files = task.files || [];
  if (files.length) {
    const clip = el("button", "clip");
    clip.type = "button";
    clip.title = files.map((file) => file.name).join("\n");
    clip.append(icon("paperclip", 12, 2), String(files.length));
    clip.addEventListener("click", () => openTaskEdit(task.id));
    row.appendChild(clip);
  }

  if (filed) row.appendChild(el("span", "tag tag-neutral tag-small", "filed"));

  const meta = el("span", "task-meta", clockOf(task.done_at || task.created_at));
  meta.title = "Added " + (task.created_at || "?") +
    (task.done_at ? "\nTicked " + task.done_at : "") + (filed ? "\nFiled " + task.filed_at : "");
  row.appendChild(meta);

  row.appendChild(miniButton("pencil", "Edit task & files", false, () => openTaskEdit(task.id)));
  row.appendChild(miniButton("trash", "Delete", true, () => deleteTask(task)));

  row.addEventListener("keydown", (event) => {
    if (event.target !== row) return;
    if (event.altKey && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
      event.preventDefault();
      nudgeTask(task, event.key === "ArrowUp" ? -1 : 1);
    } else if (event.key === "Enter") {
      event.preventDefault();
      openTaskEdit(task.id);
    } else if (event.key === " ") {
      event.preventDefault();
      if (!filed) toggleTask(task, !task.done);
    }
  });
  wireTaskDrag(row, task);
  return row;
}

/* ------------------------------------------------------------- drag & drop */

// Within a day only: a task belongs to its day, and moving it to another day
// is a date change, which the edit dialog does.  Dropping on another project's
// header in the same day moves the task into that project.
let dragging = null;

function dayIds(date) {
  return boardTasks.filter((task) => task.date === date).map((task) => task.id);
}

function clearDropMarks() {
  boardScroll.querySelectorAll(".drop-before, .drop-after, .drop-into").forEach((node) => {
    node.classList.remove("drop-before", "drop-after", "drop-into");
  });
}

function wireTaskDrag(row, task) {
  row.addEventListener("dragstart", (event) => {
    if (!row.draggable) { event.preventDefault(); return; }
    dragging = { id: task.id, date: task.date, project: task.project };
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", task.text);
    row.classList.add("dragging");
  });
  row.addEventListener("dragend", () => {
    row.draggable = false;
    row.classList.remove("dragging");
    dragging = null;
    clearDropMarks();
  });
  row.addEventListener("mouseup", () => { row.draggable = false; });
  row.addEventListener("dragover", (event) => {
    if (!dragging || dragging.date !== task.date || dragging.id === task.id) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
    const box = row.getBoundingClientRect();
    const after = event.clientY > box.top + box.height / 2;
    clearDropMarks();
    row.classList.add(after ? "drop-after" : "drop-before");
  });
  row.addEventListener("drop", (event) => {
    if (!dragging || dragging.date !== task.date || dragging.id === task.id) return;
    event.preventDefault();
    const after = row.classList.contains("drop-after");
    const moved = dragging;
    clearDropMarks();
    const ids = dayIds(task.date).filter((id) => id !== moved.id);
    let at = ids.indexOf(task.id);
    if (after) at += 1;
    ids.splice(at, 0, moved.id);
    commitOrder(task.date, ids, moved.id, task.project);
  });
}

function wireProjectDrop(label) {
  label.addEventListener("dragover", (event) => {
    if (!dragging || dragging.date !== label.dataset.date) return;
    event.preventDefault();
    clearDropMarks();
    label.classList.add("drop-into");
  });
  label.addEventListener("dragleave", () => label.classList.remove("drop-into"));
  label.addEventListener("drop", (event) => {
    if (!dragging || dragging.date !== label.dataset.date) return;
    event.preventDefault();
    const moved = dragging;
    const project = label.dataset.project;
    clearDropMarks();
    // To the end of that project's tasks.
    const day = boardTasks.filter((task) => task.date === moved.date && task.id !== moved.id);
    const ids = day.map((task) => task.id);
    let last = -1;
    day.forEach((task, index) => { if (task.project === project) last = index; });
    ids.splice(last === -1 ? ids.length : last + 1, 0, moved.id);
    commitOrder(moved.date, ids, moved.id, project);
  });
}

// Alt+Up / Alt+Down: swap with the neighbour in the same project, which is the
// neighbour on screen.
function nudgeTask(task, step) {
  const day = boardTasks.filter((item) => item.date === task.date);
  const ids = day.map((item) => item.id);
  const peers = day.filter((item) => item.project === task.project &&
    (showFiledTasks || !item.filed_at || item.id === task.id));
  const at = peers.findIndex((item) => item.id === task.id);
  const other = peers[at + step];
  if (!other) return;
  const a = ids.indexOf(task.id);
  const b = ids.indexOf(other.id);
  ids[a] = other.id;
  ids[b] = task.id;
  commitOrder(task.date, ids, task.id, task.project);
}

async function commitOrder(date, ids, movedId, project) {
  const moved = boardTasks.find((task) => task.id === movedId);
  const changesProject = moved && moved.project !== project;
  // Shown at once; the server catches up behind it.
  const position = new Map(ids.map((id, index) => [id, index]));
  if (moved) moved.project = project;
  boardTasks.sort((a, b) =>
    a.date !== b.date ? (a.date < b.date ? -1 : 1)
      : a.date === date ? position.get(a.id) - position.get(b.id) : 0);
  renderBoard();
  const row = boardScroll.querySelector('.task[data-task-id="' + CSS.escape(movedId) + '"]');
  if (row) row.focus();
  try {
    if (changesProject) {
      const out = await api("/api/tasks/update", { id: movedId, project: project });
      if (!out.ok) { flashBoard(out.message || "The task was not moved", "danger"); await loadBoard(); return; }
    }
    const out = await api("/api/tasks/reorder", { date: date, ids: ids });
    if (!out.ok) flashBoard(out.message || "The new order was not saved", "danger");
    else if (changesProject) flashBoard("Moved to " + (project ? "[" + project + "]" : "(no project)"), "success");
  } catch (err) {
    flashBoard("Lost contact with the reporter - the order was not saved", "danger");
  }
  await loadBoard();
}

/* --------------------------------------------------------------- board edits */

async function boardCall(path, body, failure) {
  try {
    const out = await api(path, body);
    if (!out.ok) {
      flashBoard(out.message || failure, "danger");
      await loadBoard();
      return null;
    }
    await loadBoard();
    return out;
  } catch (err) {
    flashBoard("Lost contact with the reporter - " + failure.toLowerCase(), "danger");
    await loadBoard();
    return null;
  }
}

// The add-a-task box grows with what is typed in it, one to five lines.
function fitTaskBox(box) {
  box.style.height = "auto";
  const style = getComputedStyle(box);
  const line = parseFloat(style.lineHeight) || 20;
  const chrome = parseFloat(style.paddingTop) + parseFloat(style.paddingBottom) +
    parseFloat(style.borderTopWidth) + parseFloat(style.borderBottomWidth);
  const lines = Math.min(5, Math.max(1, Math.round((box.scrollHeight - chrome + 2) / line)));
  box.style.height = Math.max(40, lines * line + chrome) + "px";
}

async function addTask(event) {
  if (event) event.preventDefault();
  const textBox = $("taskText");
  const text = textBox.value.trim();
  if (!text) {
    flashBoard("Type the task first", "danger");
    textBox.focus();
    return;
  }
  const out = await boardCall(
    "/api/tasks/add",
    { date: $("taskDate").value, project: $("taskProject").value.trim(), text: text },
    "The task was not added"
  );
  if (out) {
    // Only the text clears.  Tasks arrive in runs under one day and one
    // project, so those two boxes are left exactly where they were.
    textBox.value = "";
    fitTaskBox(textBox);
    clearDraft("task");
    flashBoard("Task added", "success");
  }
  textBox.focus();
}

async function toggleTask(task, done) {
  const out = await boardCall("/api/tasks/update", { id: task.id, done: done }, "The tick was not saved");
  if (out) flashBoard(done ? "Ticked - ready to file" : "Unticked", done ? "success" : null);
}

async function deleteTask(task) {
  const out = await boardCall("/api/tasks/delete", { id: task.id }, "Nothing was deleted");
  if (!out) return;
  const gone = out.task || task;
  showToast("Task deleted — “" + shorten(task.text, 48) + "”", async () => {
    const back = await boardCall("/api/tasks/restore", { task: gone }, "The task was not restored");
    if (back) flashBoard("Task restored", "success");
  });
}

async function clearFiled() {
  const filed = boardCounts.filed || 0;
  if (!filed) { flashBoard("Nothing filed to clear", "warn"); return; }
  const out = await boardCall("/api/tasks/clear-filed", {}, "Nothing was cleared");
  if (!out) return;
  showToast(out.removed + " filed task" + (out.removed === 1 ? "" : "s") + " cleared from the board", async () => {
    const back = await boardCall("/api/tasks/clear-filed/undo", {}, "Nothing was restored");
    if (back) flashBoard("Restored " + back.restored + " filed task" + (back.restored === 1 ? "" : "s"), "success");
  });
}

/* -------------------------------------------------------- task edit + files */

// Files added or removed in the dialog are only staged: like the text, they
// are committed by Save and thrown away by Discard.
let editingTaskId = null;
let taskOriginal = null;
let stagedFiles = [];      // { file, name, size, url }
let removedFiles = new Set();

const IMAGE_NAME = /\.(png|jpe?g|gif|webp|bmp|svg|ico)$/i;

function sizeLabel(bytes) {
  const kb = (bytes || 0) / 1024;
  if (kb >= 1024) return (kb / 1024).toFixed(1) + " MB";
  return Math.max(1, Math.round(kb)) + " KB";
}

function editingTask() {
  return boardTasks.find((item) => item.id === editingTaskId) || null;
}

function currentFileNames() {
  const task = editingTask();
  const kept = (task ? task.files || [] : []).map((file) => file.name).filter((name) => !removedFiles.has(name));
  return kept.concat(stagedFiles.map((item) => "+" + item.name));
}

function taskSnapshot() {
  return JSON.stringify({
    date: $("tDate").value, project: $("tProject").value, text: $("tText").value, files: currentFileNames(),
  });
}

function taskChanged() {
  const dirty = guarded.taskEditBackdrop.isDirty();
  $("tDirty").classList.toggle("hidden", !dirty);
  setDraft("edit", dirty && editingTaskId
    ? { kind: "task", id: editingTaskId, date: $("tDate").value, project: $("tProject").value, text: $("tText").value }
    : null);
  reportDirty();
}

function openTaskEdit(id, draft) {
  const task = boardTasks.find((item) => item.id === id);
  if (!task) return false;
  if (editingTaskId && editingTaskId !== id) dropStaged();
  editingTaskId = id;
  $("tDate").value = task.date;
  $("tProject").value = task.project;
  $("tText").value = task.text;
  $("tError").textContent = "";
  removedFiles = new Set();
  taskOriginal = taskSnapshot();
  if (draft) {
    $("tDate").value = draft.date || task.date;
    $("tProject").value = draft.project !== undefined ? draft.project : task.project;
    $("tText").value = draft.text !== undefined ? draft.text : task.text;
  }
  renderTaskFiles();
  openModal("taskEditBackdrop");
  taskChanged();
  $("tText").focus();
  return true;
}

function dropStaged() {
  for (const item of stagedFiles) if (item.url) URL.revokeObjectURL(item.url);
  stagedFiles = [];
  removedFiles = new Set();
}

function stageFiles(list) {
  let rejected = 0;
  for (const file of Array.from(list || [])) {
    if (file.size > CFG.maxFileBytes) { rejected += 1; continue; }
    let name = file.name || "";
    if (!name || name === "image.png") name = pastedName(file.type);
    stagedFiles.push({
      file: file, name: name, size: file.size,
      url: /^image\//.test(file.type) ? URL.createObjectURL(file) : null,
    });
  }
  if (rejected) {
    $("tError").textContent = rejected + (rejected === 1 ? " file is" : " files are") +
      " over " + Math.round(CFG.maxFileBytes / 1048576) + " MB and " + (rejected === 1 ? "was" : "were") + " not added.";
  }
  renderTaskFiles();
  taskChanged();
}

function pastedName(type) {
  const now = new Date();
  const ext = (String(type || "image/png").split("/")[1] || "png").replace("jpeg", "jpg").replace(/[^a-z0-9]/gi, "");
  return "pasted-" + now.getFullYear() + pad2(now.getMonth() + 1) + pad2(now.getDate()) + "-" +
    pad2(now.getHours()) + pad2(now.getMinutes()) + pad2(now.getSeconds()) + "." + (ext || "png");
}

function windowsJoin(folder, ...parts) {
  const sep = String(folder).includes("\\") ? "\\" : "/";
  return [String(folder).replace(/[\\/]+$/, "")].concat(parts).join(sep);
}

function fileRow(entry) {
  const isImage = IMAGE_NAME.test(entry.name);
  const row = el("div", "file-row" + (entry.staged ? " is-new" : ""));
  const thumb = el("div", "thumb");
  if (isImage && entry.thumb) {
    const img = document.createElement("img");
    img.alt = "";
    img.loading = "lazy";
    img.src = entry.thumb;
    img.addEventListener("error", () => { img.replaceWith(icon("image", 18)); });
    thumb.appendChild(img);
  } else if (isImage) {
    thumb.appendChild(icon("image", 18));
  } else {
    const ext = (entry.name.includes(".") ? entry.name.split(".").pop() : "file").slice(0, 4);
    thumb.appendChild(el("span", "thumb-ext", ext));
  }
  row.appendChild(thumb);

  const info = el("div", "file-info");
  info.appendChild(el("span", "file-name", entry.name));
  info.appendChild(el("span", "file-meta", entry.meta));
  row.appendChild(info);

  const action = (label, title, handler, disabled) => {
    const button = el("button", "btn btn-ghost", label);
    button.type = "button";
    if (title) button.title = title;
    button.disabled = Boolean(disabled);
    button.addEventListener("click", handler);
    row.appendChild(button);
  };
  action(isImage ? "Copy image" : "Copy path",
    isImage ? "Copy the picture to the clipboard - paste it straight into an AI chat" : "Copy the file's full path",
    () => copyTaskFile(entry, isImage), !isImage && entry.staged);
  action("Open", entry.staged ? "Save the task first" : "", () => taskFileAction("open", entry), entry.staged);
  action("Show in folder", entry.staged ? "Save the task first" : "", () => taskFileAction("reveal", entry), entry.staged);
  row.appendChild(miniButton("x", "Remove from task", true, entry.remove));
  return row;
}

function renderTaskFiles() {
  const list = $("tFilesList");
  list.innerHTML = "";
  const task = editingTask();
  const entries = [];
  for (const file of task ? task.files || [] : []) {
    if (removedFiles.has(file.name)) continue;
    entries.push({
      name: file.name, staged: false, thumb: fileUrl(task.id, file.name),
      meta: sizeLabel(file.size) + (file.added_at ? " · added " + file.added_at.slice(0, 16) : ""),
      remove: () => { removedFiles.add(file.name); renderTaskFiles(); taskChanged(); },
    });
  }
  stagedFiles.forEach((item, index) => {
    entries.push({
      name: item.name, staged: true, thumb: item.url, blob: item.file,
      meta: sizeLabel(item.size) + " · not saved yet",
      remove: () => {
        if (item.url) URL.revokeObjectURL(item.url);
        stagedFiles.splice(index, 1);
        renderTaskFiles();
        taskChanged();
      },
    });
  });
  for (const entry of entries) list.appendChild(fileRow(entry));
  const count = entries.length;
  $("tFilesCount").textContent = count ? count + (count === 1 ? " file" : " files") : "None yet";
}

async function copyTaskFile(entry, isImage) {
  if (isImage) {
    try {
      const blob = entry.blob || await (await fetch(entry.thumb, { cache: "no-store" })).blob();
      const ok = await copyImageBlob(blob);
      showToast(ok ? "Image copied — paste it with Ctrl + V" : "Could not put the image on the clipboard");
    } catch (err) {
      showToast("Could not read the image");
    }
    return;
  }
  const ok = await copyText(windowsJoin(CFG.taskFilesDir, editingTaskId, entry.name));
  showToast(ok ? "Path copied" : "Could not reach the clipboard");
}

async function taskFileAction(kind, entry) {
  try {
    const out = await api("/api/tasks/files/" + kind, { id: editingTaskId, name: entry.name });
    showToast(out.message || (out.ok ? "Done" : "That did not work"));
  } catch (err) {
    showToast("Lost contact with the reporter.");
  }
}

async function saveTaskEdit() {
  if (editingTaskId === null) return false;
  const id = editingTaskId;
  const error = $("tError");
  const text = $("tText").value.trim();
  if (!text) { error.textContent = "The task cannot be empty."; return false; }
  if (text.length > CFG.maxTaskLength) {
    error.textContent = "Task is too long (" + text.length + " chars). Limit is " + CFG.maxTaskLength + ".";
    return false;
  }
  error.textContent = "Saving…";
  const button = $("tSave");
  button.disabled = true;

  const changes = { id: id, project: $("tProject").value.trim(), text: text };
  // A blank date means "today" when adding a task, which is not what clearing
  // this box should do - leave the day where it is instead.
  const when = $("tDate").value;
  if (when) changes.date = when;

  const problems = [];
  try {
    const out = await api("/api/tasks/update", changes);
    if (!out.ok) { error.textContent = out.message || "Could not save the task."; button.disabled = false; return false; }
    // Then the files: new ones up, removed ones to the trash.
    while (stagedFiles.length) {
      const item = stagedFiles[0];
      const up = await uploadTaskFile(id, item.file, item.name);
      if (!up.ok) { problems.push(item.name + ": " + (up.message || "not stored")); }
      if (item.url) URL.revokeObjectURL(item.url);
      stagedFiles.shift();
    }
    for (const name of Array.from(removedFiles)) {
      const gone = await api("/api/tasks/files/remove", { id: id, name: name });
      if (!gone.ok) problems.push(name + ": " + (gone.message || "not removed"));
      removedFiles.delete(name);
    }
  } catch (err) {
    error.textContent = "Lost contact with the reporter. Some changes may not have been saved.";
    button.disabled = false;
    await loadBoard();
    renderTaskFiles();
    return false;
  }
  button.disabled = false;
  await loadBoard();
  if (problems.length) {
    // The text is saved; say which files were not, and stay open.
    taskOriginal = taskSnapshot();
    renderTaskFiles();
    error.textContent = problems.join(" · ");
    taskChanged();
    return false;
  }
  taskOriginal = taskSnapshot();
  closeModal("taskEditBackdrop");
  flashBoard("Task saved", "success");
  return true;
}

guarded.taskEditBackdrop = {
  confirmText: "You changed this task. If you close now, those changes are lost.",
  isDirty: () => editingTaskId !== null && taskOriginal !== null && taskSnapshot() !== taskOriginal,
  save: saveTaskEdit,
  discard: () => { dropStaged(); taskOriginal = taskSnapshot(); },
  onClosed: () => {
    dropStaged();
    editingTaskId = null;
    taskOriginal = null;
    $("tDirty").classList.add("hidden");
    clearDraft("edit");
  },
};

["tDate", "tProject", "tText"].forEach((id) => $(id).addEventListener("input", taskChanged));
$("tDate").addEventListener("change", taskChanged);
$("tSave").addEventListener("click", saveTaskEdit);
$("tFilesAdd").addEventListener("click", () => $("tFilesInput").click());
$("tDrop").addEventListener("click", () => $("tFilesInput").click());
$("tDrop").addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") { event.preventDefault(); $("tFilesInput").click(); }
});
$("tFilesInput").addEventListener("change", (event) => { stageFiles(event.target.files); event.target.value = ""; });

function wireDropZone(zone, onFiles) {
  zone.addEventListener("dragover", (event) => {
    if (!event.dataTransfer || !Array.from(event.dataTransfer.types || []).includes("Files")) return;
    event.preventDefault();
    zone.classList.add("over");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("over"));
  zone.addEventListener("drop", (event) => {
    if (!event.dataTransfer || !event.dataTransfer.files.length) return;
    event.preventDefault();
    zone.classList.remove("over");
    onFiles(event.dataTransfer.files);
  });
}
wireDropZone($("tDrop"), stageFiles);
// The whole dialog takes a drop, not just the dashed box.
wireDropZone(document.querySelector("#taskEditBackdrop .modal"), stageFiles);

// A file dropped anywhere else must not make the window navigate to it.
document.addEventListener("dragover", (event) => {
  if (event.dataTransfer && Array.from(event.dataTransfer.types || []).includes("Files")) event.preventDefault();
});
document.addEventListener("drop", (event) => {
  if (event.dataTransfer && event.dataTransfer.files.length) event.preventDefault();
});

// Ctrl+V of an image (a screenshot, mostly) while the dialog is open.
document.addEventListener("paste", (event) => {
  const target = IS_QUICK ? "quick" : topLayer() === "taskEditBackdrop" ? "task" : null;
  if (!target || !event.clipboardData) return;
  const files = Array.from(event.clipboardData.files || []);
  if (!files.length) return;
  event.preventDefault();
  if (target === "task") stageFiles(files);
  else quickStage(files);
});

/* --------------------------------------------------------------- the filing */

let workbookLocked = false;

async function startFiling() {
  try {
    const out = await api("/api/tasks/preview");
    previewGroups = out.groups || [];
    workbookLocked = Boolean(out.workbookOpen);
  } catch (err) {
    flashBoard("Lost contact with the reporter", "danger");
    return;
  }
  if (!previewGroups.length) {
    flashBoard("Nothing to file - tick the tasks you finished first", "warn");
    await loadBoard();
    return;
  }
  renderPreview();
  openModal("previewBackdrop");
  $("previewConfirm").focus();
}

function renderPreview() {
  const rows = previewGroups.length;
  const tasks = previewGroups.reduce((sum, group) => sum + group.taskCount, 0);

  // A day already in the sheet is added to its cell rather than given a row of
  // its own, so counting every group as a new row would overstate it.
  const grown = previewGroups.filter((group) => group.appendsToExisting).length;
  const fresh = rows - grown;
  const parts = [];
  if (fresh) parts.push(fresh + (fresh === 1 ? " new row" : " new rows"));
  if (grown) parts.push(grown + (grown === 1 ? " day" : " days") + " added to");
  $("previewCount").textContent = tasks + (tasks === 1 ? " task" : " tasks") + " → " + parts.join(" + ");

  // The chips only filter what the board shows; filing takes every tick.
  const filter = $("previewFilter");
  filter.classList.toggle("hidden", projectFilter === "all");
  filter.textContent = projectFilter === "all" ? "" :
    "The board is filtered to " + (projectFilter ? "[" + projectFilter + "]" : "(no project)") +
    ", but filing takes every ticked task on the board - all of them are below.";

  // Excel holding the workbook is worth saying before the button is pressed.
  const locked = $("previewLocked");
  locked.classList.toggle("hidden", !workbookLocked);
  locked.textContent = workbookLocked
    ? "task_reports.xlsx is open in Excel, so these rows will be queued and written in once Excel closes it. The tasks are still marked filed."
    : "";
  $("previewConfirm").textContent = workbookLocked ? "Queue for the Workbook" : "Write to Workbook";

  const wrap = $("previewWrap");
  wrap.innerHTML = "";
  for (const group of previewGroups) {
    const node = el("div", "pv-group");
    const head = el("div", "pv-head");
    head.appendChild(el("span", "pv-when", group.timestamp));
    const over = group.totalLength > CFG.maxCell || group.length > CFG.maxCell;
    let note = group.taskCount + " task" + (group.taskCount === 1 ? "" : "s") + " · " +
      (group.appendsToExisting ? "adds to the existing row · " : "new row · ") +
      Number(group.totalLength).toLocaleString() + " / " + Number(CFG.maxCell).toLocaleString() + " chars";
    if (over) note += " · too long for one cell";
    head.appendChild(el("span", "pv-note" + (over ? " pv-warn" : ""), note));
    node.appendChild(head);
    if (group.appendsToExisting) {
      node.appendChild(el("div", "pv-merge",
        "This day is already in the workbook. The text below goes a blank line under what is there, and the row's Date-Time moves to " +
        group.timestamp + "."));
    }
    node.appendChild(el("pre", "pv-body", group.text));
    wrap.appendChild(node);
  }
}

async function confirmFiling() {
  const button = $("previewConfirm");
  if (button.disabled) return;
  button.disabled = true;
  let out;
  try {
    // Recomputed server-side rather than trusting the preview, so whatever is
    // ticked at this moment is what gets written.
    out = await api("/api/tasks/file", {});
  } catch (err) {
    // The dialog stays up with the preview intact, so the write can be retried.
    button.disabled = false;
    showToast("Lost contact with the reporter. Nothing was filed.");
    return;
  }
  button.disabled = false;
  closeModal("previewBackdrop");
  await loadBoard();

  const written = out.written || [];
  const failures = out.failed || [];
  if (!written.length) {
    showToast(failures.length ? failures.map((f) => f.message).join(" · ") : out.message || "Nothing was filed");
    return;
  }
  const filedAt = Date.now();
  const message = "Filed " + out.filedCount + (out.filedCount === 1 ? " task" : " tasks") + " into " +
    written.length + (written.length === 1 ? " row" : " rows") +
    (out.queued ? " — queued until Excel closes the workbook" : "") +
    (failures.length ? " · " + failures.length + " day(s) could not be filed" : "");
  // Undo only for a write that went straight into the sheet, and only while
  // the toast is up.
  showToast(message, out.undoable ? async () => {
    if (Date.now() - filedAt > 6500) { showToast("Too late to undo that filing."); return; }
    const back = await boardCall("/api/tasks/file/undo", {}, "The filing was not undone");
    if (back) showToast("Filing undone — " + back.count + " task" + (back.count === 1 ? " is" : "s are") + " ready to file again");
  } : null);
  if (out.queued) refreshBanner(true, Math.max(1, lastPing.pendingCount));
  flashBoard(written.map((row) => row.dateLabel).join(", ") + " filed ✓", out.queued ? "warn" : "success");
}

/* ------------------------------------------------------------ view switching */

const VIEW_NAMES = { report: "Report", board: "Board", settings: "Settings" };
let viewBeforeSettings = "report";

function focusView() {
  if (IS_QUICK) { $("qText").focus(); return; }
  if (currentView === "board") $("taskText").focus();
  else if (currentView === "report") editor.focus();
}

function setView(name) {
  if (!VIEW_NAMES[name]) name = "report";
  if (name === "settings" && currentView !== "settings") viewBeforeSettings = currentView;
  currentView = name;
  $("reportView").classList.toggle("hidden", name !== "report");
  $("boardView").classList.toggle("hidden", name !== "board");
  $("settingsView").classList.toggle("hidden", name !== "settings");
  $("tabReport").classList.toggle("is-on", name === "report");
  $("tabBoard").classList.toggle("is-on", name === "board");
  $("settingsBtn").classList.toggle("is-on", name === "settings");
  $("winSub").textContent = "— " + VIEW_NAMES[name];
  document.title = "Task Reporter — " + VIEW_NAMES[name];
  if (name !== "settings") {
    try { localStorage.setItem("taskReporterView", name); } catch (err) { /* private mode */ }
  }
  if (name === "board") { loadBoard(); setTimeout(() => fitTaskBox($("taskText")), 0); }
  if (name === "settings") loadSettings();
  focusView();
}

/* ------------------------------------------------------------------- settings */

let settingsState = null;

async function loadSettings() {
  try {
    settingsState = await api("/api/settings");
  } catch (err) {
    $("sDirSource").textContent = "Could not read the settings.";
    return;
  }
  renderSettings();
}

function renderSettings() {
  const state = settingsState;
  if (!state) return;
  $("sDataDir").value = state.dataDir;
  let source = "Found from " + state.dataDirSource + ".";
  if (state.envOverride) source += " %TASK_REPORT_DIR% is set, so it wins over anything chosen here.";
  else source += " Setting %TASK_REPORT_DIR% overrides this.";
  if (state.restartNeeded) source += " The folder you chose is used from the next start.";
  $("sDirSource").textContent = source;
  $("sChangeDir").classList.toggle("hidden", !IS_APP);

  const files = state.files;
  const describe = (card, extra) => (card.exists ? extra + (card.modified ? " · changed " + card.modified.slice(0, 16) : "") : "Not created yet");
  $("sWorkbookNote").textContent = describe(files.workbook, files.workbook.rows + (files.workbook.rows === 1 ? " report row" : " report rows"));
  $("sBoardNote").textContent = describe(files.board, files.board.tasks + " tasks · " + files.board.projects + " projects");
  $("sLogNote").textContent = files.log.exists ? "Start-up address and save errors · " + sizeLabel(files.log.size) : "Only written by the desktop app";

  $("sKeepInTray").checked = Boolean(state.settings.keepInTray);
  $("sKeepInTray").disabled = !state.tray.available;
  $("sHotkey").value = state.settings.hotkey || "";
  $("helpHotkey").textContent = (state.settings.hotkey || "Ctrl+Alt+T").replace(/\+/g, " + ");
  let trayNote = "Left-click the tray icon to quick-add a task. Right-click: Open, Quick add task, Previous reports, Quit.";
  if (!state.tray.available) trayNote = "The tray icon is part of the Windows desktop app - it is not running in this window.";
  else if (state.tray.hotkeyError) trayNote += " " + state.tray.hotkeyError;
  $("sTrayNote").textContent = trayNote;
  refreshBanner(state.workbookOpen, state.pendingCount);
  syncThemeRadios();
}

async function saveSetting(changes, okMessage) {
  try {
    const out = await api("/api/settings", changes);
    if (!out.ok) { showToast(out.message || "The setting was not saved"); return false; }
    settingsState = out;
    renderSettings();
    if (okMessage) showToast(okMessage);
    return true;
  } catch (err) {
    showToast("Lost contact with the reporter - the setting was not saved");
    return false;
  }
}

$("settingsBtn").addEventListener("click", () => setView(currentView === "settings" ? viewBeforeSettings : "settings"));
$("settingsBack").addEventListener("click", () => setView(viewBeforeSettings));
$("sOpenDir").addEventListener("click", () => openKnown("folder", false));
document.querySelectorAll("[data-reveal]").forEach((button) => {
  button.addEventListener("click", () => openKnown(button.dataset.reveal, true));
});
document.querySelectorAll("[data-open]").forEach((button) => {
  button.addEventListener("click", () => openKnown(button.dataset.open, false));
});
$("sChangeDir").addEventListener("click", async () => {
  const folder = await callNative("pick_folder");
  if (!folder) return;
  await saveSetting({ dataDir: folder }, "Saved. Restart Task Reporter to use " + folder);
});
$("sKeepInTray").addEventListener("change", (event) => {
  saveSetting({ keepInTray: event.target.checked },
    event.target.checked ? "Closing the window now keeps Task Reporter in the tray" : "Closing the window now quits");
});
$("sHotkeySave").addEventListener("click", () => {
  saveSetting({ hotkey: $("sHotkey").value.trim() }, "Shortcut saved");
});
$("sHotkey").addEventListener("keydown", (event) => {
  if (event.key === "Enter") { event.preventDefault(); $("sHotkeySave").click(); }
});

async function openKnown(target, reveal) {
  try {
    const out = await api(reveal ? "/api/reveal" : "/api/open", { target: target });
    flashHere(out.message || (out.ok ? "Opened" : "That did not open"), out.ok ? "success" : "danger");
  } catch (err) {
    flashHere("Lost contact with the reporter.", "danger");
  }
}

/* ---------------------------------------------------------------------- theme */

const darkQuery = window.matchMedia ? matchMedia("(prefers-color-scheme: dark)") : null;

function themePref() {
  return document.documentElement.getAttribute("data-theme-pref") || "system";
}

function applyTheme(pref, persist) {
  const root = document.documentElement;
  const dark = pref === "dark" || (pref === "system" && darkQuery && darkQuery.matches);
  root.setAttribute("data-theme-pref", pref);
  root.setAttribute("data-theme", dark ? "dark" : "light");
  syncThemeRadios();
  if (!persist) return;
  try { localStorage.setItem("taskReporterTheme", pref); } catch (err) { /* private mode */ }
  // Python reads it too, so the next window opens in the right colour.
  api("/api/settings", { theme: pref }).catch(() => {});
  callNative("set_theme", dark ? "dark" : "light");
}

function syncThemeRadios() {
  const pref = themePref();
  document.querySelectorAll('input[name="themePick"]').forEach((radio) => { radio.checked = radio.value === pref; });
}

if (darkQuery) {
  darkQuery.addEventListener("change", () => { if (themePref() === "system") applyTheme("system", false); });
}
document.querySelectorAll('input[name="themePick"]').forEach((radio) => {
  radio.addEventListener("change", () => applyTheme(radio.value, true));
});
$("themeBtn").addEventListener("click", () => {
  const dark = document.documentElement.getAttribute("data-theme") === "dark";
  applyTheme(dark ? "light" : "dark", true);
});
// Another window (the quick-add one) changed it.
window.addEventListener("storage", (event) => {
  if (event.key === "taskReporterTheme" && event.newValue) applyTheme(event.newValue, false);
});

/* -------------------------------------------------------------- window chrome */

// The desktop window is frameless; this is its title bar.  Dragging and
// resizing are handed to Windows itself where the app can (so Aero Snap and
// Win+arrow keep working), and fall back to pywebview's drag region otherwise.
let nativeFrame = false;
let lastDragDown = 0;

async function setupWindowChrome() {
  if (!IS_APP) return;
  nativeFrame = Boolean(await callNative("native_frame"));
  document.documentElement.classList.toggle("native-frame", nativeFrame);
  syncMaximized(await callNative("is_maximized"));
}

function syncMaximized(state) {
  document.documentElement.classList.toggle("is-maximized", Boolean(state));
  $("winMax").title = state ? "Restore" : "Maximize";
  $("winMax").setAttribute("aria-label", state ? "Restore" : "Maximize");
}

async function toggleMaximize() {
  syncMaximized(await callNative("toggle_maximize"));
}

function wireDragRegion(region) {
  region.addEventListener("mousedown", (event) => {
    if (event.button !== 0 || !nativeFrame) return;
    // Stop pywebview's own drag-region handler: Windows does the move.
    event.stopPropagation();
    event.preventDefault();
    const now = Date.now();
    if (now - lastDragDown < 400) { lastDragDown = 0; if (!IS_QUICK) toggleMaximize(); return; }
    lastDragDown = now;
    callNative("start_drag");
  });
  region.addEventListener("dblclick", () => { if (!nativeFrame && !IS_QUICK) toggleMaximize(); });
}

wireDragRegion($("dragRegion"));
wireDragRegion($("quickDrag"));
document.querySelectorAll(".grip").forEach((grip) => {
  grip.addEventListener("mousedown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    callNative("start_resize", grip.dataset.edge);
  });
});
$("winMin").addEventListener("click", () => callNative("minimize"));
$("winMax").addEventListener("click", toggleMaximize);
$("winClose").addEventListener("click", () => requestWindowClose());
window.addEventListener("resize", () => {
  if (!IS_APP) return;
  clearTimeout(window.__maxTimer);
  window.__maxTimer = setTimeout(async () => syncMaximized(await callNative("is_maximized")), 120);
});

// Closing the window with an edit open asks first, exactly like the dialog
// would.  Python calls this too, for Alt+F4 and the taskbar's Close.
let closing = false;
async function requestWindowClose() {
  if (closing) return;
  closing = true;
  try {
    const id = Object.keys(guarded).find((key) => openStack.includes(key) && guarded[key].isDirty());
    if (id) {
      const answer = await askSaveChanges(guarded[id].confirmText);
      if (answer === "keep") return;
      if (answer === "save") { if (!(await guarded[id].save())) return; }
      if (answer === "discard") { guarded[id].discard(); closeModal(id); }
    }
    writeDrafts();
    if (IS_APP) await callNative("close", true);
    else window.close();
  } finally {
    closing = false;
  }
}
window.TR = { requestWindowClose: requestWindowClose, openHistory: () => openHistory(), setView: (v) => setView(v), quickShown: () => quickShown() };

/* ------------------------------------------------------------------ quick add */

let quickFiles = [];

function quickStage(list) {
  let rejected = 0;
  for (const file of Array.from(list || [])) {
    if (file.size > CFG.maxFileBytes) { rejected += 1; continue; }
    let name = file.name || "";
    if (!name || name === "image.png") name = pastedName(file.type);
    quickFiles.push({ file: file, name: name });
  }
  $("qMsg").textContent = rejected ? rejected + " file(s) over " + Math.round(CFG.maxFileBytes / 1048576) + " MB were not added." : "";
  renderQuickFiles();
}

function renderQuickFiles() {
  const box = $("qFiles");
  box.innerHTML = "";
  quickFiles.forEach((item, index) => {
    const chip = el("span", "quick-file");
    chip.appendChild(icon("paperclip", 12, 2));
    chip.appendChild(el("span", null, item.name));
    const drop = el("button");
    drop.type = "button";
    drop.title = "Remove";
    drop.appendChild(icon("x", 12, 2));
    drop.addEventListener("click", () => { quickFiles.splice(index, 1); renderQuickFiles(); });
    chip.appendChild(drop);
    box.appendChild(chip);
  });
  setDraft("quickText", $("qText").value);
}

function quickEmpty() {
  return !$("qText").value.trim() && !quickFiles.length;
}

function quickReset() {
  $("qText").value = "";
  quickFiles = [];
  $("qMsg").textContent = "";
  renderQuickFiles();
  clearDraft("quickText");
}

async function quickAdd() {
  const text = $("qText").value.trim();
  if (!text) { $("qMsg").textContent = "Type the task first."; $("qText").focus(); return false; }
  $("qAdd").disabled = true;
  try {
    const out = await api("/api/tasks/add", { date: $("qDate").value, project: $("qProject").value.trim(), text: text });
    if (!out.ok) { $("qMsg").textContent = out.message || "The task was not added."; return false; }
    const problems = [];
    for (const item of quickFiles) {
      const up = await uploadTaskFile(out.task.id, item.file, item.name);
      if (!up.ok) problems.push(item.name + ": " + (up.message || "not stored"));
    }
    try { localStorage.setItem("taskReporterQuickProject", $("qProject").value.trim()); } catch (err) { /* private */ }
    quickReset();
    if (problems.length) { $("qMsg").textContent = "Task added, but " + problems.join(" · "); return true; }
    await callNative("quick_hide");
    return true;
  } catch (err) {
    $("qMsg").textContent = "Lost contact with Task Reporter.";
    return false;
  } finally {
    $("qAdd").disabled = false;
  }
}

let quickAsking = false;
async function quickClose(fromBlur) {
  if (quickEmpty()) { quickReset(); await callNative("quick_hide"); return; }
  if (quickAsking) return;
  quickAsking = true;
  const answer = await askSaveChanges(fromBlur
    ? "This task has not been added yet. Add it, keep it open, or throw it away?"
    : "You have typed a task that has not been added. If you close now, it is lost.");
  quickAsking = false;
  if (answer === "save") await quickAdd();
  else if (answer === "discard") { quickReset(); await callNative("quick_hide"); }
  else $("qText").focus();
}

// Called by Python each time the tray (or the hotkey) shows the window.
function quickShown() {
  $("qDate").value = isoToday();
  $("qText").focus();
  loadBoard();
}

function setupQuick() {
  $("qDate").value = isoToday();
  let project = "";
  try { project = localStorage.getItem("taskReporterQuickProject") || ""; } catch (err) { /* private */ }
  $("qProject").value = project;
  if (drafts.quickText) $("qText").value = drafts.quickText;
  attachProjectCombo($("qProject"), $("qProjectPanel"), $("qProjectCaret"), () => $("qText").focus());
  $("qText").addEventListener("input", () => setDraft("quickText", $("qText").value));
  $("qText").addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey && !event.isComposing) { event.preventDefault(); quickAdd(); }
  });
  $("qAdd").addEventListener("click", quickAdd);
  $("qClose").addEventListener("click", () => quickClose(false));
  $("qOpenApp").addEventListener("click", () => callNative("open_main"));
  wireDropZone($("quickView"), quickStage);
  window.addEventListener("blur", () => {
    // A moment's grace: the tray click that opened the window can steal focus back.
    setTimeout(() => { if (!document.hasFocus() && !confirmResolve) quickClose(true); }, 150);
  });
  loadBoard();
}

/* -------------------------------------------------------------------- keyboard */

function inTextField(target) {
  if (!target) return false;
  const tag = target.tagName;
  return tag === "TEXTAREA" || target.isContentEditable ||
    (tag === "INPUT" && !["checkbox", "radio", "button", "submit"].includes(target.type));
}

document.addEventListener("keydown", (event) => {
  const accel = event.ctrlKey || event.metaKey;
  const top = topLayer();

  if (event.key === "Escape") {
    if (top === "confirmBackdrop") { event.preventDefault(); answerConfirm("keep"); return; }
    if (IS_QUICK) { event.preventDefault(); quickClose(false); return; }
    if (top) { event.preventDefault(); requestClose(top); return; }
    return;
  }
  if (top === "confirmBackdrop") return;

  if (accel && event.key === "Enter") {
    event.preventDefault();
    if (IS_QUICK) { quickAdd(); return; }
    if (top === "editBackdrop") saveEdit();
    else if (top === "taskEditBackdrop") saveTaskEdit();
    else if (top === "previewBackdrop") confirmFiling();
    else if (!top || top === "historyPanel") {
      // Same meaning on both views: commit what is in front of you.
      if (currentView === "board") startFiling();
      else if (currentView === "report") saveReport();
    }
    return;
  }
  if (IS_QUICK) return;
  if (accel && !event.shiftKey && (event.key === "h" || event.key === "H")) {
    event.preventDefault();
    if (historyOpen) closeHistory(); else openHistory();
    return;
  }
  if (accel && !event.shiftKey && (event.key === "b" || event.key === "B")) {
    if (top && top !== "historyPanel") return;
    event.preventDefault();
    setView(currentView === "board" ? "report" : "board");
    return;
  }
  // Ctrl+Z in a text box is the text's own undo; anywhere else it is the toast's.
  if (accel && !event.shiftKey && (event.key === "z" || event.key === "Z") && !inTextField(event.target) && toastUndo) {
    event.preventDefault();
    runToastUndo();
  }
});

/* ---------------------------------------------------------------- start-up */

function tickClock() {
  $("clock").textContent = stampNow();
}

function restoreEditDraft() {
  const draft = drafts.edit;
  if (!draft || IS_QUICK) return;
  if (draft.kind === "task") {
    if (!boardTasks.some((item) => item.id === draft.id)) { clearDraft("edit"); return; }
    if (currentView !== "board") setView("board");
    if (openTaskEdit(draft.id, draft)) {
      flashBoard("Restored unsaved changes", "warn");
    } else {
      clearDraft("edit");
    }
  } else if (draft.kind === "report") {
    api("/api/list").then((out) => {
      historyRows = out.reports || [];
      historyQueued = out.queued || [];
      const row = historyRows.find((item) => item.index === draft.index);
      // Only onto the same row it was typed against.
      if (row && row.datetime === draft.origWhen && row.text === draft.origText && openEdit(draft.index, draft)) {
        flashHere("Restored unsaved changes", "warn");
      } else {
        clearDraft("edit");
      }
    }).catch(() => {});
  }
}

function startMain() {
  editor.addEventListener("input", () => { updateCounter(); setDraft("report", editor.value); });
  saveBtn.addEventListener("click", saveReport);
  const revealBtn = $("revealBtn");
  revealBtn.title = "Show " + CFG.workbook + " in Explorer";
  revealBtn.addEventListener("click", () => openKnown("workbook", true));
  $("helpBtn").addEventListener("click", () => openModal("helpBackdrop"));
  $("tabReport").addEventListener("click", () => setView("report"));
  $("tabBoard").addEventListener("click", () => setView("board"));
  $("composer").addEventListener("submit", addTask);
  $("fileBtn").addEventListener("click", startFiling);
  $("clearFiledBtn").addEventListener("click", clearFiled);
  $("previewConfirm").addEventListener("click", confirmFiling);
  $("showFiled").addEventListener("change", (event) => { showFiledTasks = event.target.checked; renderBoard(); });

  const taskText = $("taskText");
  taskText.addEventListener("keydown", (event) => {
    // Enter adds; Shift+Enter is a new line inside the task.
    if (event.key === "Enter" && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.isComposing) {
      event.preventDefault();
      addTask();
    }
  });
  taskText.addEventListener("input", () => { fitTaskBox(taskText); setDraft("task", taskText.value); });

  attachProjectCombo($("taskProject"), $("taskProjectPanel"), $("taskProjectCaret"), () => taskText.focus());
  attachProjectCombo($("tProject"), $("tProjectPanel"), $("tProjectCaret"), () => $("tText").focus());

  // Drafts typed before the window last closed, and the project the last
  // run of tasks went under - tasks arrive in runs, so it is kept.
  if (drafts.report) { editor.value = drafts.report; }
  if (drafts.task) { taskText.value = drafts.task; }
  try { $("taskProject").value = localStorage.getItem("taskReporterProject") || ""; } catch (err) { /* private mode */ }
  $("taskProject").addEventListener("input", () => {
    try { localStorage.setItem("taskReporterProject", $("taskProject").value.trim()); } catch (err) { /* private mode */ }
  });

  tickClock();
  setInterval(tickClock, 1000);
  updateCounter();
  if (drafts.report) flashStatus("Restored the report you were writing", "warn");

  $("taskDate").value = isoToday();
  let startView = "report";
  try { startView = localStorage.getItem("taskReporterView") || "report"; } catch (err) { /* private mode */ }
  setView(startView === "board" ? "board" : "report");
  if (startView !== "board") loadBoard();
  syncThemeRadios();
}

if (IS_QUICK) setupQuick(); else startMain();
if (IS_APP) {
  if (native()) setupWindowChrome();
  else window.addEventListener("pywebviewready", setupWindowChrome);
}
ping();
setInterval(ping, CFG.pingSeconds * 1000);
</script>
</body>
</html>
"""


class WebSessionState:
    """What the request handlers share: the token, and who is looking.

    Client bookkeeping is per-tab rather than a single counter so that closing
    one of two open tabs does not end the session for the other.
    """

    def __init__(self, session: "ReporterSession", token: str):
        self.session = session
        self.token = token
        self._lock = threading.Lock()
        self._clients = {}          # client id -> monotonic time of last ping
        self._ever_connected = False
        self._page_served = False
        self._bye_at = None
        # Set when a second launch asks the running window to come forward.
        self._focus_requested = False
        # True while the page holds an edit that closing would throw away.
        self._page_dirty = False
        self._page_commands = deque()

    def set_page_dirty(self, dirty: bool):
        with self._lock:
            self._page_dirty = bool(dirty)

    @property
    def page_dirty(self) -> bool:
        with self._lock:
            return self._page_dirty

    def send_page_command(self, command: str):
        """Ask the page to do something on its next ping (e.g. open history)."""
        with self._lock:
            self._page_commands.append(command)

    def take_page_commands(self) -> list:
        with self._lock:
            commands = list(self._page_commands)
            self._page_commands.clear()
        return commands

    def request_focus(self):
        with self._lock:
            self._focus_requested = True

    def take_focus_request(self) -> bool:
        with self._lock:
            wanted = self._focus_requested
            self._focus_requested = False
        return wanted

    def note_page_served(self):
        with self._lock:
            self._page_served = True

    @property
    def page_served(self) -> bool:
        with self._lock:
            return self._page_served

    def note_ping(self, client_id: str):
        with self._lock:
            self._clients[client_id] = time.monotonic()
            self._ever_connected = True
            # A reload arrives as bye-then-ping; the ping cancels the goodbye.
            self._bye_at = None

    def note_bye(self, client_id: str):
        """Record that a page said goodbye.  Only a hint - see the grace above."""
        with self._lock:
            self._clients.pop(client_id, None)
            if not self._clients and self._bye_at is None:
                self._bye_at = time.monotonic()

    def shutdown_reason(self, allow_idle_timeout: bool):
        """Why the session should end, or None to keep going."""
        now = time.monotonic()
        with self._lock:
            if allow_idle_timeout:
                stale = [
                    client
                    for client, seen in self._clients.items()
                    if now - seen >= WEB_IDLE_TIMEOUT_SECONDS
                ]
                for client in stale:
                    del self._clients[client]
                if stale and not self._clients:
                    self._bye_at = now

            if not self._ever_connected or self._clients:
                return None
            if self._bye_at is None:
                return None
            if now - self._bye_at < WEB_BYE_GRACE_SECONDS:
                return None
        return "browser page closed"


class _ReporterHTTPServer(http.server.ThreadingHTTPServer):
    daemon_threads = True
    state = None
    # Windows lets a second socket bind a port that is already being listened
    # on when SO_REUSEADDR is set, and then splits new connections between the
    # two unpredictably.  Under WSL that is not hypothetical: wslrelay.exe
    # mirrors every port bound inside the VM onto Windows, so a session running
    # in WSL and the Windows app would otherwise fight over 8770 - and the
    # window would load whichever answered first, token and all.
    allow_reuse_address = os.name != "nt"


class _ReporterRequestHandler(http.server.BaseHTTPRequestHandler):
    server_version = "TaskReporter"
    sys_version = ""
    protocol_version = "HTTP/1.1"

    # ---------------------------------------------------------------- plumbing

    def log_message(self, fmt, *args):
        if os.environ.get("TASK_REPORT_WEB_DEBUG") == "1":
            sys.stderr.write(
                "  [web] %s %s\n"
                % (datetime.now().strftime("%H:%M:%S.%f")[:-3], fmt % args)
            )

    @property
    def _state(self) -> "WebSessionState":
        return self.server.state

    def _send(self, status: int, body: bytes, content_type: str):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        # The page never wants to be framed or sniffed, and it has no business
        # being fetched by anything other than itself.
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _send_json(self, payload: dict, status: int = 200):
        body = json.dumps(payload).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8")

    def _send_text(self, status: int, text: str):
        self._send(status, text.encode("utf-8"), "text/plain; charset=utf-8")

    def _query(self) -> dict:
        parsed = urllib.parse.urlparse(self.path)
        return urllib.parse.parse_qs(parsed.query)

    def _path(self) -> str:
        return urllib.parse.urlparse(self.path).path

    def _authorised(self) -> bool:
        """Reject anything that is not this session's own page.

        Two checks, for two different problems.  The token is the real gate:
        loopback is *not* private on WSL, because Windows forwards its own
        localhost into the VM, so any process on either side of the boundary
        can reach this port.  The Host check closes DNS rebinding, where a
        hostile page resolves its own domain to 127.0.0.1 and talks to us from
        the browser the user already trusts.
        """
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0].strip().lower()
        if host.strip("[]") not in ("127.0.0.1", "localhost", "::1"):
            self._send_text(403, "forbidden host")
            return False

        supplied = self.headers.get("X-Task-Report-Token") or ""
        if not supplied:
            values = self._query().get("t") or []
            supplied = values[0] if values else ""
        if not secrets.compare_digest(str(supplied), self._state.token):
            self._send_text(403, "forbidden")
            return False
        return True

    def _query_value(self, key: str) -> str:
        values = self._query().get(key) or []
        return values[0] if values else ""

    def _body_length(self) -> int:
        try:
            return max(0, int(self.headers.get("Content-Length") or 0))
        except ValueError:
            return 0

    def _discard_body(self, length: int):
        """Read and drop a body we will not use, so keep-alive stays in step."""
        while length > 0:
            chunk = self.rfile.read(min(length, 1 << 20))
            if not chunk:
                break
            length -= len(chunk)

    def _send_task_file(self):
        path = task_file_path(self._query_value("id"), self._query_value("name"))
        if not path or not os.path.isfile(path):
            self._send_text(404, "not found")
            return
        try:
            with open(path, "rb") as handle:
                body = handle.read()
        except OSError:
            self._send_text(404, "not found")
            return
        kind = mimetypes.guess_type(path)[0] or "application/octet-stream"
        self._send(200, body, kind)

    def _receive_task_file(self) -> dict:
        """POST /api/tasks/files/add - the one route whose body is a file.

        The file arrives raw (application/octet-stream) with the task id and
        file name in the query string, or as JSON {id, name, data_base64}.
        Either way the size is checked before the body is read.
        """
        length = self._body_length()
        # base64 is a third bigger than the file it carries.
        if length > MAX_TASK_FILE_BYTES * 4 // 3 + 4096:
            self._discard_body(length)
            return {
                "ok": False,
                "message": (
                    f"That file is too big. Task files are limited to "
                    f"{MAX_TASK_FILE_BYTES // 1048576} MB each."
                ),
            }
        raw = self.rfile.read(length) if length else b""
        kind = (self.headers.get("Content-Type") or "").split(";")[0].strip().lower()
        if kind == "application/json":
            try:
                payload = json.loads(raw.decode("utf-8"))
                task_id = str(payload.get("id") or "")
                name = str(payload.get("name") or "")
                data = base64.b64decode(str(payload.get("data_base64") or ""), validate=False)
            except Exception:
                return {"ok": False, "message": "The file did not arrive intact."}
        else:
            task_id = self._query_value("id")
            name = self._query_value("name")
            data = raw
        try:
            entry = add_task_file(task_id, name, data)
        except (ValueError, KeyError) as exc:
            return {"ok": False, "message": str(exc).strip("'\"")}
        except OSError as exc:
            return {"ok": False, "message": f"Could not store the file: {exc}"}
        return {"ok": True, "file": entry}

    def _read_json(self) -> dict:
        length = self._body_length()
        if length <= 0:
            return {}
        if length > WEB_MAX_BODY_BYTES:
            self._discard_body(length)
            return {}
        try:
            raw = self.rfile.read(length)
            data = json.loads(raw.decode("utf-8"))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    # ------------------------------------------------------------------ routes

    def do_GET(self):
        path = self._path()

        if path == "/favicon.ico":
            icon = _app_icon_bytes()
            if icon:
                self._send(200, icon, "image/x-icon")
            else:
                self._send(204, b"", "image/x-icon")
            return

        if not self._authorised():
            return

        if path in ("/", "/index.html"):
            self._state.note_page_served()
            page = _render_web_page(
                self._state.token,
                app=self._query_value("app") == "1",
                quick=self._query_value("quick") == "1",
            )
            self._send(200, page, "text/html; charset=utf-8")
            return

        if path == "/api/health":
            self._send_json({"ok": True})
            return

        if path == "/api/list":
            self._send_json(
                {"reports": _web_list_reports(), "queued": _web_list_queued()}
            )
            return

        if path == "/api/settings":
            self._send_json(_web_settings_state())
            return

        if path == "/api/tasks/files/raw":
            self._send_task_file()
            return

        if path == "/api/tasks":
            self._send_json(_web_board_state())
            return

        if path == "/api/tasks/preview":
            self._send_json(
                {
                    "groups": preview_filing(),
                    "workbookOpen": workbook_is_open_in_excel(),
                }
            )
            return

        self._send_text(404, "not found")

    def do_HEAD(self):
        self.do_GET()

    def do_POST(self):
        if not self._authorised():
            return

        path = self._path()

        if path == "/api/focus":
            # A second launch of the desktop app: there is only ever one
            # window, so the running one is brought forward instead.
            self._state.request_focus()
            self._send_json({"ok": True})
            return

        if path == "/api/bye":
            # Sent by sendBeacon on tab close, so it carries no body.
            values = self._query().get("client") or []
            self._state.note_bye(values[0] if values else "")
            self._send_json({"ok": True})
            return

        if path == "/api/tasks/files/add":
            self._send_json(self._receive_task_file())
            return

        payload = self._read_json()

        if path == "/api/ping":
            self._state.note_ping(str(payload.get("client") or ""))
            # Reports queued while Excel had the workbook go in as soon as it
            # lets go, rather than waiting for the next save to sweep them up.
            if pending_report_count() and not workbook_is_open_in_excel():
                flush_pending_reports()
            session = self._state.session
            events = [
                event
                for event in session.drain_events()
                if event.get("origin") not in ("browser", "window")
            ]
            self._send_json(
                {
                    "ok": True,
                    "events": events,
                    "shuttingDown": session.is_shutting_down(),
                    "reason": session.reason,
                    "queued": pending_report_count(),
                    "pendingCount": pending_report_count(),
                    "workbookOpen": workbook_is_open_in_excel(),
                    "boardRevision": board_revision(),
                    "commands": self._state.take_page_commands(),
                }
            )
            return

        if path == "/api/page-state":
            # The window's close button needs to know whether closing would
            # lose an edit; the page says so here every time that changes.
            self._state.set_page_dirty(bool(payload.get("dirty")))
            self._send_json({"ok": True})
            return

        if path == "/api/restore":
            self._send_json(_web_restore_report(payload))
            return

        if path == "/api/settings":
            self._send_json(_web_save_settings(payload))
            return

        if path == "/api/open":
            self._send_json(_web_open_known(payload, reveal=False))
            return

        if path == "/api/tasks/restore":
            tasks = payload.get("tasks")
            if tasks is None and isinstance(payload.get("task"), dict):
                tasks = [payload["task"]]
            self._send_json(_web_board_call(lambda: {"restored": restore_tasks(tasks or [])}))
            return

        if path == "/api/tasks/clear-filed/undo":
            self._send_json(_web_board_call(lambda: {"restored": restore_cleared_tasks()}))
            return

        if path == "/api/tasks/reorder":
            self._send_json(
                _web_board_call(
                    lambda: {"moved": reorder_tasks(payload.get("date"), payload.get("ids") or [])}
                )
            )
            return

        if path == "/api/tasks/file/undo":
            self._send_json(undo_last_filing())
            return

        if path == "/api/tasks/files/remove":
            self._send_json(_web_remove_task_file(payload))
            return

        if path == "/api/tasks/files/open":
            self._send_json(_web_task_file_action(payload, reveal=False))
            return

        if path == "/api/tasks/files/reveal":
            self._send_json(_web_task_file_action(payload, reveal=True))
            return

        if path == "/api/save":
            self._send_json(_web_save_report(self._state.session, payload))
            return

        if path == "/api/update":
            self._send_json(_web_update_report(payload))
            return

        if path == "/api/delete":
            self._send_json(_web_delete_report(payload))
            return

        if path == "/api/tasks/add":
            self._send_json(_web_add_task(payload))
            return

        if path == "/api/tasks/update":
            self._send_json(_web_update_task(payload))
            return

        if path == "/api/tasks/delete":
            self._send_json(_web_delete_task(payload))
            return

        if path == "/api/projects/forget":
            self._send_json(_web_forget_project(payload))
            return

        if path == "/api/reveal":
            self._send_json(_web_open_known(payload, reveal=True))
            return

        if path == "/api/tasks/clear-filed":
            self._send_json(_web_board_call(lambda: {"removed": delete_filed_tasks()}))
            return

        if path == "/api/tasks/file":
            self._send_json(file_checked_tasks(self._state.session, origin="browser"))
            return

        self._send_text(404, "not found")


# ---------------------------------------------------------------------------
# Browser UI actions - thin wrappers over the same workbook helpers the
# terminal console uses, so both surfaces cannot drift apart.
# ---------------------------------------------------------------------------


def _render_web_page(token: str, app: bool = False, quick: bool = False) -> bytes:
    config = {
        "token": token,
        "clientId": secrets.token_urlsafe(9),
        "maxLength": MAX_REPORT_LENGTH,
        "maxTaskLength": MAX_TASK_LENGTH,
        "maxCell": EXCEL_MAX_CELL,
        "maxFileBytes": MAX_TASK_FILE_BYTES,
        "pingSeconds": 3,
        "workbook": display_path(EXCEL_FILE_PATH),
        "pendingFile": display_path(PENDING_FILE_PATH),
        "taskFilesDir": display_path(TASK_FILES_DIR),
        "theme": load_settings()["theme"],
        # The frameless desktop window draws its own title bar; a browser tab
        # does not need one.
        "app": bool(app),
        "mode": "quick" if quick else "main",
    }
    page = (
        WEB_PAGE_TEMPLATE
        .replace("%%CONFIG%%", json.dumps(config))
        .replace("%%MAXTASK%%", str(MAX_TASK_LENGTH))
        .replace("%%MAXPROJECT%%", str(MAX_PROJECT_LENGTH))
    )
    return page.encode("utf-8")


def _web_list_reports() -> list:
    return [
        {"index": index, "datetime": when, "text": text}
        for index, (when, text) in enumerate(load_sheet_reports())
    ]


def _web_list_queued() -> list:
    return [{"datetime": when, "text": text} for when, text in list_pending_reports()]


def _app_icon_bytes() -> bytes:
    """windows/TaskReporter.ico - beside the script, or inside the bundle."""
    for folder in (
        getattr(sys, "_MEIPASS", ""),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "windows"),
    ):
        path = os.path.join(folder, "TaskReporter.ico") if folder else ""
        if path and os.path.isfile(path):
            try:
                with open(path, "rb") as handle:
                    return handle.read()
            except OSError:
                pass
    return b""


def _web_board_call(action) -> dict:
    """Run a board change and answer in the shape the page expects."""
    try:
        result = action()
    except ValueError as exc:
        return {"ok": False, "message": str(exc)}
    except KeyError as exc:
        return {"ok": False, "message": str(exc).strip("'\"")}
    except OSError as exc:
        return {"ok": False, "message": f"Could not write the task board: {exc}"}
    out = {"ok": True}
    out.update(result or {})
    return out


def _web_remove_task_file(payload: dict) -> dict:
    task_id = str(payload.get("id") or "")
    name = str(payload.get("name") or "")

    def action():
        if not remove_task_file(task_id, name):
            raise ValueError(f"{name} is not attached to that task.")
        return {}

    return _web_board_call(action)


def _web_task_file_action(payload: dict, reveal: bool) -> dict:
    path = task_file_path(payload.get("id"), payload.get("name"))
    if not path or not os.path.isfile(path):
        return {"ok": False, "message": "That file is not there any more."}
    return reveal_path(path) if reveal else open_path(path)


# What Settings and the header can show or open, by name - the page never
# sends a path, so it cannot be talked into opening anything else.
def _known_paths() -> dict:
    return {
        "workbook": EXCEL_FILE_PATH,
        "board": TASKS_FILE_PATH,
        "log": APP_LOG_PATH,
        "folder": os.path.join(BASE_DIR, ""),
        "taskFiles": os.path.join(TASK_FILES_DIR, ""),
    }


def _web_open_known(payload: dict, reveal: bool) -> dict:
    target = str(payload.get("target") or "workbook")
    path = _known_paths().get(target)
    if path is None:
        return {"ok": False, "message": "Nothing to show."}
    if target == "workbook" and reveal:
        return reveal_workbook()
    if target in ("folder", "taskFiles"):
        os.makedirs(path, exist_ok=True)
        return open_path(path.rstrip("\\/"))
    return reveal_path(path) if reveal else open_path(path)


def _file_card(path: str) -> dict:
    try:
        info = os.stat(path)
        return {"exists": True, "size": info.st_size,
                "modified": datetime.fromtimestamp(info.st_mtime).strftime(TIMESTAMP_FORMAT)}
    except OSError:
        return {"exists": False, "size": 0, "modified": ""}


def _web_settings_state() -> dict:
    settings = load_settings()
    sheet_rows = len(load_sheet_reports())
    tasks = load_tasks()
    return {
        "ok": True,
        "settings": settings,
        "dataDir": display_path(BASE_DIR),
        "dataDirSource": BASE_DIR_SOURCE,
        "envOverride": bool((os.environ.get("TASK_REPORT_DIR") or "").strip()),
        "restartNeeded": bool(settings["dataDir"])
        and os.path.abspath(settings["dataDir"]) != os.path.abspath(BASE_DIR),
        "files": {
            "workbook": dict(_file_card(EXCEL_FILE_PATH), rows=sheet_rows),
            "board": dict(_file_card(TASKS_FILE_PATH), tasks=len(tasks),
                          projects=len(known_projects())),
            "log": _file_card(APP_LOG_PATH),
        },
        "taskFilesDir": display_path(TASK_FILES_DIR),
        "workbookOpen": workbook_is_open_in_excel(),
        "pendingCount": pending_report_count(),
        "tray": TRAY_STATUS,
    }


def _web_save_settings(payload: dict) -> dict:
    changes = {}
    if payload.get("theme") in ("light", "dark", "system"):
        changes["theme"] = payload["theme"]
    if "keepInTray" in payload:
        changes["keepInTray"] = bool(payload["keepInTray"])
    if "hotkey" in payload:
        hotkey = str(payload.get("hotkey") or "").strip()
        if hotkey and parse_hotkey(hotkey) is None:
            return {"ok": False, "message": f"“{hotkey}” is not a shortcut Windows can register."}
        changes["hotkey"] = hotkey
    if "dataDir" in payload:
        folder = str(payload.get("dataDir") or "").strip()
        if folder and not os.path.isdir(folder):
            return {"ok": False, "message": f"{folder} is not a folder."}
        changes["dataDir"] = os.path.abspath(folder) if folder else ""
    try:
        save_settings(changes)
    except OSError as exc:
        return {"ok": False, "message": f"Could not save the settings: {exc}"}
    for listener in list(SETTINGS_LISTENERS):
        try:
            listener(changes)
        except Exception as exc:
            print(f"  [!] applying a setting failed: {exc}")
    return _web_settings_state()


# Called with the changed keys after every settings save - the tray and the
# hotkey live outside the page and need to hear about it.
SETTINGS_LISTENERS = []
TRAY_STATUS = {"available": False, "running": False, "hotkey": "", "hotkeyError": ""}


def _web_restore_report(payload: dict) -> dict:
    index = _web_row_index(payload)
    if index is None:
        return {"ok": False, "message": "That report could not be identified."}
    try:
        restore_report_row(index, str(payload.get("datetime") or ""), str(payload.get("text") or ""))
    except PermissionError as exc:
        return {"ok": False, "message": str(exc)}
    except Exception as exc:
        return {"ok": False, "message": f"Could not restore the report: {exc}"}
    return {"ok": True}


def _web_save_report(session: "ReporterSession", payload: dict) -> dict:
    text = str(payload.get("text") or "").strip()
    if not text:
        return {"ok": False, "message": "The report cannot be empty."}
    if len(text) > MAX_REPORT_LENGTH:
        return {
            "ok": False,
            "message": (
                f"Report is too long ({len(text)} chars). "
                f"The limit is {MAX_REPORT_LENGTH}."
            ),
        }
    try:
        timestamp = append_report_to_excel(text)
    except ReportQueuedError as exc:
        session.record_save("browser", exc.timestamp, queued=True)
        print(f"  [~] Filed from the browser at {exc.timestamp} (workbook locked).")
        return {"ok": True, "timestamp": exc.timestamp, "queued": True}
    except Exception as exc:
        return {"ok": False, "message": f"An unexpected error occurred: {exc}"}

    session.record_save("browser", timestamp)
    print(f"  [ok] Filed from the browser at {timestamp} -> {EXCEL_FILE_NAME}")
    return {"ok": True, "timestamp": timestamp, "queued": False}


def _web_board_state() -> dict:
    return {
        "tasks": load_tasks(),
        "projects": known_projects(),
        "counts": board_counts(),
        "revision": board_revision(),
    }


def _web_add_task(payload: dict) -> dict:
    try:
        task = add_task(
            payload.get("date"),
            str(payload.get("project") or ""),
            str(payload.get("text") or ""),
        )
    except ValueError as exc:
        return {"ok": False, "message": str(exc)}
    except OSError as exc:
        return {"ok": False, "message": f"Could not write the task board: {exc}"}
    return {"ok": True, "task": task}


def _web_update_task(payload: dict) -> dict:
    task_id = str(payload.get("id") or "").strip()
    if not task_id:
        return {"ok": False, "message": "That task could not be identified."}

    # Only the keys the page actually sent are applied, so ticking a box cannot
    # blank the text and editing the text cannot untick the box.
    changes = {}
    for key in ("text", "project", "date", "done"):
        if key in payload:
            changes[key] = payload[key]
    if not changes:
        return {"ok": False, "message": "Nothing to change."}

    try:
        task = update_task(task_id, **changes)
    except ValueError as exc:
        return {"ok": False, "message": str(exc)}
    except KeyError:
        return {"ok": False, "message": "That task no longer exists."}
    except OSError as exc:
        return {"ok": False, "message": f"Could not write the task board: {exc}"}
    return {"ok": True, "task": task}


def _web_forget_project(payload: dict) -> dict:
    name = str(payload.get("name") or "").strip()
    if not name:
        return {"ok": False, "message": "No project name was given."}
    try:
        forgotten = forget_project(name)
    except OSError as exc:
        return {"ok": False, "message": f"Could not write the task board: {exc}"}
    if not forgotten:
        return {"ok": False, "message": f"'{name}' was not in the list."}
    return {"ok": True}


def _web_delete_task(payload: dict) -> dict:
    task_id = str(payload.get("id") or "").strip()
    if not task_id:
        return {"ok": False, "message": "That task could not be identified."}
    try:
        removed = delete_task(task_id)
    except OSError as exc:
        return {"ok": False, "message": f"Could not write the task board: {exc}"}
    if not removed:
        return {"ok": False, "message": "That task no longer exists."}
    # Handed back whole so the page can offer Undo with /api/tasks/restore.
    return {"ok": True, "task": removed}


def _web_row_index(payload: dict):
    try:
        index = int(payload.get("index"))
    except (TypeError, ValueError):
        return None
    return index if index >= 0 else None


def _web_update_report(payload: dict) -> dict:
    index = _web_row_index(payload)
    if index is None:
        return {"ok": False, "message": "That report could not be identified."}
    text = str(payload.get("text") or "").strip()
    if not text:
        return {"ok": False, "message": "Report text cannot be empty."}
    if len(text) > MAX_REPORT_LENGTH:
        return {
            "ok": False,
            "message": (
                f"Report is too long ({len(text)} chars). "
                f"The limit is {MAX_REPORT_LENGTH}."
            ),
        }
    try:
        update_report_in_excel(index, str(payload.get("datetime") or "").strip(), text)
    except PermissionError:
        return {
            "ok": False,
            "message": (
                "Could not write the workbook - it is probably open in Excel. "
                "Close it and try again."
            ),
        }
    except Exception as exc:
        return {"ok": False, "message": f"An unexpected error occurred: {exc}"}
    return {"ok": True}


def _web_delete_report(payload: dict) -> dict:
    index = _web_row_index(payload)
    if index is None:
        return {"ok": False, "message": "That report could not be identified."}
    try:
        delete_report_from_excel(index)
    except PermissionError:
        return {
            "ok": False,
            "message": (
                "Could not write the workbook - it is probably open in Excel. "
                "Close it and try again."
            ),
        }
    except Exception as exc:
        return {"ok": False, "message": f"An unexpected error occurred: {exc}"}
    return {"ok": True}


# ---------------------------------------------------------------------------
# Starting the server and getting a browser pointed at it
# ---------------------------------------------------------------------------


def _port_is_answered(port: int) -> bool:
    """True when something already accepts connections on this loopback port.

    Checked before binding rather than after, because a bind that "succeeds"
    alongside an existing listener is the failure that cannot be recovered
    from: the connections are then split between the two servers.
    """
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.settimeout(0.25)
        try:
            return probe.connect_ex((WEB_HOST, port)) == 0
        except OSError:
            return False


def start_web_server(session: "ReporterSession", port_hint: int = None):
    """Bind the UI server.  Returns (server, state, error).

    A busy port is never allowed to be the reason the UI does not come up: the
    preferred range is tried in order and then the OS picks one.
    """
    state = WebSessionState(session, secrets.token_urlsafe(24))

    if port_hint:
        candidates = [port_hint]
    else:
        candidates = list(WEB_PREFERRED_PORTS) + [0]

    last_error = None
    for port in candidates:
        if port and _port_is_answered(port):
            # Something is already listening - very likely another Task
            # Reporter session, whose page must not be handed our token.
            last_error = OSError(f"port {port} is already in use")
            continue
        try:
            server = _ReporterHTTPServer(
                (WEB_HOST, port), _ReporterRequestHandler
            )
        except OSError as exc:
            last_error = exc
            continue
        server.state = state
        thread = threading.Thread(
            target=server.serve_forever,
            kwargs={"poll_interval": 0.2},
            name="task-report-web",
            daemon=True,
        )
        thread.start()
        return server, state, None

    if port_hint:
        return None, None, f"port {port_hint} is not available ({last_error})"
    return None, None, f"no loopback port could be bound ({last_error})"


def is_wsl() -> bool:
    if os.environ.get("WSL_DISTRO_NAME") or os.environ.get("WSL_INTEROP"):
        return True
    try:
        with open("/proc/version", "r", encoding="utf-8", errors="replace") as fh:
            return "microsoft" in fh.read().lower()
    except OSError:
        return False


def web_url(port: int, token: str, host: str = None) -> str:
    if host is None:
        # Windows reaches into the VM through the name `localhost`; a bare
        # 127.0.0.1 also works, and is the fallback if name resolution on the
        # Windows side prefers ::1 (where nothing is listening).
        host = "localhost" if is_wsl() else WEB_HOST
    return f"http://{host}:{port}/?t={token}"


def _browser_launchers(url: str) -> list:
    """Ways to open a URL, best first.  Each entry is (label, argv)."""
    launchers = []

    override = (os.environ.get("TASK_REPORT_BROWSER") or "").strip()
    if override:
        launchers.append((override, [*override.split(), url]))

    if is_wsl():
        # Handing the URL to Windows is the whole point - the page must be
        # drawn by a Windows browser, not by anything inside the VM.
        launchers.append(
            (
                "Windows default browser",
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    f"Start-Process '{url}'",
                ],
            )
        )
        launchers.append(("cmd start", ["cmd.exe", "/c", "start", "", url]))
        # explorer.exe reports failure even when it succeeds, so it is last and
        # its exit code is ignored (see _run_launcher).
        launchers.append(("explorer", ["explorer.exe", url]))

    for tool, label in (
        ("wslview", "wslview"),
        ("xdg-open", "xdg-open"),
        ("gio", "gio open"),
        ("open", "open"),
    ):
        resolved = shutil.which(tool)
        if not resolved:
            continue
        argv = [resolved, "open", url] if tool == "gio" else [resolved, url]
        launchers.append((label, argv))

    return launchers


def _run_launcher(label: str, argv: list) -> bool:
    try:
        result = subprocess.run(
            argv,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=20,
            check=False,
        )
    except Exception:
        return False
    # explorer.exe returns 1 on success; everything else is trusted as normal.
    if argv[0].endswith("explorer.exe"):
        return True
    return result.returncode == 0


def open_in_browser(url: str) -> str:
    """Ask the desktop to open `url`.  Returns the label that accepted it."""
    for label, argv in _browser_launchers(url):
        if _run_launcher(label, argv):
            return label
    try:
        if webbrowser.open(url, new=2):
            return "python webbrowser"
    except Exception:
        pass
    return ""


def _windows_path(path: str) -> str:
    """Translate a Linux path into the Windows spelling, or "" if it cannot be.

    Only meaningful under WSL, where Explorer is a Windows program and has
    never heard of /c/... or /home/...  wslpath handles both, including the
    \\wsl$\ form for files that live inside the VM.
    """
    try:
        result = subprocess.run(
            ["wslpath", "-w", path],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except Exception:
        return ""
    return result.stdout.strip() if result.returncode == 0 else ""


def _reveal_launchers(path: str) -> list:
    """Ways to show `path` in a file manager, best first.  (label, argv) pairs.

    Selecting the file beats opening the folder, so the exact file being talked
    about is the one highlighted - but a folder is offered as the fallback,
    because the workbook does not exist until the first report is filed.
    """
    folder = os.path.dirname(path) or "."
    exists = os.path.exists(path)
    launchers = []

    if os.name == "nt":
        explorer = shutil.which("explorer.exe") or "explorer.exe"
        if exists:
            # One argument, comma and all: Explorer does not accept /select as
            # a separate token.
            launchers.append(("Explorer", [explorer, f"/select,{path}"]))
        launchers.append(("Explorer", [explorer, folder]))
        return launchers

    if sys.platform == "darwin":
        if exists:
            launchers.append(("Finder", ["open", "-R", path]))
        launchers.append(("Finder", ["open", folder]))
        return launchers

    if is_wsl():
        explorer = shutil.which("explorer.exe")
        if explorer:
            windows_file = _windows_path(path) if exists else ""
            windows_folder = _windows_path(folder)
            if windows_file:
                launchers.append(("Explorer", [explorer, f"/select,{windows_file}"]))
            if windows_folder:
                launchers.append(("Explorer", [explorer, windows_folder]))

    for label, tool in (("xdg-open", "xdg-open"), ("gio open", "gio")):
        resolved = shutil.which(tool)
        if not resolved:
            continue
        launchers.append(
            (label, [resolved, "open", folder] if tool == "gio" else [resolved, folder])
        )

    return launchers


def reveal_workbook() -> dict:
    """Show task_reports.xlsx in the desktop's file manager."""
    return reveal_path(EXCEL_FILE_PATH, "the reports folder")


def reveal_path(path: str, folder_label: str = "its folder") -> dict:
    """Show `path` selected in the desktop's file manager.

    Answers in the shape the browser UI expects, and never raises: not being
    able to open a window is worth a message, not a broken button.
    """
    folder = os.path.dirname(path)

    launchers = _reveal_launchers(path)
    if not launchers:
        return {
            "ok": False,
            "path": path,
            "message": f"No file manager could be found. It is in {folder}",
        }

    for label, argv in launchers:
        if _run_launcher(label, argv):
            where = folder_label if not os.path.exists(path) else os.path.basename(path)
            return {
                "ok": True,
                "path": path,
                "message": f"Showing {where} in {label}",
            }

    return {
        "ok": False,
        "path": path,
        "message": f"The file manager would not open. It is in {folder}",
    }


def open_path(path: str) -> dict:
    """Open a file with whatever the desktop opens that kind of file with."""
    if not os.path.exists(path):
        return {"ok": False, "message": f"{os.path.basename(path)} is not there any more."}
    name = os.path.basename(path)
    if os.name == "nt":
        try:
            os.startfile(path)  # noqa: the Windows-only call is the point
            return {"ok": True, "message": f"Opened {name}"}
        except OSError as exc:
            return {"ok": False, "message": f"Windows would not open {name}: {exc}"}

    launchers = []
    if sys.platform == "darwin":
        launchers.append(("open", ["open", path]))
    if is_wsl():
        windows_file = _windows_path(path)
        if windows_file and shutil.which("cmd.exe"):
            # `start` picks the default app; the empty title is required.
            launchers.append(("Windows", ["cmd.exe", "/c", "start", "", windows_file]))
    for tool in ("wslview", "xdg-open"):
        resolved = shutil.which(tool)
        if resolved:
            launchers.append((tool, [resolved, path]))
    for label, argv in launchers:
        if _run_launcher(label, argv):
            return {"ok": True, "message": f"Opened {name}"}
    return {"ok": False, "message": f"Nothing would open {name}. It is in {os.path.dirname(path)}"}


@functools.lru_cache(maxsize=64)
def display_path(path: str) -> str:
    """The path as the person at the keyboard would write it.

    Under WSL that is the Windows spelling - a path copied out of the app is
    going to be pasted into something running on Windows.
    """
    if is_wsl():
        windows = _windows_path(path)
        if windows:
            return windows
    return path


def _wait_for_page(state: "WebSessionState", timeout: float) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if state.page_served:
            return True
        time.sleep(0.15)
    return state.page_served


# ---------------------------------------------------------------------------
# Terminal console
# ---------------------------------------------------------------------------

CLI_PROMPT = "report> "

CLI_COMMANDS = [
    (":m  /  :multi", "write a multi-line report (finish with a lone '.')"),
    (":l  /  :list", "show the 10 most recent reports"),
    (":t  /  :tasks", "the task board - see the board commands below"),
    (":p  /  :path", "print the workbook path"),
    (":h  /  :help", "show this help"),
    (":q  /  :quit", "close the session (Ctrl+C and Ctrl+D do the same)"),
]


def _print_cli_help():
    print("  Type your report and press Enter to file it.")
    print("  Use \\n inside the line for a manual line break.")
    for keys, desc in CLI_COMMANDS:
        print(f"    {keys:<16} {desc}")
    print("  Task board:")
    for keys, desc in CLI_BOARD_COMMANDS:
        print(f"    {keys:<22} {desc}")


def _print_cli_banner(dual: bool):
    print()
    print("=" * 68)
    print("  TASK REPORTER - terminal console")
    print("=" * 68)
    if dual:
        print("  The browser UI and this console are both live.")
        print("  File the report in whichever one you like.")
        print("  Closing either one closes the other.")
    else:
        print("  This console is the only surface for this session.")
    print(f"  Workbook: {EXCEL_FILE_PATH}")
    counts = board_counts()
    if counts["total"]:
        print(
            f"  Task board: {counts['open']} open, {counts['ready']} ticked "
            f"and ready to file  (':t' to list)"
        )
    queued = pending_report_count()
    if queued:
        print(f"  {queued} report(s) waiting to be merged into the workbook.")
    if workbook_is_open_in_excel():
        print("  NOTE: the workbook is open in Excel. Reports are queued until")
        print("        you close it, rather than being written and then lost.")
    print("-" * 68)
    _print_cli_help()
    print("-" * 68)


def cli_merge_days(dry_run: bool = False) -> int:
    """--merge-days: one row per day for a workbook that predates that rule."""
    plan = plan_day_merge()
    if not plan:
        print("  Every day in the workbook is already a single row.")
        return 0

    print(f"  {len(plan)} day(s) are spread over more than one row:")
    for day in plan:
        print(
            f"    {day['dateLabel']}  {day['rows']} rows -> 1 row "
            f"({day['chars']} chars)"
        )

    if dry_run:
        print()
        print("  Nothing was written (--dry-run). Drop --dry-run to merge them.")
        return 0

    if workbook_is_open_in_excel():
        print()
        print(f"  {EXCEL_FILE_NAME} is open in Excel. Close it and run this again -")
        print("  rewriting the sheet now would be undone the next time Excel saves.")
        return 1

    try:
        result = merge_existing_days()
    except PermissionError as exc:
        print(f"  {exc}")
        return 1
    except Exception as exc:
        print(f"  The workbook could not be rewritten: {exc}")
        return 1

    print()
    print(
        f"  Merged {result['merged']} day(s), {result['rowsRemoved']} fewer row(s)."
    )
    return 0


def _print_recent_reports(limit: int = 10):
    rows = load_reports_from_excel()
    if not rows:
        print("  No reports yet.")
        return
    recent = rows[-limit:]
    start = len(rows) - len(recent) + 1
    print(f"  Showing {len(recent)} of {len(rows)} report(s):")
    for offset, (timestamp, report) in enumerate(recent):
        single_line = " ".join(report.split())
        if len(single_line) > 96:
            single_line = single_line[:93] + "..."
        print(f"    {start + offset:>4}. [{timestamp}] {single_line}")


def save_report_text(text: str, origin: str, session=None) -> bool:
    """Save one report from either surface.  Returns True when it was kept."""
    try:
        timestamp = append_report_to_excel(text)
    except ValueError as exc:
        print(f"  [!] {exc}")
        return False
    except ReportQueuedError as exc:
        if session is not None:
            session.record_save(origin, exc.timestamp, queued=True)
        print(f"  [~] Saved at {exc.timestamp}, but the workbook is locked.")
        print(f"      Queued in {os.path.basename(PENDING_FILE_PATH)} and it will")
        print("      be merged automatically once Excel releases the file.")
        return True
    except Exception as exc:
        print(f"  [!] Could not save: {exc}")
        return False

    if session is not None:
        session.record_save(origin, timestamp)
    print(f"  [ok] Saved at {timestamp} -> {EXCEL_FILE_NAME}")
    return True


def _read_multiline_report() -> str:
    print("  Multi-line mode. Finish with a single '.' on its own line.")
    lines = []
    while True:
        try:
            line = input("  | ")
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if line.strip() == ".":
            break
        lines.append(line)
    return "\n".join(lines).strip()


# ---------------------------------------------------------------------------
# The task board from the terminal
# ---------------------------------------------------------------------------
#
# Same board, same file, same filing - just typed instead of clicked, so a
# session that never opened a browser is not a session without a board.  The
# add syntax deliberately mirrors the task list itself:
#
#   :t add [onex-academy] fix the links comparison issue
#   :t add 21.08.2026 [onex-academy] fix the links comparison issue

CLI_BOARD_COMMANDS = [
    (":t  /  :tasks", "show the task board"),
    (":t all", "show it including tasks already filed"),
    (":t add [proj] TEXT", "add a task to today ([proj] and a leading date"),
    ("", "are both optional: :t add 21.08.2026 [proj] TEXT)"),
    (":t x N...", "tick tasks by number (also :t done N)"),
    (":t o N...", "untick tasks by number (also :t open N)"),
    (":t rm N...", "delete tasks by number"),
    (":t file", "write every ticked task into the workbook"),
    (":t clear", "drop already-filed tasks off the board"),
    (":t projects", "list the remembered project names"),
    (":t forget NAME", "drop a project name off that list"),
]

# Numbers refer to the listing the console last printed, so ":t x 3" always
# means the line the user is looking at, whatever has changed since.
_CLI_LISTING = []

_CLI_DATE_TOKEN = re.compile(r"^\d{1,4}[./-]\d{1,2}[./-]\d{1,4}$")


def _relative_day_text(date_iso: str) -> str:
    try:
        day = datetime.strptime(date_iso, DATE_STORE_FORMAT).date()
    except (ValueError, TypeError):
        return ""
    diff = (day - datetime.now().date()).days
    if diff == 0:
        return "today"
    if diff == -1:
        return "yesterday"
    if diff == 1:
        return "tomorrow"
    name = day.strftime("%A")
    if diff < 0:
        return f"{name}, {-diff} days ago"
    return f"{name}, in {diff} days"


def _print_board(show_filed: bool = False):
    """Print the board, and remember the numbering it used."""
    global _CLI_LISTING

    tasks = load_tasks()
    counts = board_counts()

    if not tasks:
        print("  The task board is empty.")
        print("  Add one with:  :t add [project] what needs doing")
        _CLI_LISTING = []
        return

    shown = tasks if show_filed else [t for t in tasks if not t["filed_at"]]
    hidden = len(tasks) - len(shown)

    print(
        f"  Task board - {counts['total']} task(s), {counts['open']} open, "
        f"{counts['ready']} ticked and ready to file"
    )
    print("  " + "-" * 66)

    _CLI_LISTING = []
    if not shown:
        print("  Everything on the board is filed.  ':t all' shows it anyway.")
    else:
        for date_iso, projects in group_by_day(shown):
            day = [task for _name, group in projects for task in group]
            done = sum(1 for task in day if task["done"])
            when = _relative_day_text(date_iso)
            print()
            print(
                f"  {display_date(date_iso)}"
                + (f"  ({when})" if when else "")
                + f"  -  {done} of {len(day)} done"
            )
            for name, group in projects:
                print(f"    [{name}]" if name else "    (no project)")
                for task in group:
                    _CLI_LISTING.append(task["id"])
                    text = " ".join(task["text"].split())
                    suffix = "  (filed)" if task["filed_at"] else ""
                    room = 58 - len(suffix)
                    if len(text) > room:
                        text = text[: room - 3] + "..."
                    mark = "x" if task["done"] else " "
                    print(f"      {len(_CLI_LISTING):>3}. [{mark}] {text}{suffix}")

    print()
    print("  " + "-" * 66)
    if hidden:
        print(f"  {hidden} filed task(s) hidden - ':t all' shows them, "
              "':t clear' removes them.")
    if counts["ready"]:
        rows = len(preview_filing())
        print(
            f"  {counts['ready']} ticked -> ':t file' writes "
            f"{rows} report row(s)."
        )
    else:
        print("  Nothing ticked yet.  Tick with ':t x N'.")


def _print_projects():
    names = known_projects()
    if not names:
        print("  No project names remembered yet.")
        print("  One is recorded the first time you use it:")
        print("    :t add [onex-academy] fix the links comparison issue")
        return
    print(f"  {len(names)} project name(s), most recently used first:")
    for index, name in enumerate(names, 1):
        print(f"    {index:>3}. [{name}]")
    print("  ':t forget NAME' drops one off the list.")


def _cli_forget_project(name: str):
    name = name.strip()
    # The bracketed form is how these are written everywhere else, so accept it.
    if name.startswith("[") and name.endswith("]"):
        name = name[1:-1].strip()
    if not name:
        print("  [!] Which one?  e.g.  :t forget onex-academy")
        return
    try:
        forgotten = forget_project(name)
    except OSError as exc:
        print(f"  [!] Could not write the task board: {exc}")
        return
    if forgotten:
        print(f"  [ok] '{name}' removed from the project list.")
        print("       Tasks already using it keep it.")
    else:
        print(f"  [!] '{name}' is not in the list.  ':t projects' shows it.")


def _resolve_task_numbers(words: list):
    """Turn the numbers the user typed into task ids.  Returns (ids, bad)."""
    if not _CLI_LISTING:
        _print_board()
        print()
    ids = []
    bad = []
    for word in words:
        try:
            number = int(word)
        except ValueError:
            bad.append(word)
            continue
        if 1 <= number <= len(_CLI_LISTING):
            ids.append(_CLI_LISTING[number - 1])
        else:
            bad.append(word)
    return ids, bad


def parse_task_line(rest: str):
    """`[21.08.2026] [project] the task text` -> (date, project, text).

    Both prefixes are optional and in that order, which is how the task list
    writes them: the day heading, then the project in brackets, then the task.
    """
    date = None
    project = ""
    text = rest.strip()

    head = text.split(None, 1)
    if head and _CLI_DATE_TOKEN.match(head[0]):
        date = head[0]
        text = head[1].strip() if len(head) > 1 else ""

    if text.startswith("["):
        close = text.find("]")
        if close != -1:
            project = text[1:close].strip()
            text = text[close + 1:].strip()

    return date, project, text


def _cli_add_task(rest: str) -> bool:
    date, project, text = parse_task_line(rest)
    if not text:
        print("  [!] Nothing to add.")
        print("      Try:  :t add [onex-academy] fix the links comparison issue")
        return False
    try:
        task = add_task(date or today_iso(), project, text)
    except ValueError as exc:
        print(f"  [!] {exc}")
        return False
    except OSError as exc:
        print(f"  [!] Could not write the task board: {exc}")
        return False
    where = f"[{task['project']}] " if task["project"] else ""
    print(f"  [ok] Added to {display_date(task['date'])} {where}-> {TASKS_FILE_NAME}")
    return True


def _cli_mark_tasks(words: list, done: bool):
    if not words:
        print(f"  [!] Which one?  e.g.  :t {'x' if done else 'o'} 3   "
              f"or  :t {'x' if done else 'o'} 3 4 5")
        return
    ids, bad = _resolve_task_numbers(words)
    for word in bad:
        print(f"  [!] '{word}' is not a task number from the list above.")

    # A filed task is already a row in the workbook, so its box means nothing
    # now - the browser greys the same box out for the same reason.
    filed = {task["id"] for task in load_tasks() if task["filed_at"]}
    skipped = [task_id for task_id in ids if task_id in filed]
    ids = [task_id for task_id in ids if task_id not in filed]
    if skipped:
        print(f"  [!] Skipped {len(skipped)} already-filed task(s) - they are "
              "in the workbook already.")

    changed = 0
    for task_id in ids:
        try:
            update_task(task_id, done=done)
            changed += 1
        except (KeyError, ValueError, OSError) as exc:
            print(f"  [!] {exc}")
    if changed:
        verb = "Ticked" if done else "Unticked"
        ready = board_counts()["ready"]
        print(f"  [ok] {verb} {changed} task(s). {ready} ready to file.")


def _cli_remove_tasks(words: list):
    if not words:
        print("  [!] Which one?  e.g.  :t rm 3   or  :t rm 3 4 5")
        return
    ids, bad = _resolve_task_numbers(words)
    for word in bad:
        print(f"  [!] '{word}' is not a task number from the list above.")
    removed = 0
    for task_id in ids:
        try:
            if delete_task(task_id):
                removed += 1
        except OSError as exc:
            print(f"  [!] Could not write the task board: {exc}")
    if removed:
        print(f"  [ok] Deleted {removed} task(s).")
        # The numbering just moved, so the remembered listing is now a lie.
        _CLI_LISTING.clear()


def _print_filing_preview(groups: list):
    grown = sum(1 for group in groups if group["appendsToExisting"])
    fresh = len(groups) - grown
    where = []
    if fresh:
        where.append(f"{fresh} new row(s)")
    if grown:
        where.append(f"{grown} day(s) added to")
    print(
        f"  {sum(g['taskCount'] for g in groups)} ticked task(s) "
        f"-> {' + '.join(where)}:"
    )
    for group in groups:
        print()
        print(f"    {group['timestamp']}   ({group['taskCount']} task(s), "
              f"{group['length']} chars)")
        if group["appendsToExisting"]:
            print(
                f"      goes a blank line under the {group['existingLength']} "
                f"characters already filed for {group['dateLabel']}"
            )
        for line in group["text"].splitlines():
            print(f"      | {line}" if line else "      |")


def _report_filing_result(result: dict):
    for row in result.get("written") or []:
        status = "queued (workbook locked)" if row["queued"] else "filed"
        print(
            f"  [{'~' if row['queued'] else 'ok'}] {row['dateLabel']} "
            f"{status} at {row['timestamp']} ({row['taskCount']} task(s))"
        )
    for failure in result.get("failed") or []:
        print(f"  [!] {failure['dateLabel']}: {failure['message']}")
    if result.get("queued"):
        print(f"      Queued in {os.path.basename(PENDING_FILE_PATH)}; it merges")
        print("      itself in once Excel releases the file.")
    if not result.get("written"):
        message = result.get("message")
        if message:
            print(f"  [!] {message}")


def cli_file_checked_tasks(session=None, confirm: bool = True) -> bool:
    groups = preview_filing()
    if not groups:
        print("  [!] Nothing ticked.  Tick what you finished with ':t x N'.")
        return False

    _print_filing_preview(groups)
    print()
    if workbook_is_open_in_excel():
        print(f"  NOTE: {EXCEL_FILE_NAME} is open in Excel, so these go to the")
        print("        pending queue and merge in once you close it.")
        print()

    if confirm:
        try:
            answer = input(f"  Write {len(groups)} row(s) into "
                           f"{EXCEL_FILE_NAME}? [y/N] ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            answer = ""
        if answer not in ("y", "yes"):
            print("  Nothing was written.")
            return False

    result = file_checked_tasks(session, origin="terminal")
    _report_filing_result(result)
    _CLI_LISTING.clear()
    return bool(result.get("written"))


def handle_task_command(rest: str, session=None):
    """Everything after ':t'.  An empty rest just prints the board."""
    words = rest.split()
    if not words:
        _print_board()
        return

    action = words[0].lower()
    tail = rest[len(words[0]):].strip()

    if action in ("all", "-a", "--all"):
        _print_board(show_filed=True)
    elif action in ("add", "a", "+"):
        if _cli_add_task(tail):
            _CLI_LISTING.clear()
    elif action in ("x", "done", "tick"):
        _cli_mark_tasks(words[1:], True)
    elif action in ("o", "open", "undone", "untick"):
        _cli_mark_tasks(words[1:], False)
    elif action in ("rm", "del", "delete"):
        _cli_remove_tasks(words[1:])
    elif action in ("projects", "proj"):
        _print_projects()
    elif action == "forget":
        _cli_forget_project(tail)
    elif action == "file":
        cli_file_checked_tasks(session)
    elif action == "clear":
        try:
            removed = delete_filed_tasks()
        except OSError as exc:
            print(f"  [!] Could not write the task board: {exc}")
            return
        print(
            f"  [ok] Dropped {removed} filed task(s) off the board."
            if removed
            else "  Nothing filed to clear."
        )
        _CLI_LISTING.clear()
    else:
        print(f"  [!] ':t {action}' is not a board command.")
        for keys, desc in CLI_BOARD_COMMANDS:
            print(f"    {keys:<22} {desc}")


def cli_console_loop(session: ReporterSession, dual: bool):
    """Read reports from the terminal until the session ends."""
    session.cli_active = True
    _print_cli_banner(dual)

    while not session.is_shutting_down():
        try:
            line = input(CLI_PROMPT)
        except EOFError:
            print()
            session.request_shutdown("terminal closed")
            break
        except KeyboardInterrupt:
            print()
            session.request_shutdown("terminal interrupted (Ctrl+C)")
            break
        except Exception:
            session.request_shutdown("terminal closed")
            break

        command = line.strip()
        if not command:
            continue

        lowered = command.lower()
        if lowered in (":q", ":quit", ":exit"):
            session.request_shutdown("closed from the terminal")
            break
        if lowered in (":h", ":help", ":?"):
            _print_cli_help()
            continue
        if lowered in (":l", ":list"):
            _print_recent_reports()
            continue
        if lowered in (":t", ":tasks", ":board") or lowered.startswith(
            (":t ", ":tasks ", ":board ")
        ):
            handle_task_command(command.split(None, 1)[1] if " " in command else "",
                                session)
            continue
        if lowered in (":p", ":path"):
            print(f"  {EXCEL_FILE_PATH}")
            continue
        if lowered in (":m", ":multi"):
            text = _read_multiline_report()
            if text:
                save_report_text(text, "terminal", session)
            else:
                print("  [!] Nothing entered - not saved.")
            continue

        report = command.replace("\\n", "\n").strip()
        if len(report) > MAX_REPORT_LENGTH:
            print(
                f"  [!] Too long ({len(report)} chars). "
                f"The limit is {MAX_REPORT_LENGTH}."
            )
            continue

        save_report_text(report, "terminal", session)

    session.cli_active = False


# ---------------------------------------------------------------------------
# Run modes
# ---------------------------------------------------------------------------


def _install_signal_handlers(session: ReporterSession):
    """Closing the terminal window sends SIGHUP - treat it as 'end session'."""

    def _handler(signum, _frame):
        names = {
            getattr(signal, "SIGHUP", None): "terminal window closed",
            getattr(signal, "SIGTERM", None): "session terminated",
            getattr(signal, "SIGINT", None): "interrupted (Ctrl+C)",
        }
        session.request_shutdown(names.get(signum) or f"signal {signum}")

    for name in ("SIGHUP", "SIGTERM", "SIGINT"):
        sig = getattr(signal, name, None)
        if sig is None:
            continue
        try:
            signal.signal(sig, _handler)
        except (ValueError, OSError):
            pass


def _describe_gui_failure(attempts: list) -> str:
    if not attempts:
        return "no display was detected"
    return "; ".join(f"{name}: {detail}" for name, _ok, detail in attempts)


def run_gui(session: ReporterSession, start_cli: bool) -> int:
    qt_app = QApplication(sys.argv)
    qt_app.setApplicationName("Task Reporter")
    qt_app.setStyleSheet(DARK_QSS)
    qt_app.setQuitOnLastWindowClosed(True)

    window = TaskReporterApp(session=session)
    window.show()

    # Polls the shared shutdown flag so the terminal side can close the window,
    # and gives the interpreter a slice in which to run signal handlers - Qt's
    # exec() otherwise blocks in C and Ctrl+C never arrives.
    watchdog = QTimer()
    watchdog.setInterval(200)
    watchdog.timeout.connect(
        lambda: qt_app.quit() if session.is_shutting_down() else None
    )
    watchdog.start()

    if start_cli:
        thread = threading.Thread(
            target=cli_console_loop,
            args=(session, True),
            name="task-report-cli",
            daemon=True,
        )
        # Start once the window is actually up, so the banner is not buried
        # under Qt's own start-up chatter.
        QTimer.singleShot(0, thread.start)

    exit_code = qt_app.exec()
    session.request_shutdown("window closed")
    return exit_code


def wait_for_quiet_workbook(timeout: float = 15.0):
    """Block until no save is in flight.

    The hard exit below kills the console thread wherever it happens to be, so
    it must not land in the middle of openpyxl rewriting the .xlsx.
    """
    if _EXCEL_LOCK.acquire(timeout=timeout):
        _EXCEL_LOCK.release()


def run_stdin_oneshot(session: ReporterSession) -> int:
    """Non-interactive stdin: treat everything piped in as a single report."""
    try:
        text = sys.stdin.read().strip()
    except Exception:
        text = ""
    if not text:
        print("Error: The report cannot be empty.")
        return 1
    return 0 if save_report_text(text, "stdin", session) else 1


def _redraw_cli_prompt(session: ReporterSession):
    """Put the prompt back after a background thread has printed over it."""
    if session.cli_active and not session.is_shutting_down():
        sys.stdout.write("\n" + CLI_PROMPT)
        sys.stdout.flush()


def _print_web_banner(url: str, launcher: str, opened: bool):
    print()
    print("=" * 68)
    print("  TASK REPORTER - browser UI")
    print("=" * 68)
    if opened and launcher:
        print(f"  Opening in your browser via {launcher}.")
    elif opened:
        print("  No browser opener answered - open the address below yourself.")
    else:
        print("  Browser launch was skipped (--no-browser).")
    print(f"  Address : {url}")
    print("  This address is single-use: the token changes every session.")
    print("-" * 68)


def _ensure_page_loaded(session: ReporterSession, state: WebSessionState, port: int, primary_url: str):
    """Make sure a browser really did get the page, and say what to do if not.

    The one loopback quirk worth retrying: some Windows setups resolve
    `localhost` to ::1 first, where nothing is listening, while the literal
    127.0.0.1 goes straight through WSL's forwarder.
    """
    if _wait_for_page(state, WEB_FIRST_CLIENT_TIMEOUT):
        return

    alternate = web_url(port, state.token, host=WEB_HOST)
    if alternate != primary_url:
        print()
        print(f"  The page has not loaded yet - retrying with {alternate}")
        if open_in_browser(alternate) and _wait_for_page(
            state, WEB_FIRST_CLIENT_TIMEOUT
        ):
            _redraw_cli_prompt(session)
            return

    print()
    print("  The browser did not load the page on its own.")
    print("  Copy one of these into any browser window:")
    print(f"    {primary_url}")
    if alternate != primary_url:
        print(f"    {alternate}")
    print("  Nothing is blocked in the meantime - the console below still files")
    print("  reports, and --doctor explains what went wrong.")
    _redraw_cli_prompt(session)


def _watch_web_clients(
    session: ReporterSession, state: WebSessionState, allow_idle_timeout: bool
):
    """End the session once the last browser page is gone.

    This is the browser-side half of "closing either surface closes the other".
    """
    while not session.is_shutting_down():
        reason = state.shutdown_reason(allow_idle_timeout)
        if reason:
            session.request_shutdown(reason)
            return
        time.sleep(0.5)


def _start_shutdown_watchdog(session: ReporterSession):
    """Take the process down when the *other* surface ends the session.

    The console is parked inside input() and cannot be woken from here, so - as
    before - the process is exited from under it once the workbook is quiet.
    That is what makes closing the browser close the terminal too.
    """

    def _wait_and_exit():
        session.wait_for_shutdown()
        # A beat for the page's next ping to collect the shutdown, so it can
        # show "session ended" rather than a failed connection.
        time.sleep(0.6)
        wait_for_quiet_workbook()
        _print_session_summary(session)
        sys.stdout.flush()
        sys.stderr.flush()
        os._exit(0)

    threading.Thread(
        target=_wait_and_exit, name="task-report-exit", daemon=True
    ).start()


def run_web(
    session: ReporterSession,
    start_cli: bool,
    port_hint: int = None,
    open_browser: bool = True,
):
    """Run the browser UI.  Returns an exit code, or None if it could not start."""
    server, state, error = start_web_server(session, port_hint)
    if server is None:
        print(f"  The browser UI could not start: {error}")
        return None

    port = server.server_address[1]
    url = web_url(port, state.token)

    launcher = open_in_browser(url) if open_browser else ""
    _print_web_banner(url, launcher, open_browser)

    if open_browser:
        # In a thread: a slow browser must not hold up the console prompt.
        threading.Thread(
            target=_ensure_page_loaded,
            args=(session, state, port, url),
            name="task-report-web-check",
            daemon=True,
        ).start()

    # Only the browser-only session may end itself on silence; with a console
    # present, a throttled background tab must not take the console down.
    threading.Thread(
        target=_watch_web_clients,
        args=(session, state, not start_cli),
        name="task-report-web-watch",
        daemon=True,
    ).start()

    if start_cli:
        _start_shutdown_watchdog(session)
        cli_console_loop(session, dual=True)
    else:
        while not session.is_shutting_down():
            session.wait_for_shutdown(0.3)

    session.request_shutdown("session ended")
    try:
        server.shutdown()
    except Exception:
        pass
    return 0


# ---------------------------------------------------------------------------
# Desktop app - the Windows window
# ---------------------------------------------------------------------------

# The window is the same page as the browser UI, drawn inside the process that
# serves it.  Nothing about the UI changes; what changes is that there is no
# terminal to keep open and no browser tab to keep track of.
#
# Windows draws it with WebView2 (the Edge runtime, present on every supported
# Windows install), so this is a native window with a native title bar and a
# taskbar entry - not a browser in disguise.  The whole WSLg failure class the
# Qt window suffered from is still avoided, because no X11, Wayland or
# compositor is involved.

APP_WINDOW_TITLE = "Task Reporter"
# In CSS pixels - what the page is laid out in.  See _app_window_geometry().
APP_WINDOW_SIZE = (1080, 780)
APP_WINDOW_MIN_SIZE = (760, 560)

# Two files the app keeps beside the workbook.  The lock is how a second
# launch finds the window that is already open; the log is where the output
# that used to go to the terminal goes instead.
APP_LOCK_PATH = os.path.join(BASE_DIR, ".task_reporter_app.lock")
APP_LOG_PATH = os.path.join(BASE_DIR, ".task_reporter_app.log")
APP_LOG_MAX_BYTES = 512 * 1024


def _app_window_geometry():
    """The window size to ask for, and the smallest it may be dragged to.

    pywebview's width and height are logical pixels - the same units the page
    is laid out in - so a scaled display needs no arithmetic here, and adding
    any produces a window proportionally too big.  (Measured: a window asked
    for at 1350 gives the page 1336 CSS pixels on a 125% display.  Screenshot
    tools that are not themselves DPI-aware report such a window at its
    logical size and crop the capture, which looks exactly like the layout
    overflowing.  It is not.)

    The only adjustment worth making is the one that is not about DPI: a
    default that does not fit the screen.  GetSystemMetrics answers in logical
    pixels for this process, which is the unit wanted.
    """
    width, height = APP_WINDOW_SIZE
    min_width, min_height = APP_WINDOW_MIN_SIZE
    if os.name != "nt":
        return width, height, min_width, min_height

    try:
        import ctypes

        user32 = ctypes.windll.user32
        screen_w = user32.GetSystemMetrics(0)  # SM_CXSCREEN
        screen_h = user32.GetSystemMetrics(1)  # SM_CYSCREEN
        if screen_w > 0 and screen_h > 0:
            width = min(width, int(screen_w * 0.94))
            # Short of the full height, so the title bar does not start out
            # underneath the taskbar.
            height = min(height, int(screen_h * 0.88))
            # The minimum is a floor, so it must never end up above the size.
            min_width = min(min_width, width)
            min_height = min(min_height, height)
    except Exception:
        pass

    return width, height, min_width, min_height


def start_app_log():
    """Send stdout/stderr to a file, because a windowed app has neither.

    PyInstaller's windowed build leaves sys.stdout as None, and this program
    prints freely - so without this, the first print() anywhere would raise.
    """
    try:
        if os.path.exists(APP_LOG_PATH) and os.path.getsize(APP_LOG_PATH) > APP_LOG_MAX_BYTES:
            os.replace(APP_LOG_PATH, APP_LOG_PATH + ".1")
    except OSError:
        pass
    try:
        handle = open(APP_LOG_PATH, "a", encoding="utf-8", errors="replace", buffering=1)
    except OSError:
        handle = open(os.devnull, "w", encoding="utf-8")
    sys.stdout = handle
    sys.stderr = handle
    # stdin is never read in app mode, but leaving it as None makes any stray
    # input() raise something unhelpful instead of ending cleanly.
    if sys.stdin is None:
        try:
            sys.stdin = open(os.devnull, "r", encoding="utf-8")
        except OSError:
            pass
    print(f"\n=== {datetime.now().strftime(TIMESTAMP_FORMAT)}  app start ===")
    print(f"  data folder: {BASE_DIR}")


def show_app_error(message: str, fatal: bool = True):
    """Put a problem in front of someone running the windowed app.

    Without this the exe would simply vanish, or do something unasked for:
    there is no console for the explanation to land in, and the log file is
    only useful once you know to look for it.
    """
    print(message)
    if os.name != "nt" or not FROZEN:
        return
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(
            None,
            f"{message}\n\nThe full details are in:\n{APP_LOG_PATH}",
            f"{APP_WINDOW_TITLE} - could not start"
            if fatal
            else f"{APP_WINDOW_TITLE}",
            (0x10 if fatal else 0x30) | 0x1000,  # ICONERROR/ICONWARNING, SYSTEMMODAL
        )
    except Exception:
        pass


def _write_app_lock(port: int, token: str):
    try:
        with open(APP_LOCK_PATH, "w", encoding="utf-8") as handle:
            json.dump({"port": port, "token": token, "pid": os.getpid()}, handle)
    except OSError:
        pass


def _clear_app_lock():
    try:
        os.remove(APP_LOCK_PATH)
    except OSError:
        pass


def focus_running_app() -> bool:
    """Hand a second launch over to the window that is already open.

    Returns True when a running app answered, in which case this process has
    nothing left to do.  A stale lock file simply fails to answer and is
    removed, so a crash never blocks the next launch.
    """
    try:
        with open(APP_LOCK_PATH, "r", encoding="utf-8") as handle:
            lock = json.load(handle)
        port = int(lock["port"])
        token = str(lock["token"])
    except (OSError, ValueError, KeyError, TypeError):
        return False

    url = f"http://{WEB_HOST}:{port}/api/focus?t={urllib.parse.quote(token)}"
    request = urllib.request.Request(url, data=b"", method="POST")
    try:
        # Generous for a loopback call to a process that is already running:
        # a timeout here would clear a live lock and open a second window.
        with urllib.request.urlopen(request, timeout=5) as response:
            if response.status == 200:
                return True
    except Exception:
        pass
    _clear_app_lock()
    return False


def _watch_app_focus(session: ReporterSession, state: WebSessionState, app: "DesktopApp"):
    """Bring the window forward when another launch asks for it."""
    while not session.is_shutting_down():
        if state.take_focus_request():
            app.show_main()
        time.sleep(0.3)


def _watch_app_shutdown(session: ReporterSession, app: "DesktopApp"):
    """Close every window when the session ends from somewhere other than one."""
    session.wait_for_shutdown()
    app.shutdown()


# ---------------------------------------------------------------------------
# The frameless window: title bar buttons, drag and resize
# ---------------------------------------------------------------------------
#
# The window has no Windows frame - the page draws the 40 px title bar from
# the design.  Moving and resizing are still done by Windows: a mousedown on
# the title bar or on one of the 6 px edge grips is handed to the window as
# if it had landed on a real caption or border (WM_NCLBUTTONDOWN), so the
# move and size loops, Aero Snap and the minimum size are all Windows' own.
# Where that is not available (not Windows), pywebview's drag region moves it.

_WM_NCLBUTTONDOWN = 0x00A1
_HIT_TESTS = {
    "caption": 2, "w": 10, "e": 11, "n": 12, "nw": 13, "ne": 14, "s": 15, "sw": 16, "se": 17,
}


def _user32():
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    return user32


def _on_ui_thread(window, action) -> bool:
    """Run `action` on the window's own thread, without waiting for it.

    Every change to a window goes through here.  Touching a form from another
    thread - even pywebview's own `window.on_top = True` - makes the .NET call
    wait for the window's thread while this thread still holds Python's GIL,
    and the window's thread needs the GIL to run its event handlers: the whole
    app freezes.  BeginInvoke never waits, so it cannot.
    """
    form = getattr(window, "native", None) if window is not None else None
    if form is None:
        return False
    try:
        from System import Action  # pythonnet - already loaded by pywebview

        def guarded():
            try:
                action()
            except Exception as exc:
                print(f"  [!] window action failed: {exc}")

        form.BeginInvoke(Action(guarded))
        return True
    except Exception as exc:
        print(f"  [!] could not reach the window: {exc}")
        return False


def _enable_native_caption(window):
    """Let the page's `app-region: drag` areas act as the window's caption.

    WebView2 (runtime 123+, SDK 1.0.2420+) can report those areas to Windows
    as title bar, which makes dragging, snapping and double-click-to-maximize
    Windows' own, with no round trip through Python.  Runs on the window's
    thread, before the page has finished loading.
    """
    form = getattr(window, "native", None)
    control = getattr(getattr(form, "browser", None), "webview", None)
    if control is None:
        return

    def apply(*_):
        try:
            control.CoreWebView2.Settings.IsNonClientRegionSupportEnabled = True
        except Exception as exc:
            # An older runtime: the page falls back to its own drag handling.
            print(f"  native title bar not available ({exc})")

    if control.CoreWebView2 is not None:
        apply()
    else:
        control.CoreWebView2InitializationCompleted += apply


def _fit_maximized_bounds(form):
    """A frameless form maximizes over the taskbar unless told the work area.

    Kept up to date as the window moves, because Windows can maximize it by
    itself too (double-clicking the title bar, Win+Up, dragging to the top).
    """
    from System.Drawing import Rectangle
    from System.Windows.Forms import Screen

    screen = Screen.FromHandle(form.Handle)
    work, bounds = screen.WorkingArea, screen.Bounds
    form.MaximizedBounds = Rectangle(
        work.X - bounds.X, work.Y - bounds.Y, work.Width, work.Height
    )


def _focus_page(form):
    """Give the keyboard to the page inside the form, not just the form.

    pywebview does this once, the first time a window is shown; a window that
    is hidden and shown again (the quick-add one, every time) needs it again.
    """
    browser = getattr(form, "browser", None)
    control = getattr(browser, "webview", None) if browser is not None else None
    if control is not None:
        control.Focus()


class _WindowApi:
    """What the page can ask its own window to do (window.pywebview.api).

    pywebview publishes every public method; anything starting with an
    underscore stays private, which is why the references below do.
    """

    def __init__(self, app: "DesktopApp", which: str):
        self._app = app
        self._which = which

    def _window(self):
        return self._app.main if self._which == "main" else self._app.quick

    def native_frame(self) -> bool:
        window = self._window()
        return os.name == "nt" and window is not None and getattr(window, "native", None) is not None

    def start_drag(self):
        self._native_press("caption")

    def start_resize(self, edge):
        if edge in _HIT_TESTS and edge != "caption" and not self._app.is_maximized(self._which):
            self._native_press(edge)

    def _native_press(self, where: str):
        window = self._window()
        if not self.native_frame():
            return
        form = window.native

        def press():
            hwnd = int(form.Handle.ToInt64())
            user32 = _user32()
            # The button may already be up by the time this runs (a quick
            # click).  Starting the size loop then would leave the window
            # glued to the pointer until the next click, so only go on while
            # it is still held.
            if not user32.GetAsyncKeyState(0x01) & 0x8000:
                return
            user32.ReleaseCapture()
            # Windows' own move / size loop runs from here until the button
            # comes up - snapping, Win+arrows and the minimum size included.
            user32.SendMessageW(hwnd, _WM_NCLBUTTONDOWN, _HIT_TESTS[where], 0)

        _on_ui_thread(window, press)

    def minimize(self):
        window = self._window()
        form = getattr(window, "native", None)
        if form is None:
            window.minimize()
            return

        def iconify():
            from System.Windows.Forms import FormWindowState

            form.WindowState = FormWindowState.Minimized

        _on_ui_thread(window, iconify)

    def is_maximized(self) -> bool:
        return self._app.is_maximized(self._which)

    def toggle_maximize(self) -> bool:
        return self._app.toggle_maximize()

    def close(self, force=False):
        """The title bar's close button (the page has already asked about edits)."""
        if self._which == "quick":
            self._app.hide_quick()
        else:
            self._app.close_main(force=bool(force))

    def pick_folder(self):
        window = self._window()
        try:
            import webview

            kind = getattr(getattr(webview, "FileDialog", None), "FOLDER", None)
            if kind is None:
                kind = webview.FOLDER_DIALOG
            chosen = window.create_file_dialog(kind, directory=BASE_DIR)
        except Exception as exc:
            print(f"  [!] folder picker failed: {exc}")
            return ""
        if not chosen:
            return ""
        return chosen[0] if isinstance(chosen, (list, tuple)) else str(chosen)

    def set_theme(self, theme):
        self._app.set_background("dark" if theme == "dark" else "light")

    def quick_hide(self):
        self._app.hide_quick()

    def open_main(self):
        self._app.hide_quick()
        self._app.show_main()


APP_BACKGROUND_LIGHT = "#f3f2f2"
APP_BACKGROUND_DARK = "#201e1d"
QUICK_WINDOW_SIZE = (420, 360)


def _windows_prefers_dark() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return value == 0
    except OSError:
        return False


def app_background() -> str:
    """The colour the window is painted before the page loads.

    It matches the theme the page is about to draw, so a dark window does not
    flash light on start (and the other way round).
    """
    theme = load_settings()["theme"]
    dark = theme == "dark" or (theme == "system" and _windows_prefers_dark())
    return APP_BACKGROUND_DARK if dark else APP_BACKGROUND_LIGHT


# ---------------------------------------------------------------------------
# The global quick-add shortcut
# ---------------------------------------------------------------------------

_HOTKEY_MODIFIERS = {"ctrl": 0x0002, "control": 0x0002, "alt": 0x0001, "shift": 0x0004, "win": 0x0008}
_HOTKEY_NAMED_KEYS = {
    "space": 0x20, "enter": 0x0D, "tab": 0x09, "insert": 0x2D, "delete": 0x2E, "home": 0x24,
    "end": 0x23, "pageup": 0x21, "pagedown": 0x22, "up": 0x26, "down": 0x28, "left": 0x25,
    "right": 0x27, "`": 0xC0, "-": 0xBD, "=": 0xBB, "[": 0xDB, "]": 0xDD, ";": 0xBA,
    "'": 0xDE, ",": 0xBC, ".": 0xBE, "/": 0xBF, "\\": 0xDC,
}


def parse_hotkey(text: str):
    """"Ctrl+Alt+T" -> (modifiers, virtual key), or None if it is not one."""
    parts = [part.strip().lower() for part in str(text or "").split("+") if part.strip()]
    modifiers = 0
    key = None
    for part in parts:
        if part in _HOTKEY_MODIFIERS:
            modifiers |= _HOTKEY_MODIFIERS[part]
        elif key is None:
            key = part
        else:
            return None
    if not modifiers or key is None:
        return None
    if len(key) == 1 and key.isalnum():
        return modifiers, ord(key.upper())
    match = re.fullmatch(r"f([1-9]|1[0-9]|2[0-4])", key)
    if match:
        return modifiers, 0x70 + int(match.group(1)) - 1
    if key in _HOTKEY_NAMED_KEYS:
        return modifiers, _HOTKEY_NAMED_KEYS[key]
    return None


class GlobalHotkey:
    """One system-wide shortcut, registered with RegisterHotKey.

    The registration belongs to the thread that made it, so a thread of its
    own owns it and runs the message loop the WM_HOTKEY arrives on.
    Re-binding is a message to that thread rather than a call from outside.
    """

    _WM_HOTKEY = 0x0312
    _WM_QUIT = 0x0012
    _WM_REBIND = 0x8001  # WM_APP + 1
    _MOD_NOREPEAT = 0x4000
    _ID = 0x5452  # "TR"

    def __init__(self, callback):
        self._callback = callback
        self._wanted = ""
        self._thread_id = None
        self._ready = threading.Event()

    def start(self, hotkey: str):
        if os.name != "nt":
            return
        self._wanted = hotkey
        threading.Thread(target=self._run, name="task-report-hotkey", daemon=True).start()
        self._ready.wait(3)

    def rebind(self, hotkey: str):
        self._wanted = hotkey
        if self._thread_id:
            import ctypes

            ctypes.windll.user32.PostThreadMessageW(self._thread_id, self._WM_REBIND, 0, 0)

    def stop(self):
        if self._thread_id:
            import ctypes

            ctypes.windll.user32.PostThreadMessageW(self._thread_id, self._WM_QUIT, 0, 0)

    def _register(self, user32) -> None:
        user32.UnregisterHotKey(None, self._ID)
        TRAY_STATUS["hotkey"] = ""
        TRAY_STATUS["hotkeyError"] = ""
        if not self._wanted:
            return
        parsed = parse_hotkey(self._wanted)
        if parsed is None:
            TRAY_STATUS["hotkeyError"] = f"“{self._wanted}” is not a shortcut Windows can register."
            return
        modifiers, key = parsed
        if user32.RegisterHotKey(None, self._ID, modifiers | self._MOD_NOREPEAT, key):
            TRAY_STATUS["hotkey"] = self._wanted
            print(f"  quick-add shortcut: {self._wanted}")
        else:
            TRAY_STATUS["hotkeyError"] = (
                f"{self._wanted} is already taken by another program - pick another shortcut."
            )
            print(f"  [!] {TRAY_STATUS['hotkeyError']}")

    def _run(self):
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        message = wintypes.MSG()
        # Makes the thread's message queue exist before anyone posts to it.
        user32.PeekMessageW(ctypes.byref(message), None, 0, 0, 0)
        self._thread_id = kernel32.GetCurrentThreadId()
        self._register(user32)
        self._ready.set()
        while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
            if message.message == self._WM_HOTKEY and message.wParam == self._ID:
                try:
                    self._callback()
                except Exception as exc:
                    print(f"  [!] quick-add shortcut failed: {exc}")
            elif message.message == self._WM_REBIND:
                self._register(user32)
        user32.UnregisterHotKey(None, self._ID)


# ---------------------------------------------------------------------------
# The desktop app: main window, quick-add window, tray icon
# ---------------------------------------------------------------------------


class DesktopApp:
    """Owns the two windows, the tray icon and the shortcut, and how they close.

    Closing has three outcomes, decided in one place (_on_main_closing):
    the page has an unsaved edit -> ask first; "Keep in tray" is on -> hide;
    otherwise -> close, which ends the session.
    """

    def __init__(self, session: ReporterSession, state: WebSessionState, url: str):
        self.session = session
        self.state = state
        self.url = url
        self.main = None
        self.quick = None
        self.tray = None
        self.hotkey = None
        self._maximized = {"main": False, "quick": False}
        self._close_approved = False
        self._quitting = False
        # Tray -> Quit with an edit open: once the page has dealt with the
        # edit, its close() must quit, not drop back into the tray.
        self._quit_pending = False

    # -------------------------------------------------------------- windows

    def create_windows(self, webview, width, height, min_width, min_height):
        frameless = os.name == "nt"
        background = app_background()
        self.main = webview.create_window(
            APP_WINDOW_TITLE,
            self.url + "&app=1",
            width=width,
            height=height,
            min_size=(min_width, min_height),
            background_color=background,
            text_select=True,
            frameless=frameless,
            easy_drag=False,
            js_api=_WindowApi(self, "main"),
        )
        if frameless:
            self.main.events.before_show += lambda: _enable_native_caption(self.main)
            self.main.events.shown += self._keep_maximized_bounds
            self.main.events.moved += self._keep_maximized_bounds
        self.main.events.closing += self._on_main_closing
        self.main.events.closed += self._on_main_closed
        self.main.events.maximized += lambda: self._maximized.__setitem__("main", True)
        self.main.events.restored += lambda: self._maximized.__setitem__("main", False)
        self.main.events.minimized += lambda: self._maximized.__setitem__("main", False)

        if os.name == "nt":
            quick_width, quick_height = QUICK_WINDOW_SIZE
            self.quick = webview.create_window(
                "Quick add task",
                self.url + "&app=1&quick=1",
                width=quick_width,
                height=quick_height,
                background_color=background,
                frameless=True,
                easy_drag=False,
                on_top=True,
                hidden=True,
                resizable=False,
                js_api=_WindowApi(self, "quick"),
            )
            self.quick.events.before_show += lambda: _enable_native_caption(self.quick)
            # The quick window is only ever hidden; closing it must not end
            # the app (Alt+F4 on it, for instance).
            self.quick.events.closing += self._on_quick_closing

    def _keep_maximized_bounds(self, *_):
        form = getattr(self.main, "native", None)
        if form is not None:
            _on_ui_thread(self.main, lambda: _fit_maximized_bounds(form))

    def is_maximized(self, which: str) -> bool:
        return self._maximized.get(which, False)

    def toggle_maximize(self) -> bool:
        form = getattr(self.main, "native", None)
        maximize = not self._maximized["main"]
        self._maximized["main"] = maximize
        if form is None:
            (self.main.maximize if maximize else self.main.restore)()
            return maximize

        def flip():
            from System.Windows.Forms import FormWindowState

            if maximize:
                _fit_maximized_bounds(form)
                form.WindowState = FormWindowState.Maximized
            else:
                form.WindowState = FormWindowState.Normal

        _on_ui_thread(self.main, flip)
        return maximize

    def set_background(self, theme: str):
        colour = APP_BACKGROUND_DARK if theme == "dark" else APP_BACKGROUND_LIGHT
        for window in (self.main, self.quick):
            form = getattr(window, "native", None) if window else None
            if form is None:
                continue

            def paint(f=form):
                from System.Drawing import ColorTranslator

                f.BackColor = ColorTranslator.FromHtml(colour)

            _on_ui_thread(window, paint)

    def show_main(self, command: str = None):
        window = self.main
        if window is None:
            return
        form = getattr(window, "native", None)

        def bring_forward():
            from System.Windows.Forms import FormWindowState

            form.Show()
            if form.WindowState == FormWindowState.Minimized:
                form.WindowState = (
                    FormWindowState.Maximized if self._maximized["main"] else FormWindowState.Normal
                )
            # A momentary topmost is what actually lifts a WebView2 window
            # above the window that launched it.
            form.TopMost = True
            form.TopMost = False
            form.Activate()
            _focus_page(form)

        if form is not None:
            _on_ui_thread(window, bring_forward)
        else:
            try:
                window.show()
            except Exception:
                pass
        if command:
            self._run_js(window, f"window.TR && TR.{command}")

    def show_quick(self):
        window = self.quick
        if window is None:
            self.show_main()
            return
        form = getattr(window, "native", None)
        if form is None:
            return

        def appear():
            from System.Drawing import Size
            from System.Windows.Forms import Screen

            # pywebview takes a frame's worth off a frameless window, so the
            # size is set here, exactly, in the form's own (physical) pixels.
            scale = getattr(form, "_scale", 1.0) or 1.0
            width, height = QUICK_WINDOW_SIZE
            form.ClientSize = Size(int(width * scale), int(height * scale))
            # Bottom right, just above the taskbar, like a flyout.
            work = Screen.PrimaryScreen.WorkingArea
            margin = int(12 * scale)
            form.Left = work.X + work.Width - form.Width - margin
            form.Top = work.Y + work.Height - form.Height - margin
            form.Show()
            form.TopMost = True
            # Being shown is not the same as having the keyboard.
            form.Activate()
            _focus_page(form)

        _on_ui_thread(window, appear)
        self._run_js(window, "window.TR && TR.quickShown()")

    def hide_quick(self):
        form = getattr(self.quick, "native", None) if self.quick else None
        if form is not None:
            _on_ui_thread(self.quick, form.Hide)

    def hide_main(self):
        form = getattr(self.main, "native", None) if self.main else None
        if form is not None:
            _on_ui_thread(self.main, form.Hide)

    def _run_js(self, window, script: str):
        def run():
            try:
                window.evaluate_js(script)
            except Exception as exc:
                print(f"  [!] page call failed: {exc}")

        threading.Thread(target=run, daemon=True).start()

    # -------------------------------------------------------------- closing

    def close_main(self, force: bool = False):
        """The page's close button: it has already dealt with unsaved edits."""
        if self._quit_pending:
            self.shutdown()
            return
        if load_settings()["keepInTray"] and self.tray is not None:
            print("  close button -> hidden in the tray")
            self.hide_main()
            return
        self._close_approved = True
        self.main.destroy()

    def _on_main_closing(self):
        """Alt+F4, the taskbar's Close, or our own destroy().  False cancels."""
        if self._quitting or self._close_approved:
            return True
        if self.state.page_dirty:
            # The page asks "Save changes?" and calls close() itself after.
            print("  close held: the page has an unsaved edit - asking first")
            self._run_js(self.main, "window.TR && TR.requestWindowClose()")
            return False
        if load_settings()["keepInTray"] and self.tray is not None:
            print("  close -> hidden in the tray")
            self.hide_main()
            return False
        return True

    def _on_quick_closing(self):
        if self._quitting:
            return True
        self.hide_quick()
        return False

    def _on_main_closed(self):
        self.session.request_shutdown("window closed")

    def quit(self):
        """Tray -> Quit.  Asks about unsaved edits first, like the close button."""
        if self.state.page_dirty:
            self._quit_pending = True
            self.show_main()
            self._run_js(self.main, "window.TR && TR.requestWindowClose()")
            return
        self.shutdown()

    def shutdown(self):
        if self._quitting:
            return
        self._quitting = True
        if self.hotkey is not None:
            self.hotkey.stop()
        if self.tray is not None:
            try:
                self.tray.stop()
            except Exception:
                pass
        for window in (self.quick, self.main):
            if window is None:
                continue
            try:
                window.destroy()
            except Exception:
                pass
        self.session.request_shutdown("window closed")

    # ----------------------------------------------------- tray and hotkey

    def start_tray(self):
        """The notification-area icon.  Optional: without pystray, no tray."""
        if os.name != "nt":
            return
        try:
            import io

            import pystray
            from PIL import Image
        except ImportError as exc:
            print(f"  no tray icon ({exc}) - pip install pystray pillow")
            return

        icon_bytes = _app_icon_bytes()
        try:
            image = Image.open(io.BytesIO(icon_bytes)) if icon_bytes else Image.new("RGB", (32, 32), "#ec3013")
        except Exception:
            image = Image.new("RGB", (32, 32), "#ec3013")

        menu = pystray.Menu(
            pystray.MenuItem("Open Task Reporter", lambda icon, item: self.show_main()),
            # The default item is what a left-click does, and is drawn bold.
            pystray.MenuItem("Quick add task…", lambda icon, item: self.show_quick(), default=True),
            pystray.MenuItem("Previous reports", lambda icon, item: self.show_main("openHistory()")),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", lambda icon, item: self.quit()),
        )
        self.tray = pystray.Icon("TaskReporter", image, APP_WINDOW_TITLE, menu)
        threading.Thread(target=self.tray.run, name="task-report-tray", daemon=True).start()
        TRAY_STATUS["available"] = True
        TRAY_STATUS["running"] = True
        print("  tray icon running")

    def start_hotkey(self):
        if os.name != "nt":
            return
        self.hotkey = GlobalHotkey(self.show_quick)
        self.hotkey.start(load_settings()["hotkey"])
        SETTINGS_LISTENERS.append(self._settings_changed)

    def _settings_changed(self, changes: dict):
        if "hotkey" in changes and self.hotkey is not None:
            self.hotkey.rebind(changes["hotkey"])
            # Give the hotkey thread a moment, so Settings shows the outcome.
            time.sleep(0.2)


def run_desktop_app(session: ReporterSession, port_hint: int = None):
    """Run the UI in a native window.  Returns an exit code, or None if the
    window could not be created at all."""
    try:
        import webview
    except ImportError as exc:
        # ModuleNotFoundError when pywebview is absent, plain ImportError when
        # it is present but one of its DLLs will not load.
        print(f"  The desktop window needs pywebview ({exc}).")
        return None

    server, state, error = start_web_server(session, port_hint)
    if server is None:
        print(f"  The desktop window could not start: {error}")
        return None

    port = server.server_address[1]
    # Same process, same machine: the window always talks to 127.0.0.1, never
    # to the `localhost` spelling the WSL browser path needs.
    url = web_url(port, state.token, host=WEB_HOST)
    _write_app_lock(port, state.token)
    print(f"  serving the window on {url}")

    width, height, min_width, min_height = _app_window_geometry()
    print(f"  window {width}x{height} (minimum {min_width}x{min_height})")

    app = DesktopApp(session, state, url)
    try:
        app.create_windows(webview, width, height, min_width, min_height)

        threading.Thread(
            target=_watch_app_focus,
            args=(session, state, app),
            name="task-report-app-focus",
            daemon=True,
        ).start()
        threading.Thread(
            target=_watch_app_shutdown,
            args=(session, app),
            name="task-report-app-close",
            daemon=True,
        ).start()

        def started():
            app.start_tray()
            app.start_hotkey()

        storage = app_storage_path()
        os.makedirs(storage, exist_ok=True)
        # start() owns the main thread until the last window is closed.
        webview.start(started, private_mode=False, storage_path=storage)
    except Exception as exc:
        print(f"  The desktop window failed to open ({exc}).")
        print("  Falling back to the browser for this session.")
        launcher = open_in_browser(url)
        if not launcher:
            _clear_app_lock()
            try:
                server.shutdown()
            except Exception:
                pass
            return None
        _watch_web_clients(session, state, allow_idle_timeout=True)

    app.shutdown()
    session.request_shutdown("window closed")
    _clear_app_lock()
    try:
        server.shutdown()
    except Exception:
        pass
    return 0


def _http_probe(url: str):
    """Fetch `url` from this shell.  Returns (ok, detail)."""
    try:
        with urllib.request.urlopen(url, timeout=5) as response:
            return response.status == 200, f"HTTP {response.status}"
    except urllib.error.HTTPError as exc:
        return False, f"HTTP {exc.code}"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"


def _windows_http_probe(url: str):
    """Fetch `url` from the Windows side, which is where the browser lives."""
    if not shutil.which("powershell.exe"):
        return False, "powershell.exe is not on PATH (no Windows interop)"
    command = (
        "try { "
        f"$r = Invoke-WebRequest -UseBasicParsing -Uri '{url}' -TimeoutSec 6; "
        "'STATUS ' + [int]$r.StatusCode "
        "} catch { 'ERROR ' + $_.Exception.Message }"
    )
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True,
            text=True,
            timeout=25,
            check=False,
        )
    except Exception as exc:
        return False, f"could not run powershell.exe: {exc}"
    output = " ".join((result.stdout or "").split()).strip()
    if output.startswith("STATUS 200"):
        return True, "HTTP 200"
    return False, output or "no answer from powershell.exe"


# Both the console's normal return and the shutdown watchdog reach for the
# summary, and whichever loses the race must not print it twice.
_SUMMARY_LOCK = threading.Lock()
_SUMMARY_PRINTED = False


def _print_session_summary(session: ReporterSession):
    global _SUMMARY_PRINTED
    with _SUMMARY_LOCK:
        if _SUMMARY_PRINTED:
            return
        _SUMMARY_PRINTED = True

    print()
    print(f"Task Reporter closed - {session.reason or 'session ended'}.")
    if session.saved_count:
        plural = "s" if session.saved_count != 1 else ""
        print(f"{session.saved_count} report{plural} filed this session.")
    queued = pending_report_count()
    if queued:
        print(f"{queued} report(s) still queued for the workbook.")
    ready = board_counts()["ready"]
    if ready:
        print(f"{ready} ticked task(s) still waiting on the board.")


def _doctor_web_section() -> bool:
    """Check the browser UI end to end.  Returns True when it can be used."""
    print("  Browser UI - the default surface")
    print("  " + "-" * 64)

    server, state, error = start_web_server(ReporterSession())
    if server is None:
        print(f"  [x ] no loopback port could be bound -> {error}")
        print("       This is the only way the browser UI can fail outright.")
        return False

    port = server.server_address[1]
    try:
        print(f"  [ok] serving on {WEB_HOST}:{port}")

        ok_here, detail = _http_probe(
            f"http://{WEB_HOST}:{port}/api/health?t={state.token}"
        )
        print(f"  [{'ok' if ok_here else 'x '}] reachable from this shell -> {detail}")

        reachable = ok_here
        if is_wsl():
            # This is the check that matters: the page is fetched and drawn by
            # a Windows browser, so Windows is what has to reach the port.
            windows_ok = False
            for host in ("localhost", WEB_HOST):
                ok, detail = _windows_http_probe(
                    f"http://{host}:{port}/api/health?t={state.token}"
                )
                print(
                    f"  [{'ok' if ok else 'x '}] reachable from Windows "
                    f"via {host} -> {detail}"
                )
                windows_ok = windows_ok or ok
            reachable = reachable and windows_ok
        else:
            print("  [--] not WSL - the local browser is used directly")

        openers = [label for label, _argv in _browser_launchers("http://127.0.0.1/")]
        print(f"  browser openers    : {', '.join(openers) if openers else '(none)'}")
        if not openers:
            print("       No opener found. The URL is printed at start-up so it")
            print("       can be pasted into a browser by hand.")
    finally:
        server.shutdown()
        server.server_close()

    print()
    if reachable:
        print("  [ok] The browser UI works. This is what runs by default.")
    else:
        print("  [x ] The browser could not reach the server.")
        print("       Set TASK_REPORT_BROWSER to a specific opener, or paste the")
        print("       URL printed at start-up into a browser yourself.")
    return reachable


def _doctor_qt_section(forced_platform: str = None) -> bool:
    print()
    print("  Qt window - opt-in fallback (--qt)")
    print("  " + "-" * 64)

    platform_name, attempts = probe_qt_platform(forced_platform)
    for name, ok, detail in attempts:
        print(f"  [{'ok' if ok else 'x '}] platform '{name}' -> {detail}")

    print()
    if platform_name:
        print(f"  [ok] A window renders under '{platform_name}'.")
        print("       Note that this probe cannot promise the next window will")
        print("       render too - WSLg can degrade between one and the next,")
        print("       which is exactly why the browser UI is now the default.")
        return True

    print("  [x ] No Qt platform could put a window on screen.")
    print("       Nothing is lost - the browser UI does not use Qt at all.")
    print("       If you specifically want the Qt window back, a degraded WSLg")
    print("       session is usually cleared from Windows with:")
    print("         wsl.exe --shutdown        (then reopen)")
    print("       or force a platform:  ./task-report --qt --platform xcb")
    return False


def _webview2_runtime_version() -> str:
    """The installed Edge WebView2 runtime, or "" when it cannot be found.

    Read from the registry, because that is where the evergreen runtime records
    itself whether it was installed per-machine or per-user.
    """
    if os.name != "nt":
        return ""
    try:
        import winreg
    except ModuleNotFoundError:
        return ""

    key_path = (
        r"SOFTWARE\Microsoft\EdgeUpdate\Clients"
        r"\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    )
    for root, flags in (
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_CURRENT_USER, 0),
    ):
        try:
            with winreg.OpenKey(
                root, key_path, 0, winreg.KEY_READ | flags
            ) as key:
                version, _ = winreg.QueryValueEx(key, "pv")
            if version and version != "0.0.0.0":
                return str(version)
        except OSError:
            continue
    return ""


def _doctor_app_section() -> bool:
    """Can the native desktop window (--app) come up on this machine?"""
    print()
    print("Desktop window (--app)")
    print("-" * 68)

    try:
        import webview  # noqa: F401
    except ModuleNotFoundError as exc:
        print(f"  [no] pywebview is not installed here ({exc.name}).")
        print("       The packaged Windows app ships it; a plain checkout")
        print("       needs: pip install pywebview")
        return False
    print("  [ok] pywebview importable")

    if os.name != "nt":
        print("  [--] not on Windows: the window would use this desktop's")
        print("       own toolkit, which is the thing the browser UI avoids.")
        print("       Use the browser UI here.")
        return False

    version = _webview2_runtime_version()
    if version:
        print(f"  [ok] Edge WebView2 runtime {version}")
    else:
        print("  [no] no Edge WebView2 runtime found in the registry.")
        print("       Install it from https://go.microsoft.com/fwlink/p/?LinkId=2124703")
        return False

    print(f"  [--] running as a packaged app : {FROZEN}")
    if os.path.exists(APP_LOCK_PATH):
        print(f"  [--] an app window looks open  : {APP_LOCK_PATH}")
    return True


def run_doctor(forced_platform: str = None) -> int:
    print("Task Reporter - environment check")
    print("=" * 68)
    print(f"  interpreter        : {sys.executable}")
    print(f"  script folder      : {BASE_DIR}")
    print(f"  workbook           : {EXCEL_FILE_PATH}")
    print(f"  workbook exists    : {os.path.exists(EXCEL_FILE_PATH)}")
    print(f"  queued reports     : {pending_report_count()}")
    print(f"  open in Excel now  : {workbook_is_open_in_excel()}")
    print(f"  task board         : {TASKS_FILE_PATH}")
    _counts = board_counts()
    print(
        f"  board tasks        : {_counts['total']} "
        f"({_counts['open']} open, {_counts['ready']} ready to file, "
        f"{_counts['filed']} filed)"
    )
    print(f"  running under WSL  : {is_wsl()}")
    print(f"  windows interop    : {bool(shutil.which('powershell.exe'))}")
    print(f"  PySide6 importable : {GUI_AVAILABLE}")
    print(f"  DISPLAY            : {os.environ.get('DISPLAY') or '(unset)'}")
    print(f"  WAYLAND_DISPLAY    : {os.environ.get('WAYLAND_DISPLAY') or '(unset)'}")
    print(f"  stdin is a tty     : {bool(sys.stdin) and sys.stdin.isatty()}")
    print("=" * 68)

    web_ok = _doctor_web_section()
    app_ok = _doctor_app_section()
    _doctor_qt_section(forced_platform)

    print()
    print("-" * 68)
    if web_ok and app_ok:
        print("  Verdict: the desktop window will come up - run TaskReporter.exe")
        print("           (or ./task-report --app), and the browser UI also works.")
    elif web_ok:
        print("  Verdict: run ./task-report - the browser UI will come up.")
    else:
        print("  Verdict: the terminal console still works everywhere:")
        print("    ./task-report --cli")
    return 0


def main(argv: list) -> int:
    # A windowed .exe has no stdout at all, so this has to come before the
    # first print() anywhere - including argparse's own error messages.
    if FROZEN and (sys.stdout is None or sys.stderr is None):
        start_app_log()

    # Without a tty, Python block-buffers stdout - which would hide the one
    # line that matters most when a browser fails to open: the UI's address.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(line_buffering=True)
        except Exception:
            pass

    parser = argparse.ArgumentParser(
        prog="task-report-maker.py",
        description=(
            "File task reports into task_reports.xlsx. By default the browser "
            "UI and the terminal console run together; closing either one ends "
            "the session."
        ),
    )
    parser.add_argument(
        "--gui", action="store_true", help="UI only (no terminal console)"
    )
    parser.add_argument(
        "--cli", action="store_true", help="Terminal console only (no UI)"
    )
    parser.add_argument(
        "--app",
        action="store_true",
        help=(
            "Show the UI in a native desktop window instead of a browser tab. "
            "This is what the packaged Windows app does by default."
        ),
    )
    parser.add_argument(
        "--qt",
        action="store_true",
        help=(
            "Use the old PySide6 window instead of the browser UI. Depends on "
            "a working display; the browser UI does not."
        ),
    )
    parser.add_argument(
        "--port",
        type=int,
        metavar="N",
        help=(
            f"Serve the browser UI on port N instead of the first free port "
            f"from {WEB_PREFERRED_PORTS[0]}-{WEB_PREFERRED_PORTS[-1]}."
        ),
    )
    parser.add_argument(
        "--no-browser",
        dest="no_browser",
        action="store_true",
        help="Start the UI server but print the address instead of opening it",
    )
    parser.add_argument(
        "-m",
        "--report",
        metavar="TEXT",
        help="File TEXT as a report and exit immediately",
    )
    parser.add_argument(
        "--list",
        dest="list_reports",
        action="store_true",
        help="Print recent reports and exit",
    )
    parser.add_argument(
        "--board",
        action="store_true",
        help="Print the task board and exit",
    )
    parser.add_argument(
        "-t",
        "--task",
        metavar="TEXT",
        help=(
            "Add TEXT to the task board and exit. A leading date and a "
            "[project] prefix are both optional, e.g. "
            "-t '21.08.2026 [onex-academy] fix the links'"
        ),
    )
    parser.add_argument(
        "--merge-days",
        dest="merge_days",
        action="store_true",
        help=(
            "Fold days that already have several rows in the workbook into one "
            "row each and exit. Only needed once, on a workbook written before "
            "a day became a single cell; use --dry-run to see it first"
        ),
    )
    parser.add_argument(
        "--dry-run",
        dest="dry_run",
        action="store_true",
        help="With --merge-days, report what would change and write nothing",
    )
    parser.add_argument(
        "--file-checked",
        dest="file_checked",
        action="store_true",
        help=(
            "Write every ticked task on the board into the workbook (one row "
            "per day, grouped by project) and exit"
        ),
    )
    parser.add_argument(
        "--platform",
        metavar="NAME",
        help=(
            "With --qt, force a Qt platform plugin (e.g. wayland, xcb) instead "
            "of auto-detecting. Also settable as TASK_REPORT_QT_PLATFORM."
        ),
    )
    parser.add_argument(
        "--doctor",
        action="store_true",
        help="Check both UIs end to end and explain anything that cannot start",
    )
    args = parser.parse_args(argv)

    if args.gui and args.cli:
        print("Please choose only one mode: --gui or --cli")
        return 1

    if args.app and args.cli:
        print("Please choose only one mode: --app or --cli")
        return 1

    # The packaged app is a window by default; the flags are still there for
    # anyone running the exe from a command prompt.
    want_app = args.app or (
        FROZEN and not (args.cli or args.qt or args.no_browser)
    )
    if want_app and focus_running_app():
        # Task Reporter is already open - that window was raised instead.
        return 0

    flush_pending_reports()

    if args.doctor:
        return run_doctor(args.platform)

    if args.list_reports:
        _print_recent_reports(limit=50)
        return 0

    if args.board:
        _print_board(show_filed=True)
        return 0

    if args.merge_days:
        return cli_merge_days(dry_run=args.dry_run)

    session = ReporterSession()
    _install_signal_handlers(session)

    if args.task is not None:
        return 0 if _cli_add_task(args.task) else 1

    if args.file_checked:
        # No prompt here: asking a script a question is how a cron job hangs.
        return 0 if cli_file_checked_tasks(session, confirm=False) else 1

    if args.report is not None:
        return 0 if save_report_text(args.report.strip(), "argument", session) else 1

    try:
        stdin_is_tty = sys.stdin is not None and sys.stdin.isatty()
    except (ValueError, AttributeError):
        stdin_is_tty = False

    # Piped input (echo "..." | task-report-maker.py --cli) is a one-shot save.
    if args.cli and not stdin_is_tty:
        return run_stdin_oneshot(session)

    script_path = os.path.abspath(__file__)
    platform_name = None
    attempts = []

    start_cli = not args.gui and not want_app and stdin_is_tty
    if not args.cli and not args.gui and not want_app and not stdin_is_tty:
        print("No interactive terminal attached - running the UI only.")

    # The desktop window first when it was asked for: it is the surface that
    # needs neither a terminal to stay open nor a browser tab to stay found.
    if want_app:
        exit_code = run_desktop_app(session, port_hint=args.port)
        if exit_code is not None:
            wait_for_quiet_workbook()
            _print_session_summary(session)
            sys.stdout.flush()
            sys.stderr.flush()
            os._exit(exit_code)
        # Getting a browser tab you did not ask for is only reasonable if it
        # comes with the reason - and the reason is in the log.
        show_app_error(
            "The desktop window could not open, so Task Reporter is falling "
            "back to your browser for this session.",
            fatal=False,
        )

    # The browser UI first, because it is the surface that does not depend on a
    # display server.  It only steps aside if no loopback port can be bound.
    if not args.cli and not args.qt:
        exit_code = run_web(
            session,
            start_cli=start_cli,
            port_hint=args.port,
            open_browser=not args.no_browser,
        )
        if exit_code is not None:
            wait_for_quiet_workbook()
            _print_session_summary(session)
            sys.stdout.flush()
            sys.stderr.flush()
            # Hard exit for the same reason as below: the console thread may be
            # parked inside input().
            os._exit(exit_code)
        print("  Trying the Qt window instead.")

    if not args.cli:
        # Only the Qt path needs an interpreter that can import PySide6, so the
        # re-exec dance is confined to it.
        ensure_movs_python(script_path, argv)
        relaunch_with_gui_if_possible(script_path, argv)
        platform_name, attempts = probe_qt_platform(args.platform)

    if platform_name:
        # The platform was verified in a subprocess to both load AND paint a
        # window, so QApplication() below will neither abort nor hang on an
        # invisible surface.
        os.environ["QT_QPA_PLATFORM"] = platform_name
        if platform_name not in VISUAL_QT_PLATFORMS:
            print(
                f"Warning: '{platform_name}' does not draw on a real screen. "
                "The window will be invisible;"
            )
            print("         drop --qt for the browser UI, or use --cli.")
        try:
            exit_code = run_gui(session, start_cli=start_cli)
        except Exception as exc:
            print(f"The Qt window failed to start ({exc}).")
        else:
            wait_for_quiet_workbook()
            _print_session_summary(session)
            sys.stdout.flush()
            sys.stderr.flush()
            # Hard exit: the console thread may be parked inside input(), and
            # this is what makes closing the window close the terminal too.
            os._exit(exit_code)

    if not args.cli:
        print("No UI could start, so the terminal console is taking over.")
        print(f"  Reason: {_describe_gui_failure(attempts)}")
        print("  Run with --doctor for the full check.")

    if not stdin_is_tty:
        return run_stdin_oneshot(session)

    if session.cli_active:
        # A console thread from the aborted GUI attempt already owns stdin;
        # let it finish rather than racing it with a second reader.
        session.wait_for_shutdown()
    else:
        cli_console_loop(session, dual=False)

    _print_session_summary(session)
    return 0


def _run_frozen() -> int:
    """main() for the packaged app, with nowhere for a traceback to go."""
    import traceback

    try:
        return main(sys.argv[1:])
    except SystemExit as exc:
        return exc.code if isinstance(exc.code, int) else 0
    except BaseException:
        traceback.print_exc()
        show_app_error("Task Reporter hit an error it could not recover from.")
        return 1


if __name__ == "__main__":
    if FROZEN:
        sys.exit(_run_frozen())
    try:
        sys.exit(main(sys.argv[1:]))
    except BrokenPipeError:
        # `--board | head` closes the pipe as soon as it has enough lines, which
        # is not a failure.  Point the fd at devnull first, or the interpreter's
        # own flush on the way out raises the same thing again and prints a
        # traceback after we have already decided all is well.
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        os._exit(0)
