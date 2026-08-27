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
import argparse
import http.server
import json
import os
import secrets
import shutil
import signal
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_NAME = "task_reports.xlsx"
EXCEL_FILE_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
# Reports land here when the workbook cannot be written (usually because it is
# open in Excel).  They are merged back in automatically on the next save.
PENDING_FILE_PATH = os.path.join(BASE_DIR, ".task_reports_pending.jsonl")

MAX_REPORT_LENGTH = 2000
TIMESTAMP_FORMAT = "%d/%m/%Y %H:%M:%S"

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
    wb.save(EXCEL_FILE_PATH)


def _ensure_workbook():
    if not os.path.exists(EXCEL_FILE_PATH):
        _create_workbook()


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
        try:
            _ensure_workbook()
            wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
            ws = wb.active
            for timestamp, report in entries:
                ws.append([timestamp, report])
            wb.save(EXCEL_FILE_PATH)
        except Exception:
            return 0
        _drop_pending()
        return len(entries)


def append_report_to_excel(report_text: str) -> str:
    """Append a report and return the timestamp that was written.

    Raises ReportQueuedError when the workbook cannot be written - the report
    is safely queued in that case rather than lost.
    """
    if not report_text or not report_text.strip():
        raise ValueError("The report cannot be empty.")

    text = report_text.strip()
    timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)

    with _EXCEL_LOCK:
        pending = _read_pending()
        try:
            _ensure_workbook()
            wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
            ws = wb.active
            for queued_timestamp, queued_report in pending:
                ws.append([queued_timestamp, queued_report])
            ws.append([timestamp, text])
            wb.save(EXCEL_FILE_PATH)
        except (PermissionError, OSError) as exc:
            _queue_pending(timestamp, text)
            raise ReportQueuedError(timestamp, exc) from exc

        if pending:
            _drop_pending()

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
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 100
        wb.save(EXCEL_FILE_PATH)


def delete_report_from_excel(row_index: int):
    with _EXCEL_LOCK:
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
        wb.save(EXCEL_FILE_PATH)
        compact_excel()


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
<meta name="color-scheme" content="dark">
<title>Task Reporter</title>
<style>
*, *::before, *::after { box-sizing: border-box; }
html, body { height: 100%; }
body {
  margin: 0;
  background: #0F1117;
  color: #E2E8F0;
  font-family: "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  font-size: 14px;
  -webkit-font-smoothing: antialiased;
}
button { font-family: inherit; cursor: pointer; }
.page {
  display: flex; flex-direction: column; gap: 12px;
  padding: 20px; height: 100%; min-height: 0;
}
.card { background: #161B27; border: 1px solid #1E2640; border-radius: 14px; }

/* header */
.header { display: flex; align-items: center; gap: 12px; padding: 16px 22px; flex: none; }
.title { margin: 0 auto 0 0; font-size: 28px; font-weight: 700; color: #E8F0FE; }
.header-right { display: flex; align-items: center; gap: 12px; }
.clock { font-size: 13px; color: #5B6EA6; font-variant-numeric: tabular-nums; }
.icon-btn {
  width: 28px; height: 28px; padding: 0; border-radius: 14px;
  background: #1E2640; color: #7B90D4; border: 1px solid #2D3860;
  font-size: 14px; font-weight: 700; line-height: 1;
  display: grid; place-items: center;
}
.icon-btn:hover { background: #2D3860; color: #A8BFFF; }
.icon-btn:active { background: #3B4A80; }
.separator { height: 1px; background: #1E2640; flex: none; }

/* editor card */
.content {
  display: flex; flex-direction: column; gap: 12px;
  padding: 20px 24px; flex: 1; min-height: 0;
}
.heading { margin: 0; font-size: 20px; font-weight: 700; color: #E8F0FE; }
#editor {
  flex: 1; min-height: 160px; resize: none; padding: 12px;
  background: #0D1020; color: #CBD5E1;
  border: 2px solid #1E2640; border-radius: 10px;
  font-size: 15px; font-family: inherit; line-height: 1.55; outline: none;
}
#editor:focus { border-color: #3B82F6; }
#editor::placeholder { color: #3D4F7A; }
#editor::selection { background: #2563EB; color: #FFFFFF; }

.footer { display: flex; align-items: center; gap: 12px; padding-top: 4px; flex: none; }
.status { margin-right: auto; min-width: 160px; font-size: 13px; font-weight: 600; color: #3B82F6; }
.status.danger  { color: #EF4444; }
.status.success { color: #22C55E; }
.status.warn    { color: #F59E0B; }
.progress { flex: none; width: 200px; height: 8px; background: #1E2640; border-radius: 4px; overflow: hidden; }
.progress-fill { width: 0; height: 100%; background: #3B82F6; border-radius: 4px; transition: width .12s linear; }
.progress-fill.danger { background: #EF4444; }
.counter { min-width: 80px; text-align: right; font-size: 13px; color: #5B6EA6; font-variant-numeric: tabular-nums; }
.counter.danger { color: #EF4444; }
.save-btn {
  flex: none; min-width: 156px; height: 40px; border: none; border-radius: 10px;
  background: linear-gradient(#3B82F6, #2563EB); color: #FFFFFF;
  font-size: 14px; font-weight: 700;
}
.save-btn:hover  { background: linear-gradient(#60A5FA, #3B82F6); }
.save-btn:active { background: #1D4ED8; }
.save-btn:disabled { background: #1E2640; color: #3D4F7A; cursor: default; }

/* dialogs */
.backdrop {
  position: fixed; inset: 0; z-index: 20; padding: 24px;
  background: rgba(5, 7, 14, .74);
  display: none; align-items: center; justify-content: center;
}
.backdrop.open { display: flex; }
.modal {
  display: flex; flex-direction: column; gap: 12px;
  background: #161B27; border: 1px solid #1E2640; border-radius: 12px;
  padding: 20px 24px 16px; max-height: 100%; width: 100%;
}
.modal-sm { max-width: 480px; }
.modal-md { max-width: 600px; }
.modal-lg { max-width: 940px; height: 100%; }
.modal-head { display: flex; align-items: baseline; gap: 12px; flex: none; }
.modal-head h3 { margin: 0 auto 0 0; font-size: 17px; font-weight: 700; color: #E8F0FE; }
.modal-count { font-size: 12px; color: #5B6EA6; }
.modal-actions { display: flex; justify-content: flex-end; gap: 10px; flex: none; padding-top: 4px; }
.field-label { font-size: 12px; font-weight: 600; color: #5B6EA6; }

.btn {
  border-radius: 8px; font-size: 13px; font-weight: 600;
  min-height: 34px; padding: 0 24px;
  background: #1E2640; color: #7B90D4; border: 1px solid #2D3860;
}
.btn:hover  { background: #2D3860; color: #A8BFFF; }
.btn:active { background: #3B4A80; }
.btn-primary {
  border: none; color: #FFFFFF;
  background: linear-gradient(#3B82F6, #2563EB); font-weight: 700;
}
.btn-primary:hover  { background: linear-gradient(#60A5FA, #3B82F6); color: #FFFFFF; }
.btn-primary:active { background: #1D4ED8; }
.btn-row { min-height: 28px; padding: 0 14px; font-size: 12px; border-radius: 6px; white-space: nowrap; }
.btn-danger { background: #2A1520; color: #F87171; border-color: #5B2030; }
.btn-danger:hover  { background: #3D1A2A; color: #FCA5A5; }
.btn-danger:active { background: #4A1F30; }

/* history table */
.table-wrap {
  flex: 1; min-height: 0; overflow: auto;
  background: #0D1020; border: 1px solid #1E2640; border-radius: 8px;
}
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th {
  position: sticky; top: 0; z-index: 1; text-align: left;
  background: #161B27; color: #7B90D4; font-weight: 700;
  padding: 8px 12px; border-bottom: 2px solid #2D3860;
}
td { padding: 8px 12px; color: #CBD5E1; border-bottom: 1px solid #1E2640; vertical-align: top; }
tbody tr:nth-child(even) { background: #131728; }
tbody tr:hover { background: #1A2036; }
.col-num  { width: 52px;  text-align: center; color: #5B6EA6; font-variant-numeric: tabular-nums; }
.col-when { width: 160px; white-space: nowrap; font-variant-numeric: tabular-nums; }
.col-text { white-space: pre-wrap; overflow-wrap: anywhere; }
.col-act  { width: 1%; }
.empty { padding: 28px 12px; text-align: center; color: #5B6EA6; }

/* edit dialog fields */
#editWhen, #editText {
  width: 100%; padding: 6px 10px;
  background: #0D1020; color: #CBD5E1;
  border: 2px solid #1E2640; border-radius: 8px;
  font-size: 14px; font-family: inherit; outline: none;
}
#editText { min-height: 200px; resize: vertical; line-height: 1.5; }
#editWhen:focus, #editText:focus { border-color: #3B82F6; }
.edit-error { min-height: 16px; font-size: 12px; font-weight: 600; color: #EF4444; }

/* shortcuts table */
.keys { width: 100%; border-collapse: collapse; }
.keys tr:nth-child(odd)  { background: #0F1117; }
.keys tr:nth-child(even) { background: #1A1F33; }
.keys td { border: none; padding: 7px 10px; }
.keys td:first-child { width: 1%; white-space: nowrap; }
kbd {
  display: inline-block; padding: 2px 8px; border-radius: 4px;
  background: #1E2640; color: #A8BFFF;
  font-family: ui-monospace, SFMono-Regular, Consolas, monospace; font-size: 12px;
}
.hint { margin: 0; font-size: 12px; color: #5B6EA6; line-height: 1.6; }

/* the session-over curtain */
#gone {
  position: fixed; inset: 0; z-index: 40; display: none;
  align-items: center; justify-content: center; text-align: center;
  background: #0F1117; color: #5B6EA6; font-size: 15px; padding: 24px;
}
#gone.open { display: flex; }
</style>
</head>
<body>
<div class="page">
  <header class="card header">
    <h1 class="title">Task Reporter</h1>
    <div class="header-right">
      <span class="clock" id="clock"></span>
      <button class="icon-btn" id="historyBtn" title="View previous reports" aria-label="View previous reports">&#9776;</button>
      <button class="icon-btn" id="helpBtn" title="View keyboard shortcuts" aria-label="View keyboard shortcuts">?</button>
    </div>
  </header>

  <div class="separator"></div>

  <main class="card content">
    <h2 class="heading">What did you accomplish?</h2>
    <textarea id="editor" placeholder="Describe your work clearly and concisely&#8230;" autofocus></textarea>
    <div class="footer">
      <span class="status" id="status">Ready</span>
      <div class="progress"><div class="progress-fill" id="progressFill"></div></div>
      <span class="counter" id="counter"></span>
      <button class="save-btn" id="saveBtn">Save Report</button>
    </div>
  </main>
</div>

<div class="backdrop" id="historyBackdrop">
  <div class="modal modal-lg">
    <div class="modal-head">
      <h3>Previous Reports</h3>
      <span class="modal-count" id="historyCount"></span>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th class="col-num">#</th>
            <th class="col-when">Date-Time</th>
            <th class="col-text">Task Report</th>
            <th class="col-act"></th>
            <th class="col-act"></th>
          </tr>
        </thead>
        <tbody id="historyBody"></tbody>
      </table>
    </div>
    <div class="modal-actions">
      <button class="btn" data-close>Close</button>
    </div>
  </div>
</div>

<div class="backdrop" id="editBackdrop">
  <div class="modal modal-md">
    <div class="modal-head"><h3>Edit Report</h3></div>
    <span class="field-label">Date-Time</span>
    <input type="text" id="editWhen" autocomplete="off" spellcheck="false">
    <span class="field-label">Task Report</span>
    <textarea id="editText"></textarea>
    <div class="edit-error" id="editError"></div>
    <div class="modal-actions">
      <button class="btn" data-close>Cancel</button>
      <button class="btn btn-primary" id="editSave">Save Changes</button>
    </div>
  </div>
</div>

<div class="backdrop" id="helpBackdrop">
  <div class="modal modal-sm">
    <div class="modal-head"><h3>Keyboard Shortcuts</h3></div>
    <table class="keys">
      <tr><td><kbd>Ctrl + Enter</kbd></td><td>Save the report</td></tr>
      <tr><td><kbd>Ctrl + A</kbd></td><td>Select all text</td></tr>
      <tr><td><kbd>Ctrl + Z</kbd></td><td>Undo</td></tr>
      <tr><td><kbd>Ctrl + Y</kbd> / <kbd>Ctrl + Shift + Z</kbd></td><td>Redo</td></tr>
      <tr><td><kbd>Ctrl + Backspace</kbd></td><td>Delete previous word</td></tr>
      <tr><td><kbd>Ctrl + Delete</kbd></td><td>Delete next word</td></tr>
      <tr><td><kbd>Ctrl + C</kbd> / <kbd>Ctrl + X</kbd> / <kbd>Ctrl + V</kbd></td><td>Copy / cut / paste</td></tr>
      <tr><td><kbd>Ctrl + H</kbd></td><td>Previous reports</td></tr>
      <tr><td><kbd>Esc</kbd></td><td>Close a dialog</td></tr>
      <tr><td>Right-click</td><td>Open context menu</td></tr>
    </table>
    <p class="hint">
      The terminal console is live at the same time - a report filed there
      shows up here, and closing either one closes the other.
    </p>
    <div class="modal-actions">
      <button class="btn" data-close>Close</button>
    </div>
  </div>
</div>

<div id="gone"><div id="goneText"></div></div>

<script>
"use strict";
const CFG = %%CONFIG%%;

const $ = (id) => document.getElementById(id);
const editor = $("editor");
const statusEl = $("status");
const counterEl = $("counter");
const fillEl = $("progressFill");
const saveBtn = $("saveBtn");

/* ---------------------------------------------------------------- transport */

// Every request carries the session token.  The server refuses anything
// without it, which matters on WSL: the port is reachable from Windows, so
// "bound to loopback" is not on its own a closed door.
async function api(path, body) {
  const res = await fetch(path + "?t=" + encodeURIComponent(CFG.token), {
    method: body === undefined ? "GET" : "POST",
    headers: body === undefined ? {} : { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
    cache: "no-store",
  });
  if (!res.ok) throw new Error("HTTP " + res.status);
  return res.json();
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

function updateCounter() {
  const length = editor.value.length;
  const over = length > CFG.maxLength;
  counterEl.textContent = length + " / " + CFG.maxLength;
  counterEl.className = "counter" + (over ? " danger" : "");
  fillEl.style.width = Math.min(100, (length / CFG.maxLength) * 100) + "%";
  fillEl.className = "progress-fill" + (over ? " danger" : "");
  if (statusTimer) return;
  if (over) {
    setStatus("Report is too long  (" + (length - CFG.maxLength) + " chars over limit)", "danger");
  } else {
    setStatus("Ready", null);
  }
}

/* --------------------------------------------------------------------- save */

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
      updateCounter();
      flashStatus(
        (out.queued ? "Queued at " : "Saved at ") + out.timestamp + " ✓",
        out.queued ? "warn" : "success"
      );
      if (out.queued) {
        alert(
          "The workbook could not be written:\n" + CFG.workbook + "\n\n" +
          "It is probably open in Excel. Your report was kept in\n" +
          CFG.pendingFile + "\n\nand will be merged in automatically once the " +
          "file is free."
        );
      }
    } else {
      flashStatus(out.message || "Could not save the report", "danger");
    }
  } catch (err) {
    // The report is still in the box, so nothing is lost by retrying.
    flashStatus("Lost contact with the reporter - your text is still here", "danger");
  } finally {
    saving = false;
    saveBtn.disabled = false;
    editor.focus();
  }
}

/* ------------------------------------------------------------------ dialogs */

const openStack = [];

function openModal(id) {
  const el = $(id);
  el.classList.add("open");
  if (!openStack.includes(id)) openStack.push(id);
}

function closeModal(id) {
  $(id).classList.remove("open");
  const at = openStack.indexOf(id);
  if (at !== -1) openStack.splice(at, 1);
  if (!openStack.length) editor.focus();
}

document.querySelectorAll("[data-close]").forEach((btn) => {
  btn.addEventListener("click", () => {
    const backdrop = btn.closest(".backdrop");
    if (backdrop) closeModal(backdrop.id);
  });
});

document.querySelectorAll(".backdrop").forEach((backdrop) => {
  backdrop.addEventListener("mousedown", (event) => {
    if (event.target === backdrop) closeModal(backdrop.id);
  });
});

/* ------------------------------------------------------------------ history */

let historyRows = [];

async function loadHistory() {
  const body = $("historyBody");
  try {
    const out = await api("/api/list");
    historyRows = out.reports || [];
  } catch (err) {
    body.innerHTML = "";
    const cell = document.createElement("td");
    cell.className = "empty";
    cell.colSpan = 5;
    cell.textContent = "Could not read the workbook.";
    const row = document.createElement("tr");
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }

  const total = historyRows.length;
  $("historyCount").textContent = total + (total === 1 ? " report" : " reports");
  body.innerHTML = "";

  if (!total) {
    const cell = document.createElement("td");
    cell.className = "empty";
    cell.colSpan = 5;
    cell.textContent = "No reports yet.";
    const row = document.createElement("tr");
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }

  // Newest first, but numbered by their real position in the workbook.
  for (let i = total - 1; i >= 0; i--) {
    const item = historyRows[i];
    const row = document.createElement("tr");

    const num = document.createElement("td");
    num.className = "col-num";
    num.textContent = String(i + 1);
    row.appendChild(num);

    const when = document.createElement("td");
    when.className = "col-when";
    when.textContent = item.datetime;
    row.appendChild(when);

    // textContent, not innerHTML: report text is arbitrary user input.
    const text = document.createElement("td");
    text.className = "col-text";
    text.textContent = item.text;
    row.appendChild(text);

    const editCell = document.createElement("td");
    editCell.className = "col-act";
    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-row";
    editBtn.textContent = "Edit";
    editBtn.addEventListener("click", () => openEdit(item.index));
    editCell.appendChild(editBtn);
    row.appendChild(editCell);

    const delCell = document.createElement("td");
    delCell.className = "col-act";
    const delBtn = document.createElement("button");
    delBtn.className = "btn btn-row btn-danger";
    delBtn.textContent = "Delete";
    delBtn.addEventListener("click", () => deleteReport(item.index));
    delCell.appendChild(delBtn);
    row.appendChild(delCell);

    body.appendChild(row);
  }
}

async function openHistory() {
  openModal("historyBackdrop");
  await loadHistory();
}

async function deleteReport(index) {
  const item = historyRows.find((row) => row.index === index);
  if (!item) return;
  if (!confirm("Delete report #" + (index + 1) + "?\n\nThis action cannot be undone.")) return;
  try {
    const out = await api("/api/delete", { index });
    if (!out.ok) { alert(out.message || "Could not delete the report."); return; }
    await loadHistory();
    flashStatus("Report #" + (index + 1) + " deleted", "success");
  } catch (err) {
    alert("Lost contact with the reporter. Nothing was deleted.");
  }
}

/* --------------------------------------------------------------------- edit */

let editIndex = null;

function openEdit(index) {
  const item = historyRows.find((row) => row.index === index);
  if (!item) return;
  editIndex = index;
  $("editWhen").value = item.datetime;
  $("editText").value = item.text;
  $("editError").textContent = "";
  openModal("editBackdrop");
  $("editText").focus();
}

async function saveEdit() {
  if (editIndex === null) return;
  const text = $("editText").value.trim();
  const when = $("editWhen").value.trim();
  const error = $("editError");

  if (!text) { error.textContent = "Report text cannot be empty."; return; }
  if (text.length > CFG.maxLength) {
    error.textContent = "Report is too long (" + text.length + " chars). Limit is " + CFG.maxLength + ".";
    return;
  }

  error.textContent = "";
  try {
    const out = await api("/api/update", { index: editIndex, datetime: when, text });
    if (!out.ok) { error.textContent = out.message || "Could not save the changes."; return; }
    closeModal("editBackdrop");
    editIndex = null;
    await loadHistory();
    flashStatus("Report updated", "success");
  } catch (err) {
    error.textContent = "Lost contact with the reporter. Nothing was changed.";
  }
}

/* --------------------------------------------------------- session heartbeat */

// The page tells the server it is alive, and says goodbye on the way out, so
// that closing the tab ends the session the way closing a window used to.
// Missed pings alone never end it: browsers throttle timers in background
// tabs, and a false "the browser is gone" would take the terminal with it.
let ended = false;

function endSession(reason) {
  if (ended) return;
  ended = true;
  saveBtn.disabled = true;
  editor.readOnly = true;
  $("goneText").textContent =
    "Task Reporter session ended" + (reason ? " - " + reason : "") + ".\n" +
    "You can close this tab.";
  $("gone").classList.add("open");
}

async function ping() {
  if (ended) return;
  try {
    const out = await api("/api/ping", { client: CFG.clientId });
    if (out.shuttingDown) { endSession(out.reason); return; }
    for (const event of out.events || []) {
      flashStatus(
        "Filed from the terminal at " + event.timestamp +
          (event.queued ? " (queued - workbook is locked)" : ""),
        event.queued ? "warn" : "success"
      );
    }
  } catch (err) {
    // A single miss means nothing - the server may just be busy saving.
  }
}

window.addEventListener("pagehide", (event) => {
  // persisted means the page is going into the back/forward cache and will be
  // reused - it has not been closed, so it must not say goodbye.
  if (ended || event.persisted) return;
  const url = "/api/bye?t=" + encodeURIComponent(CFG.token) +
              "&client=" + encodeURIComponent(CFG.clientId);
  // sendBeacon survives teardown; fetch(keepalive) is the fallback.
  if (!(navigator.sendBeacon && navigator.sendBeacon(url))) {
    fetch(url, { method: "POST", keepalive: true }).catch(() => {});
  }
});

// Restored from the back/forward cache, unfrozen, or brought back to the
// foreground: check in at once so any goodbye already in flight is cancelled
// before its countdown can run out.
window.addEventListener("pageshow", () => { ping(); });
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") ping();
});

/* ------------------------------------------------------------------ wire-up */

saveBtn.addEventListener("click", saveReport);
editor.addEventListener("input", updateCounter);
$("historyBtn").addEventListener("click", openHistory);
$("helpBtn").addEventListener("click", () => openModal("helpBackdrop"));
$("editSave").addEventListener("click", saveEdit);

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && openStack.length) {
    event.preventDefault();
    closeModal(openStack[openStack.length - 1]);
    return;
  }
  const accel = event.ctrlKey || event.metaKey;
  if (accel && event.key === "Enter") {
    event.preventDefault();
    if (openStack[openStack.length - 1] === "editBackdrop") saveEdit();
    else if (!openStack.length) saveReport();
    return;
  }
  if (accel && !event.shiftKey && (event.key === "h" || event.key === "H")) {
    event.preventDefault();
    if (!openStack.includes("historyBackdrop")) openHistory();
  }
});

function tickClock() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  $("clock").textContent =
    pad(now.getDate()) + "/" + pad(now.getMonth() + 1) + "/" + now.getFullYear() +
    "  " + pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());
}

tickClock();
setInterval(tickClock, 1000);
updateCounter();
editor.focus();
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

    def _read_json(self) -> dict:
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            return {}
        if length <= 0:
            return {}
        if length > WEB_MAX_BODY_BYTES:
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
            self._send(204, b"", "image/x-icon")
            return

        if not self._authorised():
            return

        if path in ("/", "/index.html"):
            self._state.note_page_served()
            self._send(200, _render_web_page(self._state.token), "text/html; charset=utf-8")
            return

        if path == "/api/health":
            self._send_json({"ok": True})
            return

        if path == "/api/list":
            self._send_json({"reports": _web_list_reports()})
            return

        self._send_text(404, "not found")

    def do_HEAD(self):
        self.do_GET()

    def do_POST(self):
        if not self._authorised():
            return

        path = self._path()

        if path == "/api/bye":
            # Sent by sendBeacon on tab close, so it carries no body.
            values = self._query().get("client") or []
            self._state.note_bye(values[0] if values else "")
            self._send_json({"ok": True})
            return

        payload = self._read_json()

        if path == "/api/ping":
            self._state.note_ping(str(payload.get("client") or ""))
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
                }
            )
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

        self._send_text(404, "not found")


# ---------------------------------------------------------------------------
# Browser UI actions - thin wrappers over the same workbook helpers the
# terminal console uses, so both surfaces cannot drift apart.
# ---------------------------------------------------------------------------


def _render_web_page(token: str) -> bytes:
    config = {
        "token": token,
        "clientId": secrets.token_urlsafe(9),
        "maxLength": MAX_REPORT_LENGTH,
        "pingSeconds": 3,
        "workbook": EXCEL_FILE_PATH,
        "pendingFile": PENDING_FILE_PATH,
    }
    page = WEB_PAGE_TEMPLATE.replace("%%CONFIG%%", json.dumps(config))
    return page.encode("utf-8")


def _web_list_reports() -> list:
    return [
        {"index": index, "datetime": when, "text": text}
        for index, (when, text) in enumerate(load_reports_from_excel())
    ]


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
    (":p  /  :path", "print the workbook path"),
    (":h  /  :help", "show this help"),
    (":q  /  :quit", "close the session (Ctrl+C and Ctrl+D do the same)"),
]


def _print_cli_help():
    print("  Type your report and press Enter to file it.")
    print("  Use \\n inside the line for a manual line break.")
    for keys, desc in CLI_COMMANDS:
        print(f"    {keys:<16} {desc}")


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
    queued = pending_report_count()
    if queued:
        print(f"  {queued} report(s) waiting to be merged into the workbook.")
    print("-" * 68)
    _print_cli_help()
    print("-" * 68)


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


def run_doctor(forced_platform: str = None) -> int:
    print("Task Reporter - environment check")
    print("=" * 68)
    print(f"  interpreter        : {sys.executable}")
    print(f"  script folder      : {BASE_DIR}")
    print(f"  workbook           : {EXCEL_FILE_PATH}")
    print(f"  workbook exists    : {os.path.exists(EXCEL_FILE_PATH)}")
    print(f"  queued reports     : {pending_report_count()}")
    print(f"  running under WSL  : {is_wsl()}")
    print(f"  windows interop    : {bool(shutil.which('powershell.exe'))}")
    print(f"  PySide6 importable : {GUI_AVAILABLE}")
    print(f"  DISPLAY            : {os.environ.get('DISPLAY') or '(unset)'}")
    print(f"  WAYLAND_DISPLAY    : {os.environ.get('WAYLAND_DISPLAY') or '(unset)'}")
    print(f"  stdin is a tty     : {bool(sys.stdin) and sys.stdin.isatty()}")
    print("=" * 68)

    web_ok = _doctor_web_section()
    _doctor_qt_section(forced_platform)

    print()
    print("-" * 68)
    if web_ok:
        print("  Verdict: run ./task-report - the browser UI will come up.")
    else:
        print("  Verdict: the terminal console still works everywhere:")
        print("    ./task-report --cli")
    return 0


def main(argv: list) -> int:
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

    flush_pending_reports()

    if args.doctor:
        return run_doctor(args.platform)

    if args.list_reports:
        _print_recent_reports(limit=50)
        return 0

    session = ReporterSession()
    _install_signal_handlers(session)

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

    start_cli = not args.gui and stdin_is_tty
    if not args.cli and not args.gui and not stdin_is_tty:
        print("No interactive terminal attached - running the UI only.")

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


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
