"""Task Reporter - files task reports into task_reports.xlsx.

Two surfaces are available and, whenever a desktop session exists, both run at
the same time from a single process:

  * the PySide6 window  - type the report, press Ctrl+Enter
  * the terminal console - type the report at the `report>` prompt, press Enter

Closing either surface ends the whole session, so the GUI window and the
terminal window always disappear together.  When no usable display exists the
terminal console simply becomes the only surface instead of the program dying.
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
    from PySide6.QtGui import QKeySequence, QShortcut, QTextCursor, QAction

    GUI_AVAILABLE = True
except ModuleNotFoundError:
    GUI_AVAILABLE = False
    # Allow class declarations to be parsed when GUI libs are missing.
    QDialog = object
    QMainWindow = object

import openpyxl
from openpyxl import Workbook
import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import threading
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

_QT_PROBE_SNIPPET = (
    "import sys\n"
    "from PySide6.QtWidgets import QApplication\n"
    "QApplication([])\n"
    "sys.stdout.write('QT_PROBE_OK')\n"
)


def _qt_platform_candidates() -> list:
    candidates = []

    forced = (os.environ.get("QT_QPA_PLATFORM") or "").strip()
    if forced:
        candidates.append(forced)
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


def probe_qt_platform():
    """Find a Qt platform plugin that actually starts.

    This is the fix for "sometimes the UI doesn't open".  When no platform
    plugin can be initialised Qt calls qFatal(), which aborts the *entire
    process* with SIGABRT - not a Python exception, so a `try/except` around
    QApplication() never gets to run its fallback.  Each candidate is therefore
    tried inside a throwaway subprocess, where an abort costs us nothing.

    Returns (platform_name, failures).  platform_name is None when the GUI
    cannot start at all; failures lists (candidate, reason) pairs for `--doctor`.
    """
    if not GUI_AVAILABLE:
        return None, [("-", "PySide6 is not importable in this interpreter")]

    candidates = _qt_platform_candidates()
    if not candidates:
        return None, [("-", "neither DISPLAY nor WAYLAND_DISPLAY is set")]

    failures = []
    for candidate in candidates:
        env = os.environ.copy()
        env["QT_QPA_PLATFORM"] = candidate
        try:
            result = subprocess.run(
                [sys.executable, "-c", _QT_PROBE_SNIPPET],
                env=env,
                capture_output=True,
                text=True,
                timeout=QT_PROBE_TIMEOUT,
                check=False,
            )
        except subprocess.TimeoutExpired:
            failures.append((candidate, "probe timed out"))
            continue
        except Exception as exc:
            failures.append((candidate, f"probe could not run: {exc}"))
            continue

        if result.returncode == 0 and "QT_PROBE_OK" in (result.stdout or ""):
            return candidate, failures

        stderr_lines = [
            line.strip()
            for line in (result.stderr or "").splitlines()
            if line.strip() and "xcb-cursor0" not in line
        ]
        detail = stderr_lines[0] if stderr_lines else f"exit code {result.returncode}"
        failures.append((candidate, detail))

    return None, failures


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
        screen = QApplication.primaryScreen().availableGeometry()
        frame = self.frameGeometry()
        frame.moveCenter(screen.center())
        self.move(frame.topLeft())

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
        print("  The window and this console are both live.")
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


def _describe_gui_failure(failures: list) -> str:
    if not failures:
        return "no display was detected"
    return "; ".join(f"{name}: {reason}" for name, reason in failures)


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


def _print_session_summary(session: ReporterSession):
    print()
    print(f"Task Reporter closed - {session.reason or 'session ended'}.")
    if session.saved_count:
        plural = "s" if session.saved_count != 1 else ""
        print(f"{session.saved_count} report{plural} filed this session.")
    queued = pending_report_count()
    if queued:
        print(f"{queued} report(s) still queued for the workbook.")


def run_doctor() -> int:
    print("Task Reporter - environment check")
    print("-" * 68)
    print(f"  interpreter        : {sys.executable}")
    print(f"  script folder      : {BASE_DIR}")
    print(f"  workbook           : {EXCEL_FILE_PATH}")
    print(f"  workbook exists    : {os.path.exists(EXCEL_FILE_PATH)}")
    print(f"  queued reports     : {pending_report_count()}")
    print(f"  PySide6 importable : {GUI_AVAILABLE}")
    print(f"  DISPLAY            : {os.environ.get('DISPLAY') or '(unset)'}")
    print(f"  WAYLAND_DISPLAY    : {os.environ.get('WAYLAND_DISPLAY') or '(unset)'}")
    print(f"  stdin is a tty     : {bool(sys.stdin) and sys.stdin.isatty()}")
    print("-" * 68)

    platform_name, failures = probe_qt_platform()
    for name, reason in failures:
        print(f"  [x] platform '{name}' -> {reason}")
    if platform_name:
        print(f"  [ok] GUI will start using the '{platform_name}' platform plugin.")
        print("       Both the window and the terminal console will be available.")
        return 0

    print("  [x] No Qt platform plugin could start - the GUI is unavailable.")
    print("      The terminal console will be used instead; nothing is lost.")
    print("      On WSL this usually means WSLg is not running. Try:")
    print("        wsl.exe --shutdown      (from Windows, then reopen)")
    return 0


def main(argv: list) -> int:
    parser = argparse.ArgumentParser(
        prog="task-report-maker.py",
        description=(
            "File task reports into task_reports.xlsx. By default the GUI "
            "window and the terminal console run together; closing either one "
            "ends the session."
        ),
    )
    parser.add_argument(
        "--gui", action="store_true", help="GUI only (no terminal console)"
    )
    parser.add_argument(
        "--cli", action="store_true", help="Terminal console only (no window)"
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
        "--doctor",
        action="store_true",
        help="Explain whether the GUI can start, and why not if it cannot",
    )
    args = parser.parse_args(argv)

    if args.gui and args.cli:
        print("Please choose only one mode: --gui or --cli")
        return 1

    flush_pending_reports()

    if args.doctor:
        return run_doctor()

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
    failures = []

    if not args.cli:
        ensure_movs_python(script_path, argv)
        relaunch_with_gui_if_possible(script_path, argv)
        platform_name, failures = probe_qt_platform()

    if platform_name:
        # A GUI-capable platform plugin was verified in a subprocess, so
        # QApplication() below will not abort the process.
        os.environ["QT_QPA_PLATFORM"] = platform_name
        start_cli = not args.gui and stdin_is_tty
        if not args.gui and not stdin_is_tty:
            print("No interactive terminal attached - running the window only.")
        try:
            exit_code = run_gui(session, start_cli=start_cli)
        except Exception as exc:
            print(f"GUI failed to start ({exc}). Falling back to the terminal.")
        else:
            wait_for_quiet_workbook()
            _print_session_summary(session)
            sys.stdout.flush()
            sys.stderr.flush()
            # Hard exit: the console thread may be parked inside input(), and
            # this is what makes closing the window close the terminal too.
            os._exit(exit_code)

    if not args.cli:
        print("The GUI could not start, so the terminal console is taking over.")
        print(f"  Reason: {_describe_gui_failure(failures)}")
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
