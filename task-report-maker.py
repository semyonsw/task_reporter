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
import os
import sys
import argparse
import shutil
import subprocess
from datetime import datetime

EXCEL_FILE_NAME = "task_reports.xlsx"
EXCEL_FILE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), EXCEL_FILE_NAME
)
MAX_REPORT_LENGTH = 2000
RELAUNCH_ENV_FLAG = "TASK_REPORT_TK_RELAUNCH"
MOVS_PYTHON = "/root/miniconda3/envs/movs/bin/python"
MOVS_RELAUNCH_FLAG = "TASK_REPORT_MOVS_RELAUNCH"


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

    known_conda_python = "/root/miniconda3/envs/movs/bin/python"
    if os.path.exists(known_conda_python):
        candidates.append(known_conda_python)

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
    if os.environ.get(MOVS_RELAUNCH_FLAG) == "1":
        return

    current_python = os.path.abspath(sys.executable)
    target_python = os.path.abspath(MOVS_PYTHON)

    if current_python == target_python:
        return

    if not os.path.exists(target_python):
        return

    env = os.environ.copy()
    env[MOVS_RELAUNCH_FLAG] = "1"
    os.execvpe(target_python, [target_python, script_path, *args], env)


def append_report_to_excel(report_text: str):
    if not report_text or not report_text.strip():
        raise ValueError("The report cannot be empty.")

    if not os.path.exists(EXCEL_FILE_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "Reports"
        ws.append(["Date-Time", "Task Report"])
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 100
        wb.save(EXCEL_FILE_PATH)

    current_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    ws = wb.active
    ws.append([current_timestamp, report_text.strip()])
    wb.save(EXCEL_FILE_PATH)


def load_reports_from_excel() -> list:
    if not os.path.exists(EXCEL_FILE_PATH):
        return []
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH, read_only=True)
        ws = wb.active
        rows = []
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            dt_val = str(row[0]) if row[0] is not None else ""
            rpt_val = str(row[1]) if len(row) > 1 and row[1] is not None else ""
            rows.append((dt_val, rpt_val))
        wb.close()
        return rows
    except Exception:
        return []


def compact_excel():
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
                    str(row[0]) if row[0] is not None else "",
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
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Task Reporter")
        self.resize(1000, 720)
        self.setMinimumSize(920, 680)

        self._check_or_create_excel()
        self._build_ui()

        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._tick_clock)
        self._clock_timer.start(1000)
        self._tick_clock()

        self._bind_shortcuts()

        self.editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self._show_context_menu)

        self._update_counter()
        self._center_window()
        self.editor.setFocus()

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
            append_report_to_excel(report_text)
            self.editor.clear()
            self._update_counter()
            self.status_label.setText("Saved successfully \u2713")
            self.status_label.setStyleSheet(
                "font-size: 13px; font-weight: 600;"
                " color: #22C55E; background: transparent;"
            )
            QMessageBox.information(self, "Success", "Task saved successfully!")

        except PermissionError:
            self._set_status("Could not save: file is in use", danger=True)
            QMessageBox.critical(
                self,
                "Permission Error",
                f"Could not save to:\n{EXCEL_FILE_PATH}\n\n"
                "Is the Excel file currently open? Please close it and try again.",
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


def run_headless_mode(reason: str = None):
    display_available = bool(
        os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")
    )
    if reason == "forced_cli":
        print("Running in terminal mode.")
    elif display_available:
        print(
            "GUI mode unavailable in the current Python environment. "
            "Running in terminal mode."
        )
    else:
        print("No GUI display detected. Running in terminal mode.")

    print("Type your task report and press Enter.")

    if not sys.stdin.isatty():
        report_text = sys.stdin.read().strip()
    else:
        try:
            report_text = input("Task report: ").strip()
        except EOFError:
            report_text = ""

    if not report_text:
        print("Error: The report cannot be empty.")
        return 1

    try:
        append_report_to_excel(report_text)
        print(f"Task saved successfully to: {EXCEL_FILE_PATH}")
        return 0
    except PermissionError:
        print(
            f"Permission Error: Could not save to '{EXCEL_FILE_PATH}'. "
            "Is the file open?"
        )
        return 1
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Task reporter – GUI and CLI modes.")
    parser.add_argument("--gui", action="store_true", help="Force GUI mode")
    parser.add_argument("--cli", action="store_true", help="Force terminal mode")
    args = parser.parse_args()

    if args.gui and args.cli:
        print("Please choose only one mode: --gui or --cli")
        sys.exit(1)

    display_available = bool(
        os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")
    )
    should_run_gui = args.gui or (not args.cli and display_available)

    if should_run_gui:
        ensure_movs_python(os.path.abspath(__file__), sys.argv[1:])
        relaunch_with_gui_if_possible(os.path.abspath(__file__), sys.argv[1:])

        if not GUI_AVAILABLE:
            print(
                "PySide6 is not available in the current Python environment. "
                "Falling back to terminal mode."
            )
            sys.exit(run_headless_mode(reason="missing_gui"))

        if not display_available:
            print(
                "GUI requested, but no display detected. Falling back to terminal mode."
            )
            sys.exit(run_headless_mode(reason="missing_display"))

        try:
            qt_app = QApplication(sys.argv)
            qt_app.setApplicationName("Task Reporter")
            qt_app.setStyleSheet(DARK_QSS)

            window = TaskReporterApp()
            window.show()

            sys.exit(qt_app.exec())

        except Exception as exc:
            print(f"GUI failed to start ({exc}). Falling back to terminal mode.")
            sys.exit(run_headless_mode(reason="gui_failed"))

    terminal_reason = "forced_cli" if args.cli else None
    sys.exit(run_headless_mode(reason=terminal_reason))
