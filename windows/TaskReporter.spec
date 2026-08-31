# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the Windows app.  Built by windows\\build.bat."""

import os
import sys

PROJECT_DIR = os.path.abspath(os.path.join(SPECPATH, ".."))


def conda_runtime_dlls():
    """OpenSSL and libffi, which conda keeps outside the interpreter folder.

    The build venv is layered on a conda interpreter, whose `_ssl.pyd` and
    `_ctypes.pyd` load their DLLs from the environment's Library\bin - a folder
    that only exists on PATH inside an activated conda shell.  PyInstaller
    therefore cannot see them on its own, and the exe would die on the first
    `import ssl` (which pywebview does).
    """
    root = os.path.join(sys.base_prefix, "Library", "bin")
    names = (
        "libssl-3-x64.dll",
        "libcrypto-3-x64.dll",
        "ffi-8.dll",
        "ffi-7.dll",
        "ffi.dll",
    )
    return [
        (os.path.join(root, name), ".")
        for name in names
        if os.path.isfile(os.path.join(root, name))
    ]

a = Analysis(
    [os.path.join(PROJECT_DIR, "task-report-maker.py")],
    pathex=[PROJECT_DIR],
    binaries=conda_runtime_dlls(),
    # Recorded so the exe can still find the workbook after it is copied or
    # shortcutted somewhere else - see _resolve_base_dir().
    datas=[(os.path.join(SPECPATH, "project_home.txt"), ".")],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    # The Qt window is a Linux/WSLg fallback and would add ~150 MB here for a
    # code path the app never takes on Windows.
    excludes=["PySide6", "shiboken6", "tkinter", "test", "unittest"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TaskReporter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    # No console: the whole point of the app is that there is no terminal to
    # keep open.  Output goes to .task_reporter_app.log instead.
    console=False,
    # The app catches its own crashes and shows them in a message box
    # naming the log file, which is more use than a raw traceback dialog.
    disable_windowed_traceback=True,
    icon=os.path.join(SPECPATH, "TaskReporter.ico"),
)
