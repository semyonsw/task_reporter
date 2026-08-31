# Installing Task Reporter

| Your machine | What to do |
|---|---|
| **Windows** | double-click **`Install.bat`** |
| **Linux / macOS / WSL** | `./install.sh` |

Everything below is only here for when that does not work.

---

## What the installer actually does

### On Windows — `Install.bat`

1. Finds a Python 3.10+ that is **actually usable** — it rejects a Python whose
   `ssl`, `sqlite3` or `venv` module is broken (a common state for
   Anaconda/Miniconda installs). If there is none, it offers to install Python
   3.12 for you with `winget`. Nothing is hardcoded to one machine's paths.
2. Creates the private build environment `.winenv\`, reusing an existing one
   only if it is healthy.
3. Installs `openpyxl`, `pywebview` and `pyinstaller` into it, retrying with a
   longer timeout, then relaxed certificate checks, then pre-built wheels only.
4. Runs the app from source (`--board`) to prove it works before packaging it.
5. **Builds `TaskReporter.exe`** with PyInstaller — the whole app in one file,
   which needs no Python to run. If the app is open at the time, it says so and
   keeps your working copy instead of half-replacing it.
6. Writes `Start Task Reporter.bat` (runs from source; used automatically if the
   exe could not be built) and puts a *Task Reporter* shortcut on your Desktop
   and in the Start menu.

Everything is logged to `install.log`. Re-running is safe. `task_reports.xlsx`
and `task_board.json` are never touched by the installer.

### On Linux / macOS / WSL — `./install.sh`

Creates `.venv/`, installs `openpyxl`, proves the app starts, makes
`./task-report` executable, and offers `pywebview` for the desktop-window mode.
There is no `.exe` on this side — `./task-report` is the launcher.

---

## Troubleshooting

| Message | What it means | Fix |
|---|---|---|
| *Python is required and no usable copy could be installed* | No Python 3.10+ | Install 3.12 from [python.org](https://www.python.org/downloads/windows/), tick "Add python.exe to PATH" |
| *the existing .winenv environment is unusable: its SSL support is broken* | The build env was layered on a conda Python without its OpenSSL DLLs | Nothing to do — the installer deletes and rebuilds it for you |
| *no ready-made build of `<package>` exists for this Python version* | Python newer than the packages | The installer fetches 3.12 and retries by itself; otherwise install 3.12, delete `.winenv\`, re-run |
| *Task Reporter is running right now, so its .exe cannot be replaced* | Exactly that | Close the window, then run `Install.bat` again. Your current exe keeps working meanwhile |
| *antivirus blocked the build* | PyInstaller output looks suspicious to some scanners | Allow this folder in your antivirus, or use `Start Task Reporter.bat`, which runs the same app from source |
| *the exe could not be built* | Anything else | Not fatal — `Start Task Reporter.bat` runs the same app. Details at the end of `install.log` |
| *the app cannot load openpyxl* | Broken environment | Delete `.winenv\` (or `.venv/`) and re-run the installer |

### The window opens and closes again

The app writes what a terminal would have shown to `.task_reporter_app.log`.
Read the last lines of that file — it names the real error.

### "Another copy is already running"

`.task_reporter_app.lock` records the window that is already open, and a second
launch brings that one to the front instead of opening a duplicate. If the app
was killed rather than closed, delete the lock file.

### The workbook is open in Excel

Excel locks the file, so a report cannot be written. The app queues it in
`.task_reports_pending.jsonl` and files it once Excel lets go. Close Excel.

### The Desktop shortcut opens the wrong folder

The shortcut points at `TaskReporter.exe` where it lives, and the exe finds the
project through `windows\project_home.txt`, which is baked in at build time. If
you move the project, run `Install.bat` again — it rewrites that file.

`task_reporter.bat` (the WSL version with the terminal) has its own
`FALLBACKDIR` line near the top for the same purpose, or set
`TASK_REPORT_DIR`, which overrides everything.

### Starting over from scratch

```bat
rmdir /s /q .winenv build
del TaskReporter.exe install.log
Install.bat
```

Your reports and board are untouched by this.

---

## Uninstalling

1. Close the app.
2. `windows\install_shortcut.bat /remove` (or delete the shortcuts by hand).
3. Copy `task_reports.xlsx` somewhere safe if you want to keep it.
4. Delete this folder.
