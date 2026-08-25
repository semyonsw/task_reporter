# task_reporter

Files task reports into `task_reports.xlsx` with an automatic timestamp.

There are two ways in, and by default **both run at once** from a single
process: a desktop window and a terminal console. File the report in whichever
one is in front of you — **closing either one closes the other.**

## Running it

From Windows, double-click `task_reporter.bat`.
From a Linux shell in this folder:

```bash
./task-report              # window + terminal console together
./task-report --cli        # terminal console only
./task-report --gui        # window only
./task-report -m "text"    # file one report and exit (scriptable)
./task-report --list       # print recent reports
./task-report --doctor     # explain whether the window can open, and why not
./task-report --platform xcb   # force a Qt platform (wayland / xcb)
```

Piped input works too: `echo "fixed the launcher" | ./task-report --cli`

### Launching it from the Desktop

Either works:

* **A shortcut** (right-click `task_reporter.bat` → *Send to* → *Desktop
  (create shortcut)*). A shortcut runs the original file where it lives, so
  nothing else is needed. This is the tidier option.
* **A plain copy** of `task_reporter.bat`. The file finds the project in this
  order: `%TASK_REPORT_DIR%` if set, then its own folder, then the
  `FALLBACKDIR` path recorded at the top of the file. A copy falls through to
  `FALLBACKDIR`, so **if you ever move the project, update that one line** —
  or set `TASK_REPORT_DIR` and it takes precedence over everything.

## Terminal console

Type the report at the `report>` prompt and press Enter. `\n` inside the line
becomes a line break.

| Command | Effect |
| --- | --- |
| `:m` / `:multi` | multi-line report, finish with a lone `.` |
| `:l` / `:list` | show the 10 most recent reports |
| `:p` / `:path` | print the workbook path |
| `:h` / `:help` | show the command list |
| `:q` / `:quit` | close the session (Ctrl+C and Ctrl+D do the same) |

## When the window will not open

Run `./task-report --doctor`. It tries each Qt platform plugin and reports
which one can actually put a window on screen.

The check is deliberately a *rendering* check, not just a "does Qt start"
check. Two different failures look identical from the outside:

* **No plugin loads.** Qt calls `qFatal()` here, which aborts the whole
  process — it is not a catchable exception, so the probe runs in a throwaway
  subprocess where an abort costs nothing.
* **A plugin loads, a window is created, and the compositor never paints it.**
  This is the WSLg "there is a taskbar button but no window" state. An
  init-only check sails straight past it, so the probe insists on a window that
  is both mapped and painted.

Either way the terminal console takes over automatically, so a display problem
never blocks a report.

### If the window is in the taskbar but not on screen

Usually a degraded WSLg session. From Windows:

```
wsl.exe --shutdown
```

then reopen. If it persists, force the X11 path instead of native Wayland:

```bash
./task-report --platform xcb          # or: --platform wayland
TASK_REPORT_QT_PLATFORM=xcb ./task-report
```

`--cli` always works regardless of the display state.

Note that a window title prefixed with something like `[WARN: COPY MODE]` does
not come from this app — that is WSLg's own annotation for a degraded
presentation path, and it is a strong hint that `wsl.exe --shutdown` is what
you want.

### Multi-monitor placement

The window is placed on the screen the mouse pointer is on, then clamped to
that screen's usable area. This matters because WSLg can report a primary
screen at a large virtual offset (here: `1920,724`) and can keep stale monitor
entries after a display change — centring on that blindly puts the window on a
monitor you may not be looking at, or one that is switched off.

## If the workbook is locked

Saving fails while `task_reports.xlsx` is open in Excel. Rather than losing the
report, it is queued in `.task_reports_pending.jsonl` and merged into the
workbook automatically on the next successful save. `--doctor` and the console
banner both show how many reports are still queued.
