# task_reporter

Files task reports into `task_reports.xlsx` with an automatic timestamp.

There are two ways in, and by default **both run at once** from a single
process: a browser window and a terminal console. File the report in whichever
one is in front of you — **closing either one closes the other.**

## Running it

From Windows, double-click `task_reporter.bat`.
From a Linux shell in this folder:

```bash
./task-report              # browser UI + terminal console together
./task-report --cli        # terminal console only
./task-report --gui        # browser UI only
./task-report -m "text"    # file one report and exit (scriptable)
./task-report --list       # print recent reports
./task-report --doctor     # check both UIs end to end
./task-report --no-browser # print the address instead of opening a browser
./task-report --port 8770  # serve the UI on a specific port
./task-report --qt         # the old desktop window (see "The Qt window")
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

## The browser UI

The UI is a page served on loopback and opened in your normal browser. Type the
report, press **Ctrl+Enter** (or click *Save Report*). `☰` shows previous
reports, where each one can be edited or deleted; `?` lists the shortcuts.

The address printed at start-up carries a **single-use token** that changes
every session, and the server refuses any request without it. That matters
here: under WSL the port is reachable from Windows too, so "bound to loopback"
is not on its own a closed door.

### Why a browser and not a desktop window

This used to be a PySide6 window, and it worked *most* of the time — which is
the problem. It rendered through WSLg, and WSLg has three failure modes that
all look like "the UI just did not work this time":

* a Qt platform plugin fails to load, and Qt answers with `qFatal()` — an abort
  rather than an exception, so nothing in the app can catch it;
* a plugin loads, a window is created, and the compositor never presents it —
  the "there is a taskbar button but no window" state;
* a stale multi-monitor layout puts the window on a screen that is switched off.

None of that is fixable from inside the app, and the render probe that used to
guard the window could only narrow the odds: it probed a frameless transparent
tool window, which does not always take the same presentation path as a real
window, and WSLg can degrade in the gap between the probe and the window
anyway. A green probe was never a promise.

Serving the page over loopback removes the whole failure class. WSL2 forwards
Windows `localhost` into the VM, so the page is fetched, drawn and placed by
native Windows processes: no X11, no Wayland, no compositor, no GPU path, and
nothing outside the Python standard library. On a plain Linux or macOS desktop
the same code opens the same page in the default browser.

### If the browser does not open

Nothing is blocked — the terminal console is already live, and the address is
printed at start-up so it can be pasted into any browser by hand.

`./task-report --doctor` checks the whole path: that a port can be bound, that
the page is reachable from this shell *and* from Windows (both `localhost` and
`127.0.0.1`, since some setups resolve `localhost` to `::1` first), and which
browser openers exist.

To force a particular opener:

```bash
TASK_REPORT_BROWSER="/mnt/c/Program Files/Google/Chrome/Application/chrome.exe" ./task-report
```

`TASK_REPORT_WEB_DEBUG=1` logs every request, which is the quickest way to see
whether a browser is reaching the page at all.

## The Qt window

The old desktop window is still there behind `--qt`, unchanged, for anyone who
prefers a native window and has a display that behaves:

```bash
./task-report --qt
./task-report --qt --platform xcb     # force X11 instead of native Wayland
```

It carries all the caveats above. If it is in the taskbar but not on screen,
that is a degraded WSLg session; `wsl.exe --shutdown` from Windows (then
reopen) usually clears it. A window title prefixed with something like
`[WARN: COPY MODE]` is WSLg's own annotation for that state, not something this
app prints.

The browser UI does not need PySide6 at all, so `--qt` is also the only mode
that cares which Python interpreter it runs under.

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

## If the workbook is locked

Saving fails while `task_reports.xlsx` is open in Excel. Rather than losing the
report, it is queued in `.task_reports_pending.jsonl` and merged into the
workbook automatically on the next successful save. `--doctor` and the console
banner both show how many reports are still queued.

Note that editing or deleting a report rewrites the sheet, which also rewrites
any Date-Time cell that Excel had stored as a real date into the same timestamp
as text. The reports themselves are untouched.
