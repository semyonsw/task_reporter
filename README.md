# task_reporter

Files task reports into `task_reports.xlsx` with an automatic timestamp.

On Windows it is a **desktop app**: double-click *Task Reporter*, a window
opens, and there is no terminal and no browser tab to keep track of. See
[The Windows app](#the-windows-app).

From a Linux shell it is the same UI served into your browser, alongside a
terminal console — by default **both run at once** from a single process, and
**closing either one closes the other.**

Reports can also be *built up* rather than typed in one go. The **task board**
keeps the running list of what has to be done — a day, the projects worked on
that day in `[square brackets]`, and the tasks under each one with a box to
tick. Ticking the boxes and pressing **File Checked Tasks** turns them into
report rows, one row per day, grouped under their project headers. See
[The task board](#the-task-board).

## Running it

On Windows, double-click **Task Reporter** on the Desktop — or
`TaskReporter.exe` in this folder. `task_reporter.bat` still runs the WSL
version with its terminal, for when you want the console commands.

From a Linux shell in this folder:

```bash
./task-report              # browser UI + terminal console together
./task-report --cli        # terminal console only
./task-report --gui        # browser UI only
./task-report --app        # native desktop window, no terminal, no browser
./task-report -m "text"    # file one report and exit (scriptable)
./task-report --list       # print recent reports
./task-report --board      # print the task board
./task-report -t "text"    # add a task to the board and exit
./task-report --file-checked   # file every ticked task and exit
./task-report --doctor     # check both UIs end to end
./task-report --no-browser # print the address instead of opening a browser
./task-report --port 8770  # serve the UI on a specific port
./task-report --qt         # the old desktop window (see "The Qt window")
```

Piped input works too: `echo "fixed the launcher" | ./task-report --cli`

### Launching it from the Desktop

The Windows app puts itself there — `windows\install_shortcut.bat` makes the
Desktop and Start menu shortcuts. For `task_reporter.bat` (the WSL version
with the terminal), either works:

* **A shortcut** (right-click `task_reporter.bat` → *Send to* → *Desktop
  (create shortcut)*). A shortcut runs the original file where it lives, so
  nothing else is needed. This is the tidier option.
* **A plain copy** of `task_reporter.bat`. The file finds the project in this
  order: `%TASK_REPORT_DIR%` if set, then its own folder, then the
  `FALLBACKDIR` path recorded at the top of the file. A copy falls through to
  `FALLBACKDIR`, so **if you ever move the project, update that one line** —
  or set `TASK_REPORT_DIR` and it takes precedence over everything.

## The Windows app

`TaskReporter.exe` is the whole program in one file: it serves the UI to
itself and draws it in a real window with a real title bar and a real taskbar
entry. Nothing else has to stay open — no terminal, no browser tab.

The window is exactly the page described in [The browser UI](#the-browser-ui),
so everything below about the Report and Board views, the shortcuts and the
history panel applies unchanged.

### Building it

Two steps, both from Windows, and neither needs WSL:

```
windows\build.bat              # produces TaskReporter.exe in this folder
windows\install_shortcut.bat   # Desktop + Start menu shortcuts
```

`build.bat` creates a private virtual environment in `.winenv\` the first time
it runs (openpyxl, pywebview and PyInstaller), then packages the same
`task-report-maker.py` the Linux launcher runs. The exe carries its own Python,
so nothing needs to be installed for it to run.

`.winenv\` and `build\` are build scratch — about 70 MB, git-ignored, and safe
to delete at any time; the next `build.bat` makes them again.

`windows\install_shortcut.bat /remove` takes the shortcuts away again.

To change the UI without repackaging every time, `windows\run_dev.bat` runs the
window straight from the source file using that same environment.

### Where the reports go

The same `task_reports.xlsx` and `task_board.json` as every other surface. The
exe finds them the way `task_reporter.bat` does — `%TASK_REPORT_DIR%` if it is
set, then the folder the exe is sitting in if that folder holds the data, then
the project path recorded when it was built. So a copy of the exe carried
somewhere else still files into the real workbook, and `%TASK_REPORT_DIR%`
overrides all of it if the project ever moves.

That path matters because a packaged app cannot use the folder its code is in:
PyInstaller unpacks itself into a temporary directory that is deleted on exit,
which is exactly where a workbook must never be written.

### One window, not several

Launching it again while it is already open raises the window that is there
rather than opening a second one. The running app records its address in
`.task_reporter_app.lock`, and the second launch hands the request over and
exits. If that file is ever left behind by a crash it simply fails to answer
and is cleared, so it can never block a start-up.

Nor does it share a port with a session running somewhere else. A session
started under WSL holds port 8770 inside the VM *and* has it mirrored onto
Windows by `wslrelay.exe`, and Windows will happily let a second process bind
the same port and then split incoming connections between the two. So the app
asks whether a port already answers before binding it, and moves to the next
one if it does — otherwise the window can end up loading a session that is not
its own, token and all.

### If it does not start

There is no terminal for it to complain in, so it says so in a message box and
writes everything to `.task_reporter_app.log` beside the workbook. That file is
also where the start-up address and any save errors go.

The one requirement is the **Edge WebView2 runtime**, which ships with Windows
10 and 11 and with Edge itself. `./task-report --doctor` reports its version
along with everything else, and links to the installer if it is missing.

### Why a window here and not under WSL

The window is drawn by WebView2 — an Edge process on Windows, placed by the
Windows window manager. That is the same reasoning as
[Why a browser and not a desktop window](#why-a-browser-and-not-a-desktop-window)
below, and not a reversal of it: what that section rules out is rendering
through WSLg, and neither the browser UI nor this window does. The Qt window
did, which is why it stays opt-in.

What the window does not have is the terminal console and its `:t` commands.
Everything they do is on the Board view; for scripting, the Linux launcher and
its flags are unchanged.

## The browser UI

The UI is a page served on loopback and opened in your normal browser. It has
two views, switched with the **Report / Board** tabs in the header or with
**Ctrl+B**:

* **Report** — type the report, press **Ctrl+Enter** (or click *Save Report*).
* **Board** — the task list. See [The task board](#the-task-board).

`☰` shows previous reports, where each one can be edited or deleted; `?` lists
the shortcuts. The number on the *Board* tab is how many ticked tasks are
waiting to be filed.

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

## The task board

The board is the running list of what has to be done, kept in the same shape it
is usually written by hand:

```
28.08.2026
[onex-academy]
  ☑ Picture zoom feature in the questions
  ☐ Paste the text using mouse right click
[ai-agents-scores]
  ☑ Editable ticket levels
```

A day holds any number of projects, and a project any number of tasks. Tasks
without a project are allowed — they file without a header.

**Ticking the boxes is the point.** Press **File Checked Tasks** and every
ticked task becomes report rows in `task_reports.xlsx`:

* **one row per day**, timestamped with *that day's* date — a backlog ticked off
  today still lands on the day the work happened, not on today;
* inside a row, the tasks are grouped under their `[project]` headers, in the
  order the projects were first written that day;
* a preview shows the exact text of every row before anything is written.

### How a filed row is laid out

Each task gets its own bulleted line under its project heading, with a blank
line before the next project, so a day covering several projects stays readable
instead of running together:

```
[onex-academy]
• Picture zoom feature in the questions
• Paste the text using mouse right click

[ai-agents-scores]
• Editable ticket levels
```

A task that is itself several lines long keeps them, indented under its own
bullet. Tasks with no project are listed first, without a heading.

The Task Report column is written with **wrap text** and top alignment, because
without it Excel draws a cell containing line breaks as one long run-together
line — which is what made a multi-project report look concatenated even though
the line breaks were there all along. Existing column widths in the workbook are
left as they are.

Filed tasks stay on the board struck through and marked `filed`, so nothing is
ever filed twice — their boxes are locked for the same reason. **Clear Filed**
drops them off the board once you no longer want to look at them; the reports
they were written into are untouched.

The day and project boxes keep their values after adding a task, because tasks
arrive in runs under one project.

### The project box

Click it (or the `▾`) and it drops down every project you have used, most recent
first. Pick one with the mouse, or with **↑ ↓** and **Enter**. Typing narrows the
list.

**A name that is not in the list is a new project** — type it, press **Enter**,
and that is the whole of creating one. Enter in this box never adds the task; it
accepts the project and moves on to the task text.

The list is remembered separately from the tasks, so a project stays on offer
after its tasks have been filed and cleared. Reusing the same name matters
because that exact spelling becomes the `[bracketed]` header in the report —
which is why picking from the list beats retyping. To drop a name you typed by
mistake, hover it in the dropdown and click the `✕`; tasks already using it keep
it, and using the name again brings it back.

### From the terminal

The same board, the same file. `:t` on its own lists it, and the numbers it
prints are what the other commands take.

| Command | Effect |
| --- | --- |
| `:t` / `:tasks` | show the board |
| `:t all` | show it including tasks already filed |
| `:t add [proj] TEXT` | add a task to today |
| `:t add 21.08.2026 [proj] TEXT` | add a task to another day |
| `:t x N...` | tick tasks by number (`:t done N` too) |
| `:t o N...` | untick tasks by number (`:t open N` too) |
| `:t rm N...` | delete tasks by number |
| `:t file` | write every ticked task into the workbook (asks first) |
| `:t clear` | drop already-filed tasks off the board |
| `:t projects` | list the remembered project names |
| `:t forget NAME` | drop a project name off that list |

Both prefixes on `:t add` are optional and in that order — a date, then the
project in brackets, then the task. The same grammar works from a shell:

```bash
./task-report -t "[onex-academy] fix the links comparison issue"
./task-report -t "21.08.2026 [onex-academy] fix the links comparison issue"
./task-report --board
./task-report --file-checked     # no prompt: this one is for scripts
```

A task added from a second shell while the app is running shows up in the open
browser page on its own, within a few seconds.

### Where it is kept

`task_board.json`, next to the workbook. It is deliberately *not* a second
worksheet: `task_reports.xlsx` cannot be written while Excel has it open, and
ticking a box must never be the thing that fails. Reports are the archive; the
board is the working surface in front of it.

Saves are atomic, so a crash cannot leave a half-written board. If the file ever
does become unreadable, it is copied aside to `task_board.json.bad` and a fresh
board is started rather than the app refusing to run.

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
| `:t` / `:tasks` | the task board — see [From the terminal](#from-the-terminal) |
| `:p` / `:path` | print the workbook path |
| `:h` / `:help` | show the command list |
| `:q` / `:quit` | close the session (Ctrl+C and Ctrl+D do the same) |

## If the workbook is open in Excel

While Excel has `task_reports.xlsx` open, nothing is written to it. Reports are
queued in `.task_reports_pending.jsonl` instead and merged in automatically once
the file is free. Editing or deleting a report is refused outright with a "close
it and try again" message, because those rewrite the whole sheet.

**Why it refuses rather than trying.** On a Windows drive reached through WSL —
which is where this project normally lives — Excel's lock does not reach Python
as an error. The write *succeeds*, and then Excel saves its own older in-memory
copy over the top a minute later and the row is simply gone, with nothing in the
app able to notice. So the app checks for the owner file Excel keeps beside an
open workbook (`~$task_reports.xlsx`) and queues on that signal instead of
finding out the hard way.

Both the filing preview and the terminal say so before you commit, `--doctor`
prints `open in Excel now`, and the console banner and `--doctor` both show how
many reports are still queued.

Filing from the board goes through the same queue, and the tasks are still
marked filed — a queued row is a kept row, so re-filing would only duplicate it.
The board itself is a separate file and is never blocked by Excel.

Note that editing or deleting a report rewrites the sheet, which also rewrites
any Date-Time cell that Excel had stored as a real date into the same timestamp
as text. The reports themselves are untouched.
