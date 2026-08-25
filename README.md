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

Run `./task-report --doctor`. It reports which Qt platform plugin works and
why the others fail.

The window needs a live display. On WSL that means WSLg must be running — if
it is not, `wsl.exe --shutdown` from Windows and reopening usually brings it
back. Either way the terminal console takes over automatically, so a missing
display never blocks a report.

## If the workbook is locked

Saving fails while `task_reports.xlsx` is open in Excel. Rather than losing the
report, it is queued in `.task_reports_pending.jsonl` and merged into the
workbook automatically on the next successful save. `--doctor` and the console
banner both show how many reports are still queued.
