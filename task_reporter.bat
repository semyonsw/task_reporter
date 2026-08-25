@echo off
setlocal EnableExtensions
TITLE Task Report Maker

:: --------------------------------------------------------------------------
:: If you move the project, update this one line. It is only used when this
:: .bat has been COPIED away from the project (e.g. onto the Desktop).
:: --------------------------------------------------------------------------
set "FALLBACKDIR=C:\Users\User\Desktop\semyONEX\semyon-not-work-lol\AI_ML_programmming\self_created_projects\task-report"

:: --------------------------------------------------------------------------
:: Find the project, in order of preference:
::   1. %TASK_REPORT_DIR%   - set this to override everything
::   2. the folder holding this .bat   - the normal case, and what makes a
::      Desktop SHORTCUT work (a shortcut runs the original file in place)
::   3. FALLBACKDIR above   - so a plain COPY of this file still works
:: --------------------------------------------------------------------------
set "PROJDIR="
if defined TASK_REPORT_DIR if exist "%TASK_REPORT_DIR%\task-report-maker.py" set "PROJDIR=%TASK_REPORT_DIR%"
if not defined PROJDIR if exist "%~dp0task-report-maker.py" set "PROJDIR=%~dp0"
if not defined PROJDIR if exist "%FALLBACKDIR%\task-report-maker.py" set "PROJDIR=%FALLBACKDIR%"

if not defined PROJDIR (
    echo [ERROR] Could not find task-report-maker.py.
    echo         Looked next to this file:
    echo           %~dp0
    echo         and at the recorded project path:
    echo           %FALLBACKDIR%
    echo.
    echo         Fix: edit FALLBACKDIR near the top of this file, or set
    echo         TASK_REPORT_DIR to the project folder.
    echo.
    pause
    exit /b 1
)

:: Drop a trailing backslash so the wslpath argument is always well formed.
if "%PROJDIR:~-1%"=="\" set "PROJDIR=%PROJDIR:~0,-1%"

:: --------------------------------------------------------------------------
:: python is called directly rather than through "conda run": conda run
:: buffers stdout and does not pass stdin through, which would break the
:: terminal console that runs alongside the window.
:: --------------------------------------------------------------------------
set "WSLDIR="
for /f "usebackq delims=" %%i in (`wsl -u root --exec wslpath -a "%PROJDIR%" 2^>nul`) do set "WSLDIR=%%i"

if not defined WSLDIR (
    echo [ERROR] Could not translate this folder into a WSL path:
    echo         %PROJDIR%
    echo         Check that WSL is installed and running ^(try: wsl --status^).
    echo.
    pause
    exit /b 1
)

echo Launching Task Reporter...
echo   Folder : %WSLDIR%
echo   User   : root
echo.
echo The window and this terminal both accept reports.
echo Closing either one closes the other.
echo.

wsl -u root -- bash -lc "cd '%WSLDIR%' && exec bash ./task-report %*"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [ERROR] Task Reporter exited with code %RC%.
    echo For a diagnosis of why the window did not open, run:
    echo   wsl -u root -- bash -lc "cd '%WSLDIR%' && bash ./task-report --doctor"
    echo.
    echo Press any key to close...
    pause >nul
)

exit /b %RC%
