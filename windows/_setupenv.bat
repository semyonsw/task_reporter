@echo off
rem ---------------------------------------------------------------------------
rem  Creates .winenv\ - the private build environment for TaskReporter.exe -
rem  and installs what the build needs into it.
rem
rem  You do not normally run this: Install.bat does the same thing, with proper
rem  error messages and a Python search that does not assume anything about
rem  where Python lives.  This file is kept for  windows\build.bat.
rem ---------------------------------------------------------------------------
setlocal EnableExtensions
set "PROJ=%~dp0.."

rem Find any Python that can create a venv. PY_FOR_BUILD wins if you set it.
set "BASEPY="
if defined PY_FOR_BUILD if exist "%PY_FOR_BUILD%" set "BASEPY=%PY_FOR_BUILD%"
if not defined BASEPY (
    for %%C in (
        "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
        "%USERPROFILE%\miniconda3\python.exe"
        "%USERPROFILE%\anaconda3\python.exe"
    ) do if not defined BASEPY if exist %%C set "BASEPY=%%~C"
)
if not defined BASEPY (
    where py >nul 2>&1 && set "BASEPY=py"
)
if not defined BASEPY (
    where python >nul 2>&1 && set "BASEPY=python"
)

if not defined BASEPY (
    echo [ERROR] No Python found to build with.
    echo         Install Python 3.12 from https://www.python.org/downloads/windows/
    echo         ^(tick "Add python.exe to PATH"^), or just run Install.bat, which
    echo         can fetch it for you.
    exit /b 1
)

if not exist "%PROJ%\.winenv\Scripts\python.exe" (
    echo Creating build venv with %BASEPY% ...
    "%BASEPY%" -m venv "%PROJ%\.winenv" || (
        echo [ERROR] Could not create %PROJ%\.winenv
        echo         Try running Install.bat instead - it says exactly what is wrong.
        exit /b 1
    )
)

"%PROJ%\.winenv\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel || exit /b 1
"%PROJ%\.winenv\Scripts\python.exe" -m pip install openpyxl pywebview pyinstaller || (
    echo [ERROR] The build dependencies could not be installed.
    echo         Run Install.bat - it retries, diagnoses and explains failures.
    exit /b 1
)
"%PROJ%\.winenv\Scripts\python.exe" -c "import openpyxl,webview,PyInstaller;print('deps ok')"
