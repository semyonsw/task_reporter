@echo off
rem ===========================================================================
rem  Task Reporter - one-click installer
rem
rem  Double-click this file.  It installs everything the app needs and puts a
rem  shortcut on your Desktop.  Safe to run again at any time.
rem
rem  All it does is call tools\install.ps1, which writes install.log next to
rem  this file.
rem ===========================================================================
setlocal EnableExtensions
title Install Task Reporter
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"

if not exist "%~dp0tools\install.ps1" (
    echo.
    echo  [ERROR] tools\install.ps1 is missing next to this file.
    echo.
    echo  This download is incomplete. Get the whole project again:
    echo      git clone https://github.com/semyonsw/task_reporter.git
    echo  or download the ZIP from GitHub and extract ALL of it.
    echo.
    pause
    exit /b 1
)

"%PS%" -NoProfile -NoLogo -ExecutionPolicy Bypass -File "%~dp0tools\install.ps1" %*
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    if "%RC%"=="9009" (
        echo.
        echo  [ERROR] Windows PowerShell could not be started, so the installer
        echo          could not run at all.
        echo.
        echo  Fix: press Start, type "powershell", right-click Windows PowerShell
        echo       and choose "Run as administrator", then run this command:
        echo.
        echo       cd /d "%~dp0" ^&^& powershell -ExecutionPolicy Bypass -File tools\install.ps1
        echo.
    )
)

echo.
echo  Press any key to close this window...
pause >nul
exit /b %RC%
