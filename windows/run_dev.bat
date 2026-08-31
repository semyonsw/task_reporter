@echo off
:: Run the app straight from source with the build venv - no packaging step.
:: Useful when changing the UI: edit task-report-maker.py, run this.
setlocal EnableExtensions
set "PROJ=%~dp0.."
"%PROJ%\.winenv\Scripts\pythonw.exe" "%PROJ%\task-report-maker.py" --app %*
