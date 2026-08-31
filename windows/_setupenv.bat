@echo off
setlocal EnableExtensions
set "PROJ=%~dp0.."
set "R=C:\Users\User\miniconda3"
set "M=%R%\envs\movs"
set "PATH=%M%;%M%\Library\mingw-w64\bin;%M%\Library\usr\bin;%M%\Library\bin;%M%\Scripts;%M%\bin;%PATH%"
if not exist "%PROJ%\.winenv\Scripts\python.exe" (
    echo Creating build venv...
    "%M%\python.exe" -m venv "%PROJ%\.winenv" || exit /b 1
)
"%PROJ%\.winenv\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel || exit /b 1
"%PROJ%\.winenv\Scripts\python.exe" -m pip install openpyxl pywebview pyinstaller || exit /b 1
"%PROJ%\.winenv\Scripts\python.exe" -c "import openpyxl,webview,PyInstaller;print('deps ok')"
