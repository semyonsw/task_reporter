@echo off
:: Build TaskReporter.exe.  Run this from Windows; it needs no WSL.
::
::   windows\build.bat
::
:: The result is TaskReporter.exe in the project folder.  Run
:: windows\install_shortcut.bat afterwards to put it on the Desktop.
setlocal EnableExtensions
set "HERE=%~dp0"
set "PROJ=%HERE%.."
pushd "%PROJ%" || exit /b 1
for %%I in ("%CD%") do set "PROJ=%%~fI"

call "%HERE%_setupenv.bat" || goto :failed

:: Recorded into the bundle so a copied exe can still find this folder.
> "%HERE%project_home.txt" echo %PROJ%

echo.
echo Building TaskReporter.exe ...
"%PROJ%\.winenv\Scripts\pyinstaller.exe" ^
    --noconfirm --clean ^
    --distpath "%PROJ%\build\dist" ^
    --workpath "%PROJ%\build\work" ^
    "%HERE%TaskReporter.spec" || goto :failed

copy /Y "%PROJ%\build\dist\TaskReporter.exe" "%PROJ%\TaskReporter.exe" >nul || goto :failed

echo.
echo   Built: %PROJ%\TaskReporter.exe
echo   Next : windows\install_shortcut.bat   (Desktop + Start menu shortcuts)
popd
exit /b 0

:failed
echo.
echo [ERROR] Build failed.
popd
exit /b 1
