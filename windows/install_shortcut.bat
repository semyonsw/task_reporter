@echo off
:: Put Task Reporter on the Desktop and in the Start menu.
::
::   windows\install_shortcut.bat            create the shortcuts
::   windows\install_shortcut.bat /remove    take them away again
::
:: Shortcuts point at TaskReporter.exe where it lives, so the workbook and the
:: task board stay in the project folder no matter where the app is started.
setlocal EnableExtensions
set "HERE=%~dp0"
for %%I in ("%HERE%..") do set "PROJ=%%~fI"
set "TARGET=%PROJ%\TaskReporter.exe"
set "ICON=%HERE%TaskReporter.ico"

if /i "%~1"=="/remove" goto :remove

if not exist "%TARGET%" (
    echo [ERROR] %TARGET% does not exist yet.
    echo         Build it first:  windows\build.bat
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$shell = New-Object -ComObject WScript.Shell;" ^
  "foreach ($dir in @($shell.SpecialFolders('Desktop'), (Join-Path $shell.SpecialFolders('Programs') ''))) {" ^
  "  $link = $shell.CreateShortcut((Join-Path $dir 'Task Reporter.lnk'));" ^
  "  $link.TargetPath = '%TARGET%';" ^
  "  $link.WorkingDirectory = '%PROJ%';" ^
  "  $link.IconLocation = '%ICON%';" ^
  "  $link.Description = 'File task reports into task_reports.xlsx';" ^
  "  $link.Save();" ^
  "  Write-Host ('  created ' + (Join-Path $dir 'Task Reporter.lnk'))" ^
  "}"
if errorlevel 1 (
    echo [ERROR] Could not create the shortcuts.
    exit /b 1
)
echo.
echo Done - 'Task Reporter' is on the Desktop and in the Start menu.
exit /b 0

:remove
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$shell = New-Object -ComObject WScript.Shell;" ^
  "foreach ($dir in @($shell.SpecialFolders('Desktop'), $shell.SpecialFolders('Programs'))) {" ^
  "  $path = Join-Path $dir 'Task Reporter.lnk';" ^
  "  if (Test-Path $path) { Remove-Item $path; Write-Host ('  removed ' + $path) }" ^
  "}"
echo.
echo Shortcuts removed. TaskReporter.exe itself is untouched.
exit /b 0
