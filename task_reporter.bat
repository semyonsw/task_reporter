@echo off
TITLE Task Report Maker

:: ECHO steps for visibility
echo Launching WSL...
echo Target Path: /home/semyon-work/task-report
echo User: root
echo Environment: movs

:: --------------------------------------------------------------------------
:: COMMAND EXPLANATION:
:: wsl -u root: Logs in as root
:: bash -ic "...": Runs bash in interactive mode (loads conda from .bashrc)
:: "cd ... && ...": Changes folder FIRST, then runs python inside conda env
:: --------------------------------------------------------------------------

wsl -u root bash -ic "export DISPLAY=:0; cd /home/semyon-work/task-report && conda run -n movs python task-report-maker.py"

:: Check if the previous command failed
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] The script failed to run. 
    echo If 'movs' is a virtualenv instead of conda, you may need to edit this file.
    echo.
    echo Press any key to see the error details before closing...
    pause
)

:: If successful, the script reaches here and exits automatically.
