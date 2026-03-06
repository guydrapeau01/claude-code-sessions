@echo off
setlocal

set SCRIPT_DIR=%~dp0
set APP_DIR=%LOCALAPPDATA%\CentrisScraper
set LOG=%SCRIPT_DIR%launch_log.txt
set READY=%APP_DIR%\.ready

echo === %date% %time% === > "%LOG%"
echo Script dir: %SCRIPT_DIR% >> "%LOG%"

:: Check Python
python --version >> "%LOG%" 2>&1
if errorlevel 1 (
    start notepad "%LOG%"
    msg * "Python not found. Install from python.org and tick 'Add to PATH'."
    exit /b 1
)

:: First-time setup
if exist "%READY%" goto LAUNCH

echo First time setup... >> "%LOG%"
mkdir "%APP_DIR%" 2>nul

python -m pip install playwright openpyxl >> "%LOG%" 2>&1
if errorlevel 1 ( start notepad "%LOG%" & exit /b 1 )

set PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\browsers
python -m playwright install chromium >> "%LOG%" 2>&1
if errorlevel 1 ( start notepad "%LOG%" & exit /b 1 )

echo ready > "%READY%"

:LAUNCH
echo Launching... >> "%LOG%"
set PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\browsers

:: Run directly from same folder as .bat - no copying needed
python "%SCRIPT_DIR%centris_gui.py" >> "%LOG%" 2>&1
if errorlevel 1 ( start notepad "%LOG%" )

endlocal
