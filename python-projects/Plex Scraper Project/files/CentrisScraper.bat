@echo off
:: ============================================================
::  Centris Scraper — One-Click Installer & Launcher
::  
::  First run: installs everything silently, then launches app
::  After that: just launches the app instantly
:: ============================================================
setlocal EnableDelayedExpansion

set APP_DIR=%LOCALAPPDATA%\CentrisScraper
set PYTHON_EMBED=https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
set PIP_URL=https://bootstrap.pypa.io/get-pip.py
set READY_FLAG=%APP_DIR%\.ready

:: Already installed — just launch
if exist "%READY_FLAG%" goto LAUNCH

:: ── FIRST-TIME SETUP ─────────────────────────────────────────
echo.
echo  ==============================================
echo    Centris Scraper - First-Time Setup
echo    This takes about 3-5 minutes. Please wait.
echo  ==============================================
echo.

mkdir "%APP_DIR%" 2>nul

:: Download embedded Python (no system Python needed)
echo  [1/4] Downloading Python...
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -Uri '%PYTHON_EMBED%' -OutFile '%APP_DIR%\python.zip'}"
powershell -Command "Expand-Archive -Path '%APP_DIR%\python.zip' -DestinationPath '%APP_DIR%\python' -Force"
del "%APP_DIR%\python.zip"

:: Enable pip in embedded Python
echo import site >> "%APP_DIR%\python\python311._pth"

:: Install pip
echo  [2/4] Installing pip...
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -Uri '%PIP_URL%' -OutFile '%APP_DIR%\get-pip.py'}"
"%APP_DIR%\python\python.exe" "%APP_DIR%\get-pip.py" --quiet
del "%APP_DIR%\get-pip.py"

:: Install packages
echo  [3/4] Installing packages (playwright, openpyxl)...
"%APP_DIR%\python\Scripts\pip.exe" install playwright openpyxl --quiet --no-warn-script-location

:: Install Chromium
echo  [4/4] Installing Chromium browser...
set PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\browsers
"%APP_DIR%\python\Scripts\playwright.exe" install chromium

:: Copy app files
copy /Y "%~dp0centris_app.py" "%APP_DIR%\" >nul
copy /Y "%~dp0centris_gui.py" "%APP_DIR%\" >nul

:: Mark as ready
echo ready > "%READY_FLAG%"

echo.
echo  ✓ Setup complete! Launching app...
echo.

:LAUNCH
set PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\browsers
"%APP_DIR%\python\python.exe" "%APP_DIR%\centris_gui.py"
endlocal
