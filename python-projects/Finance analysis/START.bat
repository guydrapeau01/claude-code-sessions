@echo off
cd /d "%~dp0"

echo === DCF Dashboard Launcher ===
echo Working folder: %~dp0
echo.

echo Installing dependencies...
py -m pip install flask requests --quiet
if errorlevel 1 (
    python -m pip install flask requests --quiet
)
echo.

echo Starting DCF Dashboard...
echo Open your browser at: http://localhost:5050
echo.
py dcf_scraper_app.py
if errorlevel 1 (
    python dcf_scraper_app.py
)

echo.
echo === Server stopped. Press any key to close. ===
pause
