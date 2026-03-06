@echo off
echo.
echo ========================================
echo   CENTRIS PLEX SCRAPER
echo ========================================
echo.

echo Installing dependencies...
py -m pip install playwright openpyxl --quiet
py -m playwright install chromium --only-if-needed

echo.
echo Starting scraper...
echo.
py centris_app.py

pause
