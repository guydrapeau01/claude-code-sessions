#!/bin/bash
# ============================================================
#  Centris Scraper — One-Click Installer & Launcher (Mac)
#
#  First run: installs everything, then launches app
#  After that: just launches the app instantly
# ============================================================

APP_DIR="$HOME/.centris_scraper"
READY_FLAG="$APP_DIR/.ready"

# Already installed — just launch
if [ -f "$READY_FLAG" ]; then
    export PLAYWRIGHT_BROWSERS_PATH="$APP_DIR/browsers"
    python3 "$APP_DIR/centris_gui.py"
    exit 0
fi

# ── FIRST-TIME SETUP ─────────────────────────────────────────
echo ""
echo " =============================================="
echo "   Centris Scraper - First-Time Setup"
echo "   This takes about 3-5 minutes. Please wait."
echo " =============================================="
echo ""

mkdir -p "$APP_DIR"

# Check Python
if ! command -v python3 &>/dev/null; then
    osascript -e 'display dialog "Python 3 is required.\n\nInstall from https://python.org then re-run this script." buttons {"OK"} with icon stop'
    exit 1
fi

echo " [1/3] Installing packages..."
pip3 install playwright openpyxl --quiet

echo " [2/3] Installing Chromium..."
export PLAYWRIGHT_BROWSERS_PATH="$APP_DIR/browsers"
playwright install chromium

echo " [3/3] Copying app files..."
cp "$(dirname "$0")/centris_app.py" "$APP_DIR/"
cp "$(dirname "$0")/centris_gui.py" "$APP_DIR/"

touch "$READY_FLAG"

echo ""
echo " ✓ Setup complete! Launching..."
echo ""

python3 "$APP_DIR/centris_gui.py"
