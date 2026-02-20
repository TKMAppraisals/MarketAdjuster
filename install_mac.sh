#!/bin/bash
# MarketAdjuster - macOS Install/Update Script
#
# USAGE: Paste this ONE command in Terminal:
#
#   curl -sL https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_mac.sh | bash
#

set -e

APP_NAME="MarketAdjuster"
INSTALL_DIR="$HOME/Applications/MarketAdjuster"
VENV_DIR="$INSTALL_DIR/venv"
APP_DIR="$INSTALL_DIR/app"
REPO_BASE="https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main"

echo ""
echo "  ============================================"
echo "   MarketAdjuster - Install / Update"
echo "  ============================================"
echo ""

if [ -f "$APP_DIR/app.py" ]; then
    echo "  Updating existing installation..."
    MODE="UPDATE"
else
    echo "  Fresh installation..."
    MODE="INSTALL"
fi

# ---- Check for Python 3 ----
echo ""
echo "[1/6] Checking for Python 3..."
if command -v python3 &>/dev/null; then
    PYTHON_CMD="python3"
    echo "  Found: $($PYTHON_CMD --version 2>&1)"
else
    echo ""
    echo "  Python 3 is not installed."
    echo "  Opening download page..."
    open "https://www.python.org/downloads/"
    echo '  Install Python, then run this script again.'
    exit 1
fi

# ---- Create directories ----
echo ""
echo "[2/6] Setting up folders..."
mkdir -p "$INSTALL_DIR"
mkdir -p "$APP_DIR"

# ---- Download latest files ----
echo ""
echo "[3/6] Downloading latest app files..."
curl -sL "$REPO_BASE/market_condition_app_v4_15_premium_plus.py" -o "$APP_DIR/app.py"
curl -sL "$REPO_BASE/MarketAdjuster_macOS_512.png" -o "$APP_DIR/MarketAdjuster_macOS_512.png"
curl -sL "$REPO_BASE/requirements.txt" -o "$APP_DIR/requirements.txt"
curl -sL "$REPO_BASE/app_icon.png" -o "$APP_DIR/app_icon.png"

# ---- Virtual environment ----
echo ""
if [ "$MODE" = "UPDATE" ] && [ -d "$VENV_DIR" ]; then
    echo "[4/6] Reusing existing environment..."
else
    echo "[4/6] Creating Python environment..."
    $PYTHON_CMD -m venv "$VENV_DIR"
fi

# ---- Install dependencies ----
echo ""
echo "[5/6] Installing dependencies..."
source "$VENV_DIR/bin/activate"
pip install --upgrade pip --quiet 2>/dev/null
pip install -r "$APP_DIR/requirements.txt" --quiet

# ---- Create .app bundle ----
echo ""
echo "[6/6] Creating application..."

APP_BUNDLE="$INSTALL_DIR/MarketAdjuster.app"
mkdir -p "$APP_BUNDLE/Contents/MacOS"
mkdir -p "$APP_BUNDLE/Contents/Resources"

cp "$APP_DIR/app_icon.png" "$APP_BUNDLE/Contents/Resources/app_icon.png" 2>/dev/null

cat > "$APP_BUNDLE/Contents/Info.plist" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>MarketAdjuster</string>
    <key>CFBundleName</key>
    <string>MarketAdjuster</string>
    <key>CFBundleIdentifier</key>
    <string>com.tkm.marketadjuster</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleIconFile</key>
    <string>app_icon</string>
    <key>LSUIElement</key>
    <true/>
</dict>
</plist>
PLIST

cat > "$APP_BUNDLE/Contents/MacOS/MarketAdjuster" << 'EXEC'
#!/bin/bash

INSTALL_DIR="$HOME/Applications/MarketAdjuster"
VENV_DIR="$INSTALL_DIR/venv"
APP_DIR="$INSTALL_DIR/app"
LOG_FILE="$INSTALL_DIR/launch.log"

export PATH="/usr/local/bin:/usr/bin:/bin:/opt/homebrew/bin:$HOME/.local/bin:$PATH"

if [ -f "$VENV_DIR/bin/activate" ]; then
    source "$VENV_DIR/bin/activate"
else
    osascript -e 'display dialog "MarketAdjuster environment not found. Please reinstall." buttons {"OK"} default button "OK" with icon caution with title "MarketAdjuster"'
    exit 1
fi

cd "$APP_DIR"

if ! command -v streamlit &>/dev/null; then
    osascript -e 'display dialog "Streamlit not found. Please reinstall MarketAdjuster." buttons {"OK"} default button "OK" with icon caution with title "MarketAdjuster"'
    exit 1
fi

if lsof -i :8501 &>/dev/null; then
    open "http://localhost:8501"
    exit 0
fi

streamlit run app.py --server.headless true --browser.gatherUsageStats false > "$LOG_FILE" 2>&1 &
STREAMLIT_PID=$!

sleep 4
if kill -0 $STREAMLIT_PID 2>/dev/null; then
    open "http://localhost:8501"
    wait $STREAMLIT_PID
else
    osascript -e 'display dialog "MarketAdjuster failed to start. Check ~/Applications/MarketAdjuster/launch.log" buttons {"OK"} default button "OK" with icon caution with title "MarketAdjuster"'
    exit 1
fi
EXEC
chmod +x "$APP_BUNDLE/Contents/MacOS/MarketAdjuster"

cat > "$INSTALL_DIR/launch.sh" << 'LAUNCHER'
#!/bin/bash
INSTALL_DIR="$HOME/Applications/MarketAdjuster"
export PATH="/usr/local/bin:/usr/bin:/bin:/opt/homebrew/bin:$HOME/.local/bin:$PATH"
source "$INSTALL_DIR/venv/bin/activate"
cd "$INSTALL_DIR/app"
if lsof -i :8501 &>/dev/null; then
    open "http://localhost:8501"
else
    streamlit run app.py --server.headless true --browser.gatherUsageStats false &
    sleep 4
    open "http://localhost:8501"
    wait
fi
LAUNCHER
chmod +x "$INSTALL_DIR/launch.sh"

# Convert PNG to icns if possible
if command -v sips &>/dev/null && [ -f "$APP_DIR/app_icon.png" ]; then
    ICONSET_DIR=$(mktemp -d)/MA.iconset
    mkdir -p "$ICONSET_DIR"
    for s in 16 32 64 128 256 512; do
        sips -z $s $s "$APP_DIR/app_icon.png" --out "$ICONSET_DIR/icon_${s}x${s}.png" 2>/dev/null
    done
    cp "$APP_DIR/app_icon.png" "$ICONSET_DIR/icon_512x512@2x.png" 2>/dev/null
    sips -z 32 32 "$APP_DIR/app_icon.png" --out "$ICONSET_DIR/icon_16x16@2x.png" 2>/dev/null
    sips -z 64 64 "$APP_DIR/app_icon.png" --out "$ICONSET_DIR/icon_32x32@2x.png" 2>/dev/null
    sips -z 256 256 "$APP_DIR/app_icon.png" --out "$ICONSET_DIR/icon_128x128@2x.png" 2>/dev/null
    sips -z 512 512 "$APP_DIR/app_icon.png" --out "$ICONSET_DIR/icon_256x256@2x.png" 2>/dev/null
    iconutil -c icns "$ICONSET_DIR" -o "$APP_BUNDLE/Contents/Resources/app_icon.icns" 2>/dev/null
fi

echo ""
if [ "$MODE" = "UPDATE" ]; then
    echo "  ============================================"
    echo "   Update Complete!"
    echo "  ============================================"
    echo "  Report history preserved."
else
    echo "  ============================================"
    echo "   Installation Complete!"
    echo "  ============================================"
fi
echo ""
echo "  A Finder window will open with the MarketAdjuster app."
echo "  Drag it to your Desktop, Dock, or wherever you'd like."
echo ""

# Open Finder with the .app selected so user can drag it where they want
open -R "$APP_BUNDLE"
