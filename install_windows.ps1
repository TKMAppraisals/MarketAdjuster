# MarketAdjuster - Windows Install/Update Script
#
# USAGE: Open PowerShell and paste:
#
#   irm https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_windows.ps1 | iex
#

$ErrorActionPreference = "Stop"
$APP_NAME = "MarketAdjuster"
$INSTALL_DIR = "$env:LOCALAPPDATA\MarketAdjuster"
$VENV_DIR = "$INSTALL_DIR\venv"
$APP_DIR = "$INSTALL_DIR\app"
$DESKTOP = [Environment]::GetFolderPath("Desktop")

$REPO_BASE = "https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main"

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "   MarketAdjuster - Install / Update" -ForegroundColor Cyan  
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path "$APP_DIR\app.py") {
    Write-Host "  Updating existing installation..." -ForegroundColor Yellow
    $MODE = "UPDATE"
} else {
    Write-Host "  Fresh installation..." -ForegroundColor Green
    $MODE = "INSTALL"
}

# ---- Find Python ----
Write-Host ""
Write-Host "[1/5] Checking for Python..." -ForegroundColor White
$PYTHON_CMD = $null
$PYTHON_ARGS = @()

try { 
    $ver = & py -3 --version 2>&1
    if ($LASTEXITCODE -eq 0) { $PYTHON_CMD = "py"; $PYTHON_ARGS = @("-3"); Write-Host "  Found: $ver" -ForegroundColor Green }
} catch {}

if (-not $PYTHON_CMD) {
    try {
        $ver = & python3 --version 2>&1
        if ($LASTEXITCODE -eq 0) { $PYTHON_CMD = "python3"; Write-Host "  Found: $ver" -ForegroundColor Green }
    } catch {}
}

if (-not $PYTHON_CMD) {
    try {
        & python -c "import sys" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { 
            $ver = & python --version 2>&1
            $PYTHON_CMD = "python"; Write-Host "  Found: $ver" -ForegroundColor Green 
        }
    } catch {}
}

if (-not $PYTHON_CMD) {
    Write-Host ""
    Write-Host "  Python 3.10+ is required but not found." -ForegroundColor Red
    Write-Host ""
    Write-Host "  Opening Python download page..." -ForegroundColor Yellow
    Write-Host '  IMPORTANT: Check "Add Python to PATH" during install!' -ForegroundColor Yellow
    Write-Host "  Then run this script again." -ForegroundColor Yellow
    Write-Host ""
    Start-Process "https://www.python.org/downloads/"
    Read-Host "  Press Enter to exit"
    exit 1
}

# ---- Create directories ----
Write-Host ""
Write-Host "[2/5] Setting up folders..." -ForegroundColor White
New-Item -ItemType Directory -Force -Path $INSTALL_DIR | Out-Null
New-Item -ItemType Directory -Force -Path $APP_DIR | Out-Null

# ---- Download app files ----
Write-Host "  Downloading latest app files..." -ForegroundColor White

$files = @(
    @{url="$REPO_BASE/market_condition_app_v4_15_premium_plus.py"; dest="$APP_DIR\app.py"},
    @{url="$REPO_BASE/MarketAdjuster_macOS_512.png"; dest="$APP_DIR\MarketAdjuster_macOS_512.png"},
    @{url="$REPO_BASE/requirements.txt"; dest="$APP_DIR\requirements.txt"},
    @{url="$REPO_BASE/app_icon.ico"; dest="$INSTALL_DIR\app_icon.ico"}
)

foreach ($f in $files) {
    try {
        Invoke-WebRequest -Uri $f.url -OutFile $f.dest -UseBasicParsing
    } catch {
        $localFile = Join-Path $PSScriptRoot (Split-Path $f.url -Leaf)
        if (Test-Path $localFile) {
            Copy-Item $localFile $f.dest -Force
        } else {
            Write-Host "  Warning: Could not download $(Split-Path $f.url -Leaf)" -ForegroundColor Yellow
        }
    }
}

# ---- Create venv ----
Write-Host ""
if ($MODE -eq "UPDATE" -and (Test-Path "$VENV_DIR\Scripts\activate.bat")) {
    Write-Host "[3/5] Reusing existing environment..." -ForegroundColor White
} else {
    Write-Host "[3/5] Creating Python environment..." -ForegroundColor White
    & $PYTHON_CMD @PYTHON_ARGS -m venv $VENV_DIR
}

# ---- Install dependencies ----
Write-Host ""
Write-Host "[4/5] Installing dependencies (may take 2-3 min)..." -ForegroundColor White
& "$VENV_DIR\Scripts\python.exe" -m pip install --upgrade pip --quiet 2>&1 | Out-Null
& "$VENV_DIR\Scripts\python.exe" -m pip install -r "$APP_DIR\requirements.txt" --quiet

# ---- Create launcher ----
Write-Host ""
Write-Host "[5/5] Creating desktop shortcut..." -ForegroundColor White

# PowerShell launcher script - runs completely hidden, opens browser automatically
@"
`$VENV = "$VENV_DIR"
`$APP = "$APP_DIR"

# Check if already running
`$listening = netstat -ano 2>`$null | Select-String ":8501.*LISTENING"
if (`$listening) {
    Start-Process "http://localhost:8501"
    exit
}

# Start streamlit completely hidden
`$proc = Start-Process -FilePath "`$VENV\Scripts\streamlit.exe" ``
    -ArgumentList "run", "app.py", "--server.headless", "true", "--browser.gatherUsageStats", "false", "--server.port", "8501" ``
    -WorkingDirectory `$APP ``
    -WindowStyle Hidden ``
    -PassThru

# Wait for streamlit to be ready, then open browser
`$timeout = 30
`$elapsed = 0
while (`$elapsed -lt `$timeout) {
    Start-Sleep -Seconds 1
    `$elapsed++
    try {
        `$tcp = New-Object System.Net.Sockets.TcpClient
        `$tcp.Connect("127.0.0.1", 8501)
        `$tcp.Close()
        # Server is ready
        Start-Process "http://localhost:8501"
        exit
    } catch {
        # Not ready yet
    }
}

# Fallback: open browser anyway after timeout
Start-Process "http://localhost:8501"
"@ | Out-File -FilePath "$INSTALL_DIR\Launch_MarketAdjuster.ps1" -Encoding UTF8

# VBS wrapper to run the PowerShell script completely hidden (no window at all)
@"
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$INSTALL_DIR\Launch_MarketAdjuster.ps1""", 0, False
Set objShell = Nothing
"@ | Out-File -FilePath "$INSTALL_DIR\Launch_MarketAdjuster.vbs" -Encoding ASCII

# Create desktop shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$DESKTOP\MarketAdjuster.lnk")
$Shortcut.TargetPath = "wscript.exe"
$Shortcut.Arguments = """$INSTALL_DIR\Launch_MarketAdjuster.vbs"""
$Shortcut.WorkingDirectory = $APP_DIR
$Shortcut.Description = "MarketAdjuster"
$Shortcut.IconLocation = "$INSTALL_DIR\app_icon.ico,0"
$Shortcut.Save()

Write-Host ""
if ($MODE -eq "UPDATE") {
    Write-Host "  ============================================" -ForegroundColor Green
    Write-Host "   Update Complete!" -ForegroundColor Green
    Write-Host "  ============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Report history preserved." -ForegroundColor White
} else {
    Write-Host "  ============================================" -ForegroundColor Green
    Write-Host "   Installation Complete!" -ForegroundColor Green
    Write-Host "  ============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  MarketAdjuster icon is on your Desktop." -ForegroundColor White
    Write-Host "  Double-click it to launch." -ForegroundColor White
}
Write-Host ""
Read-Host "  Press Enter to close"
