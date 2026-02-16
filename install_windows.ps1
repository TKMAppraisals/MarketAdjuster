# MarketAdjuster - Windows Install/Update Script
# 
# USAGE: Your tester opens PowerShell and pastes this ONE command:
#
#   irm https://raw.githubusercontent.com/YOUR_USERNAME/MarketAdjuster/main/install_windows.ps1 | iex
#
# Or if not using GitHub, they can just double-click this file:
#   Right-click > Run with PowerShell
#

$ErrorActionPreference = "Stop"
$APP_NAME = "MarketAdjuster"
$INSTALL_DIR = "$env:LOCALAPPDATA\MarketAdjuster"
$VENV_DIR = "$INSTALL_DIR\venv"
$APP_DIR = "$INSTALL_DIR\app"
$DESKTOP = [Environment]::GetFolderPath("Desktop")

# GitHub raw URL - UPDATE THIS to your repo
$REPO_BASE = "https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main"

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "   MarketAdjuster - Install / Update" -ForegroundColor Cyan  
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

# Detect install vs update
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

# Try py launcher
try { 
    $ver = & py -3 --version 2>&1
    if ($LASTEXITCODE -eq 0) { $PYTHON_CMD = "py"; $PYTHON_ARGS = @("-3"); Write-Host "  Found: $ver" -ForegroundColor Green }
} catch {}

# Try python3
if (-not $PYTHON_CMD) {
    try {
        $ver = & python3 --version 2>&1
        if ($LASTEXITCODE -eq 0) { $PYTHON_CMD = "python3"; Write-Host "  Found: $ver" -ForegroundColor Green }
    } catch {}
}

# Try python (verify it's real)
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
    @{url="$REPO_BASE/market_condition_app_v4_15_premium_plus.py"; dest="$APP_DIR\_source.py"},
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

# ---- Compile to bytecode and remove source ----
Write-Host "  Compiling application..." -ForegroundColor White
& $PYTHON_CMD @PYTHON_ARGS -c "import py_compile,os; py_compile.compile(r'$APP_DIR\_source.py',cfile=r'$APP_DIR\app.pyc',optimize=2); os.remove(r'$APP_DIR\_source.py')"

# Create thin launcher
@"
import importlib.util, sys, os
_d = os.path.dirname(os.path.abspath(__file__))
_p = os.path.join(_d, "app.pyc")
if not os.path.exists(_p):
    for _f in os.listdir(_d):
        if _f.endswith('.pyc'):
            _p = os.path.join(_d, _f)
            break
if not os.path.exists(_p):
    raise SystemExit("Application files not found. Please reinstall.")
_s = importlib.util.spec_from_file_location("__mp__", _p)
_m = importlib.util.module_from_spec(_s)
_s.loader.exec_module(_m)
"@ | Out-File -FilePath "$APP_DIR\app.py" -Encoding UTF8

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

@"
@echo off
title MarketAdjuster
cd /d "$APP_DIR"
call "$VENV_DIR\Scripts\activate.bat"
"$VENV_DIR\Scripts\streamlit.exe" run app.py --server.headless true --browser.gatherUsageStats false
"@ | Out-File -FilePath "$INSTALL_DIR\Launch_MarketAdjuster.bat" -Encoding ASCII

@"
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "$INSTALL_DIR\Launch_MarketAdjuster.bat" & chr(34), 0
Set WshShell = Nothing
"@ | Out-File -FilePath "$INSTALL_DIR\Launch_MarketAdjuster.vbs" -Encoding ASCII

# Create desktop shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$DESKTOP\MarketAdjuster.lnk")
$Shortcut.TargetPath = "$INSTALL_DIR\Launch_MarketAdjuster.vbs"
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
