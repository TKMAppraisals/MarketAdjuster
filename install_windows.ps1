# Market Adjuster – Windows Installer

Write-Host "Starting Market Adjuster installation..." -ForegroundColor Cyan

# --- Check for Python ---
$pythonCmd = $null

if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = "py -3"
}
elseif (Get-Command python3 -ErrorAction SilentlyContinue) {
    $pythonCmd = "python3"
}
elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = "python"
}
else {
    Write-Host "Python is not installed. Please install Python 3.9+ and re-run this installer." -ForegroundColor Red
    exit 1
}

Write-Host "Using Python command: $pythonCmd" -ForegroundColor Green

# --- Clone repo ---
if (!(Test-Path "MarketAdjuster")) {
    git clone https://github.com/TKMAppraisals/MarketAdjuster.git
}

Set-Location MarketAdjuster

# --- Create virtual environment ---
& $pythonCmd -m venv .venv

# --- Activate venv ---
. .\.venv\Scripts\Activate.ps1

# --- Upgrade pip ---
python -m pip install --upgrade pip

# --- Install requirements ---
pip install -r requirements.txt

# --- Launch app ---
streamlit run market_adjustments_app.py

Write-Host "Installation complete." -ForegroundColor Cyan