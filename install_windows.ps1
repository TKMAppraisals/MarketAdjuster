# Market Adjuster – Windows Installer (robust, no Git required)
# Run with:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_windows.ps1 | iex"

$ErrorActionPreference = "Stop"

Write-Host "Starting Market Adjuster installation..." -ForegroundColor Cyan

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return @{ Cmd = "py"; Args = @("-3") }
    }
    elseif (Get-Command python3 -ErrorAction SilentlyContinue) {
        return @{ Cmd = "python3"; Args = @() }
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        return @{ Cmd = "python"; Args = @() }
    }
    else {
        return $null
    }
}

$pyInfo = Get-PythonCommand
if (-not $pyInfo) {
    Write-Host "Python is not installed. Install Python 3.9+ (with 'Add to PATH') and re-run." -ForegroundColor Red
    exit 1
}

Write-Host ("Using Python: {0} {1}" -f $pyInfo.Cmd, ($pyInfo.Args -join " ")) -ForegroundColor Green

# --- Acquire source (Git if present, otherwise download ZIP) ---
$repoDir = Join-Path $PWD "MarketAdjuster"

if (-not (Test-Path $repoDir)) {
    if (Get-Command git -ErrorAction SilentlyContinue) {
        Write-Host "Git detected. Cloning repository..." -ForegroundColor Cyan
        git clone https://github.com/TKMAppraisals/MarketAdjuster.git $repoDir
    }
    else {
        Write-Host "Git not found. Downloading repository ZIP..." -ForegroundColor Yellow
        $zipUrl = "https://github.com/TKMAppraisals/MarketAdjuster/archive/refs/heads/main.zip"
        $zipPath = Join-Path $env:TEMP "MarketAdjuster-main.zip"
        $extractRoot = Join-Path $env:TEMP ("MarketAdjuster-extract-" + [guid]::NewGuid().ToString("N"))

        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath
        Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force

        $extracted = Join-Path $extractRoot "MarketAdjuster-main"
        if (-not (Test-Path $extracted)) {
            throw "Unexpected ZIP structure. Expected folder: $extracted"
        }

        Move-Item -Path $extracted -Destination $repoDir -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
else {
    Write-Host "MarketAdjuster folder already exists. Using existing files." -ForegroundColor Cyan
}

Set-Location $repoDir

# --- Create virtual environment ---
Write-Host "Creating virtual environment..." -ForegroundColor Cyan
& $pyInfo.Cmd @($pyInfo.Args) -m venv .venv

$venvPy = Join-Path $PWD ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
    throw "Virtual environment python not found at $venvPy"
}

# --- Install dependencies (avoid relying on PATH) ---
Write-Host "Upgrading pip..." -ForegroundColor Cyan
& $venvPy -m pip install --upgrade pip

Write-Host "Installing requirements..." -ForegroundColor Cyan
& $venvPy -m pip install -r requirements.txt

# --- Launch app ---
Write-Host "Launching Market Adjuster..." -ForegroundColor Cyan
& $venvPy -m streamlit run market_adjustments_app.py
