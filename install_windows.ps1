# Market Adjuster – Windows Installer (creates Desktop shortcut)
# Recommended run:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_windows.ps1 | iex"

$ErrorActionPreference = "Stop"

Write-Host "Starting Market Adjuster installation..." -ForegroundColor Cyan

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) { return @{ Cmd = "py"; Args = @("-3") } }
    elseif (Get-Command python3 -ErrorAction SilentlyContinue) { return @{ Cmd = "python3"; Args = @() } }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { return @{ Cmd = "python"; Args = @() } }
    else { return $null }
}

$pyInfo = Get-PythonCommand
if (-not $pyInfo) {
    Write-Host "Python is not installed. Install Python 3.9+ (with 'Add to PATH') and re-run." -ForegroundColor Red
    exit 1
}

Write-Host ("Using Python: {0} {1}" -f $pyInfo.Cmd, ($pyInfo.Args -join " ")) -ForegroundColor Green

# Install location (per-user)
$installRoot = Join-Path $env:LOCALAPPDATA "MarketAdjuster"
$repoDir      = Join-Path $installRoot "app"

New-Item -ItemType Directory -Force -Path $installRoot | Out-Null

# --- Acquire source (Git if present, otherwise download ZIP) ---
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
        if (-not (Test-Path $extracted)) { throw "Unexpected ZIP structure. Expected folder: $extracted" }

        Move-Item -Path $extracted -Destination $repoDir -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
else {
    Write-Host "Existing install found at $repoDir" -ForegroundColor Cyan
}

Set-Location $repoDir

# --- Choose the app entrypoint ---
$appCandidates = @(
    "market_condition_app_v4_15_premium_plus.py",
    "market_adjustments_app.py",
    "app.py"
)

$app = $null
foreach ($c in $appCandidates) {
    if (Test-Path (Join-Path $PWD $c)) { $app = $c; break }
}
if (-not $app) {
    Write-Host "Could not find the Streamlit app file in repo root." -ForegroundColor Red
    Get-ChildItem -File | Select-Object -ExpandProperty Name | ForEach-Object { Write-Host ("  - " + $_) -ForegroundColor Yellow }
    throw "No app entrypoint found."
}
Write-Host ("Using app file: {0}" -f $app) -ForegroundColor Green

# --- Create virtual environment ---
Write-Host "Creating virtual environment..." -ForegroundColor Cyan
& $pyInfo.Cmd @($pyInfo.Args) -m venv .venv

$venvPy = Join-Path $PWD ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) { throw "Virtual environment python not found at $venvPy" }

# --- Install dependencies ---
Write-Host "Upgrading pip..." -ForegroundColor Cyan
& $venvPy -m pip install --upgrade pip

Write-Host "Installing requirements..." -ForegroundColor Cyan
& $venvPy -m pip install -r requirements.txt

# --- Create launcher scripts ---
$launcherPs1 = Join-Path $installRoot "Run Market Adjuster.ps1"
$launcherBat = Join-Path $installRoot "Run Market Adjuster.bat"

$ps1Content = @"
`$ErrorActionPreference = 'Stop'
Set-Location `"$repoDir`"
`$venvPy = Join-Path (Get-Location) '.venv\Scripts\python.exe'
if (-not (Test-Path `$venvPy)) {
  Write-Host 'Venv not found. Re-run install_windows.ps1.' -ForegroundColor Red
  Pause
  exit 1
}
Write-Host 'Starting Market Adjuster (Streamlit)...' -ForegroundColor Cyan
Start-Process "http://localhost:8501" | Out-Null
& `$venvPy -m streamlit run `"$app`"
"@

Set-Content -Path $launcherPs1 -Value $ps1Content -Encoding UTF8

$batContent = @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "`"$launcherPs1`""
"@
Set-Content -Path $launcherBat -Value $batContent -Encoding ASCII

# --- Create Desktop shortcut (.lnk) pointing to the BAT launcher ---
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Market Adjuster.lnk"

$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath = $launcherBat
$sc.WorkingDirectory = $installRoot
$sc.WindowStyle = 1
$sc.IconLocation = "$env:SystemRoot\System32\shell32.dll,167"  # generic app icon
$sc.Save()

Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
Write-Host "Install complete. Launch using the 'Market Adjuster' desktop icon." -ForegroundColor Cyan

# Optional: launch now in a new window so the installer can exit without stopping Streamlit
Write-Host "Launching now..." -ForegroundColor Cyan
Start-Process -FilePath "powershell" -ArgumentList @("-NoProfile","-ExecutionPolicy","Bypass","-File",$launcherPs1) -WorkingDirectory $installRoot
