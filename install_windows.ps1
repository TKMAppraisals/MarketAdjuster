# MarketAdjuster – Windows Installer (BAT launcher + single Desktop shortcut)
# Goal: Desktop icon launches Streamlit and opens browser reliably, without PowerShell policy issues.
# Run:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_windows.ps1 | iex"

$ErrorActionPreference = "Stop"

Write-Host "Installing MarketAdjuster..." -ForegroundColor Cyan

# Paths (match your existing layout)
$installRoot = Join-Path $env:LOCALAPPDATA "MarketAdjuster"
$repoDir     = Join-Path $installRoot "app"
$venvDir     = Join-Path $installRoot "venv"

# Sanity checks
if (-not (Test-Path $repoDir)) {
  throw "App folder not found at: $repoDir  (Run your existing installer steps that create the app folder first.)"
}
$venvPy = Join-Path $venvDir "Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
  throw "Venv python not found at: $venvPy  (Your venv folder exists but is missing Scripts\\python.exe.)"
}

# App entrypoint
$app = "market_condition_app_v4_15_premium_plus.py"
if (-not (Test-Path (Join-Path $repoDir $app))) {
  throw "App file not found: $app in $repoDir"
}

# Clean old launchers / shortcuts (prevents duplicate icons)
$desktop = [Environment]::GetFolderPath("Desktop")
@(
  (Join-Path $desktop "MarketAdjuster.lnk"),
  (Join-Path $desktop "Market Adjuster.lnk"),
  (Join-Path $desktop "MarketAdjuster.url"),
  (Join-Path $desktop "Market Adjuster.url")
) | ForEach-Object { if (Test-Path $_) { Remove-Item $_ -Force -ErrorAction SilentlyContinue } }

@(
  (Join-Path $installRoot "Launch_MarketAdjuster.bat"),
  (Join-Path $installRoot "Launch_MarketAdjuster.vbs"),
  (Join-Path $installRoot "Run Market Adjuster.bat"),
  (Join-Path $installRoot "Run MarketAdjuster.bat")
) | ForEach-Object { if (Test-Path $_) { Remove-Item $_ -Force -ErrorAction SilentlyContinue } }

# Create BAT launcher (keeps window open and shows errors)
$launcherBat = Join-Path $installRoot "Run Market Adjuster.bat"
$port = 8501

$bat = @"
@echo off
setlocal
set "INSTALLROOT=%LOCALAPPDATA%\MarketAdjuster"
set "APPDIR=%INSTALLROOT%\app"
set "PY=%INSTALLROOT%\venv\Scripts\python.exe"
set "APP=market_condition_app_v4_15_premium_plus.py"
set "URL=http://localhost:{port}"

echo Starting MarketAdjuster...
echo App folder: %APPDIR%
echo Python: %PY%
echo URL: %URL%
echo.

if not exist "%PY%" (
  echo ERROR: Python not found at "%PY%"
  echo Re-run installer.
  pause
  exit /b 1
)

if not exist "%APPDIR%\%APP%" (
  echo ERROR: App file not found: "%APPDIR%\%APP%"
  pause
  exit /b 1
)

cd /d "%APPDIR%"

rem Open the browser (best effort). If Streamlit takes a moment, refresh once it loads.
start "" "%URL%"

rem Run Streamlit (this window stays open)
"%PY%" -m streamlit run "%APP%" --server.address 127.0.0.1 --server.port {port}
set "EC=%ERRORLEVEL%"

echo.
echo Streamlit exited with code %EC%
pause
exit /b %EC%
"@.Replace("{port}", $port.ToString())

Set-Content -Path $launcherBat -Value $bat -Encoding ASCII
Write-Host "Launcher created: $launcherBat" -ForegroundColor Green

# Create single Desktop shortcut to the BAT launcher
$shortcutPath = Join-Path $desktop "Market Adjuster.lnk"

$repoIcon = Join-Path $repoDir "app_icon.ico"
$iconLocation = "$env:SystemRoot\System32\shell32.dll,167"
if (Test-Path $repoIcon) { $iconLocation = $repoIcon }

$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath = $launcherBat
$sc.WorkingDirectory = $installRoot
$sc.WindowStyle = 1
$sc.IconLocation = $iconLocation
$sc.Save()

Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
Write-Host "Done. Double-click 'Market Adjuster' on the Desktop." -ForegroundColor Cyan
