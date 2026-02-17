# MarketAdjuster – Windows Installer (clean + Desktop shortcut that stays open + logs)
# Run:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_windows.ps1 | iex"

$ErrorActionPreference = "Stop"

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) { return @{ Cmd = "py"; Args = @("-3") } }
    elseif (Get-Command python3 -ErrorAction SilentlyContinue) { return @{ Cmd = "python3"; Args = @() } }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { return @{ Cmd = "python"; Args = @() } }
    else { return $null }
}

Write-Host "Starting MarketAdjuster installation..." -ForegroundColor Cyan

$pyInfo = Get-PythonCommand
if (-not $pyInfo) { throw "Python not found. Install Python 3.10+ and re-run." }
Write-Host ("Using Python: {0} {1}" -f $pyInfo.Cmd, ($pyInfo.Args -join " ")) -ForegroundColor Green

# Install locations (match your current structure)
$installRoot = Join-Path $env:LOCALAPPDATA "MarketAdjuster"
$repoDir     = Join-Path $installRoot "app"
$venvDir1    = Join-Path $installRoot "venv"           # legacy/original installer location
$venvDir2    = Join-Path $repoDir ".venv"              # alternate location
$logsDir     = Join-Path $installRoot "logs"

New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
New-Item -ItemType Directory -Force -Path $logsDir | Out-Null

# --- Acquire source (Git if present, otherwise download ZIP) ---
if (-not (Test-Path $repoDir)) {
    if (Get-Command git -ErrorAction SilentlyContinue) {
        Write-Host "Git detected. Cloning repository..." -ForegroundColor Cyan
        git clone https://github.com/TKMAppraisals/MarketAdjuster.git $repoDir
    } else {
        Write-Host "Git not found. Downloading repository ZIP..." -ForegroundColor Yellow
        $zipUrl      = "https://github.com/TKMAppraisals/MarketAdjuster/archive/refs/heads/main.zip"
        $zipPath     = Join-Path $env:TEMP "MarketAdjuster-main.zip"
        $extractRoot = Join-Path $env:TEMP ("MarketAdjuster-extract-" + [guid]::NewGuid().ToString("N"))

        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath
        Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force

        $extracted = Join-Path $extractRoot "MarketAdjuster-main"
        if (-not (Test-Path $extracted)) { throw "Unexpected ZIP structure. Expected folder: $extracted" }

        Move-Item -Path $extracted -Destination $repoDir -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Host "Existing app folder found at $repoDir" -ForegroundColor Cyan
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

# --- Ensure venv exists (prefer existing legacy venv to avoid duplicating installs) ---
$venvPy = $null
if (Test-Path (Join-Path $venvDir1 "Scripts\python.exe")) {
    $venvPy = Join-Path $venvDir1 "Scripts\python.exe"
    Write-Host "Using existing venv: $venvDir1" -ForegroundColor Cyan
} elseif (Test-Path (Join-Path $venvDir2 "Scripts\python.exe")) {
    $venvPy = Join-Path $venvDir2 "Scripts\python.exe"
    Write-Host "Using existing venv: $venvDir2" -ForegroundColor Cyan
} else {
    # Create legacy venv at %LOCALAPPDATA%\MarketAdjuster\venv (matches your current layout)
    Write-Host "Creating virtual environment at $venvDir1 ..." -ForegroundColor Cyan
    & $pyInfo.Cmd @($pyInfo.Args) -m venv $venvDir1
    $venvPy = Join-Path $venvDir1 "Scripts\python.exe"
}

if (-not (Test-Path $venvPy)) { throw "Virtual environment python not found at $venvPy" }

Write-Host "Upgrading pip..." -ForegroundColor Cyan
& $venvPy -m pip install --upgrade pip

Write-Host "Installing requirements..." -ForegroundColor Cyan
& $venvPy -m pip install -r (Join-Path $repoDir "requirements.txt")

# --- Remove old launchers/shortcuts to prevent duplicates ---
$desktop = [Environment]::GetFolderPath("Desktop")
@(
  (Join-Path $desktop "MarketAdjuster.lnk"),
  (Join-Path $desktop "Market Adjuster.lnk"),
  (Join-Path $desktop "MarketAdjuster (1).lnk"),
  (Join-Path $desktop "Market Adjuster (1).lnk")
) | ForEach-Object { if (Test-Path $_) { Remove-Item $_ -Force -ErrorAction SilentlyContinue } }

@(
  (Join-Path $installRoot "Launch_MarketAdjuster.bat"),
  (Join-Path $installRoot "Launch_MarketAdjuster.vbs"),
  (Join-Path $installRoot "Run Market Adjuster.bat"),
  (Join-Path $installRoot "Run Market Adjuster.ps1"),
  (Join-Path $installRoot "Run MarketAdjuster.bat"),
  (Join-Path $installRoot "Run MarketAdjuster.ps1")
) | ForEach-Object { if (Test-Path $_) { Remove-Item $_ -Force -ErrorAction SilentlyContinue } }

# --- Create a launcher PS1 that NEVER closes instantly (logs + Pause on error) ---
$launcherPs1 = Join-Path $installRoot "Run Market Adjuster.ps1"
$port = 8501
$url = "http://localhost:$port"

$ps1Content = @"
`$ErrorActionPreference = 'Stop'

`$installRoot = `"$installRoot`"
`$repoDir     = `"$repoDir`"
`$venvPy      = `"$venvPy`"
`$app         = `"$app`"
`$port        = $port
`$url         = `"$url`"
`$logsDir     = Join-Path `$installRoot 'logs'

New-Item -ItemType Directory -Force -Path `$logsDir | Out-Null
`$logPath = Join-Path `$logsDir ("launch_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
Start-Transcript -Path `$logPath -Force | Out-Null

try {
  Write-Host "MarketAdjuster launcher started" -ForegroundColor Cyan
  Write-Host "Repo: `$repoDir"
  Write-Host "Python: `$venvPy"
  Write-Host "App: `$app"
  Write-Host "URL: `$url" -ForegroundColor Green

  if (-not (Test-Path `$repoDir)) { throw "Repo folder not found: `$repoDir" }
  if (-not (Test-Path `$venvPy))  { throw "Python not found: `$venvPy" }

  Set-Location `$repoDir

  # Start Streamlit as a child process
  `$args = @('-m','streamlit','run',`"$app`",'--server.port',`"$port`",'--server.address','127.0.0.1','--server.headless','true')
  Write-Host "Starting Streamlit..." -ForegroundColor Cyan
  `$p = Start-Process -FilePath `$venvPy -ArgumentList `$args -PassThru -WindowStyle Normal

  # Wait for port to open (up to ~25s)
  `$opened = `$false
  for (`$i=0; `$i -lt 50; `$i++) {
    try {
      `$client = New-Object Net.Sockets.TcpClient
      `$client.Connect('127.0.0.1', `$port)
      `$client.Close()
      `$opened = `$true
      break
    } catch { Start-Sleep -Milliseconds 500 }
  }

  if (-not `$opened) {
    Write-Host "WARNING: Streamlit did not open port `$port yet. Still trying to open browser..." -ForegroundColor Yellow
  }

  # Open browser (multiple methods)
  Write-Host "Opening browser to `$url ..." -ForegroundColor Cyan
  try { Start-Process `$url | Out-Null } catch {}
  try { Start-Process 'cmd.exe' -ArgumentList @('/c','start','',`"$url`"") | Out-Null } catch {}
  try { Start-Process 'msedge.exe' -ArgumentList `$url | Out-Null } catch {}
  try { Start-Process 'chrome.exe' -ArgumentList `$url | Out-Null } catch {}

  Write-Host "If the browser did not open, manually go to: `$url" -ForegroundColor Yellow
  Write-Host "Close this window to stop the app." -ForegroundColor Yellow

  Wait-Process -Id `$p.Id
}
catch {
  Write-Host "`nERROR: `$($_.Exception.Message)" -ForegroundColor Red
  Write-Host "See log: `$logPath" -ForegroundColor Yellow
  Write-Host "Press any key to close..." -ForegroundColor Yellow
  `$null = `$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
finally {
  Stop-Transcript | Out-Null
}
"@
Set-Content -Path $launcherPs1 -Value $ps1Content -Encoding UTF8

# --- Create ONE Desktop shortcut to powershell.exe running the launcher (shows console so errors are visible) ---
$shortcutPath = Join-Path $desktop "Market Adjuster.lnk"

$repoIcon = Join-Path $repoDir "app_icon.ico"
$iconLocation = "$env:SystemRoot\System32\shell32.dll,167"
if (Test-Path $repoIcon) { $iconLocation = $repoIcon }

$target = (Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe")
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$launcherPs1`""

$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath = $target
$sc.Arguments = $arguments
$sc.WorkingDirectory = $installRoot
$sc.WindowStyle = 1
$sc.IconLocation = $iconLocation
$sc.Save()

Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
Write-Host "Install complete. Click 'Market Adjuster' on the Desktop to run." -ForegroundColor Cyan
Write-Host "If it fails, check logs in: $logsDir" -ForegroundColor Yellow
