# MarketAdjuster – Windows Installer (single Desktop shortcut, reliable launch)
# Run:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/TKMAppraisals/MarketAdjuster/main/install_windows.ps1 | iex"

$ErrorActionPreference = "Stop"

Write-Host "Starting MarketAdjuster installation..." -ForegroundColor Cyan

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) { return @{ Cmd = "py"; Args = @("-3") } }
    elseif (Get-Command python3 -ErrorAction SilentlyContinue) { return @{ Cmd = "python3"; Args = @() } }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { return @{ Cmd = "python"; Args = @() } }
    else { return $null }
}

$pyInfo = Get-PythonCommand
if (-not $pyInfo) { throw "Python not found. Install Python 3.10+ and re-run." }
Write-Host ("Using Python: {0} {1}" -f $pyInfo.Cmd, ($pyInfo.Args -join " ")) -ForegroundColor Green

# Install location (per-user)
$installRoot = Join-Path $env:LOCALAPPDATA "MarketAdjuster"
$repoDir     = Join-Path $installRoot "app"
New-Item -ItemType Directory -Force -Path $installRoot | Out-Null

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

# --- Create virtual environment (idempotent) ---
Write-Host "Creating/updating virtual environment..." -ForegroundColor Cyan
& $pyInfo.Cmd @($pyInfo.Args) -m venv .venv

$venvPy = Join-Path $PWD ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) { throw "Virtual environment python not found at $venvPy" }

Write-Host "Upgrading pip..." -ForegroundColor Cyan
& $venvPy -m pip install --upgrade pip

Write-Host "Installing requirements..." -ForegroundColor Cyan
& $venvPy -m pip install -r requirements.txt

# --- Launcher PS1 (the shortcut runs THIS) ---
$launcherPs1 = Join-Path $installRoot "Run MarketAdjuster.ps1"
$port = 8501
$address = "127.0.0.1"
$url = "http://localhost:$port"

$ps1Content = @"
`$ErrorActionPreference = 'Stop'
Set-Location `"$repoDir`"

`$venvPy = Join-Path (Get-Location) '.venv\Scripts\python.exe'
if (-not (Test-Path `$venvPy)) {
  Write-Host 'Venv not found. Re-run install_windows.ps1.' -ForegroundColor Red
  Pause
  exit 1
}

`$port = $port
`$address = '$address'
`$url = '$url'

Write-Host "Starting MarketAdjuster..." -ForegroundColor Cyan
Write-Host "App will be available at: `$url" -ForegroundColor Green

# Start Streamlit (separate process)
`$args = @('-m','streamlit','run',`"$app`",'--server.port',`"$port`",'--server.address',`"$address`",'--server.headless','true')
`$p = Start-Process -FilePath `$venvPy -ArgumentList `$args -PassThru -WindowStyle Normal

# Wait for port to open (up to ~20s)
for (`$i=0; `$i -lt 40; `$i++) {
  try {
    `$client = New-Object Net.Sockets.TcpClient
    `$client.Connect('127.0.0.1', `$port)
    `$client.Close()
    break
  } catch {
    Start-Sleep -Milliseconds 500
  }
}

# Open browser (best-effort)
try {
  Start-Process `$url | Out-Null
} catch {
  try { Start-Process 'cmd.exe' -ArgumentList @('/c','start','',`"$url`"") | Out-Null } catch {}
}

Write-Host "If the browser did not open automatically, paste this into your browser: `$url" -ForegroundColor Yellow
Write-Host "Close this window to stop the app." -ForegroundColor Yellow

Wait-Process -Id `$p.Id
"@
Set-Content -Path $launcherPs1 -Value $ps1Content -Encoding UTF8

# --- Create ONE Desktop shortcut (.lnk) pointing directly to powershell.exe ---
$desktop = [Environment]::GetFolderPath("Desktop")
# Remove any older shortcuts we might have created
@(
  (Join-Path $desktop "MarketAdjuster.lnk"),
  (Join-Path $desktop "Market Adjuster.lnk")
) | ForEach-Object {
  if (Test-Path $_) { Remove-Item $_ -Force -ErrorAction SilentlyContinue }
}

$shortcutPath = Join-Path $desktop "Market Adjuster.lnk"

$repoIcon = Join-Path $repoDir "app_icon.ico"
$iconLocation = "$env:SystemRoot\System32\shell32.dll,167"
if (Test-Path $repoIcon) { $iconLocation = $repoIcon }

$target = (Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe")
$args = "-NoProfile -ExecutionPolicy Bypass -File `"$launcherPs1`""

$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath = $target
$sc.Arguments = $args
$sc.WorkingDirectory = $installRoot
$sc.WindowStyle = 1
$sc.IconLocation = $iconLocation
$sc.Save()

Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
Write-Host "Install complete. Launch using the 'Market Adjuster' desktop icon." -ForegroundColor Cyan
