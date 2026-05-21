# Run this script once as Administrator in PowerShell
# Right-click PowerShell -> "Run as Administrator", then run:
#   cd "C:\Users\omara\OneDrive - Al-Arabia Educational Enterprises Company\web"
#   Set-ExecutionPolicy Bypass -Scope Process -Force
#   .\install_as_service.ps1

$ServiceName  = "FlaskWebApp"
$AppDir       = "C:\Users\omara\OneDrive - Al-Arabia Educational Enterprises Company\web"
$WaitressExe  = "C:\Users\omara\AppData\Local\Programs\Python\Python313\Scripts\waitress-serve.exe"
$NssmDir      = "C:\nssm"
$NssmExe      = "$NssmDir\nssm.exe"
$LogDir       = "$AppDir\logs"

# ── 1. Create log folder ─────────────────────────────────────────────────────
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
Write-Host "[1/5] Log folder ready: $LogDir"

# ── 2. Download NSSM if not present ──────────────────────────────────────────
if (-not (Test-Path $NssmExe)) {
    Write-Host "[2/5] Downloading NSSM..."
    $zip = "$env:TEMP\nssm.zip"
    Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $zip
    Expand-Archive -Path $zip -DestinationPath $env:TEMP -Force
    New-Item -ItemType Directory -Force -Path $NssmDir | Out-Null
    Copy-Item "$env:TEMP\nssm-2.24\win64\nssm.exe" $NssmExe -Force
    Write-Host "    NSSM installed to $NssmExe"
} else {
    Write-Host "[2/5] NSSM already present, skipping download."
}

# ── 3. Remove existing service if it exists ───────────────────────────────────
$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "[3/5] Removing old service..."
    & $NssmExe stop $ServiceName | Out-Null
    & $NssmExe remove $ServiceName confirm | Out-Null
} else {
    Write-Host "[3/5] No existing service found, continuing."
}

# ── 4. Install the service ────────────────────────────────────────────────────
Write-Host "[4/5] Installing Windows service '$ServiceName'..."
& $NssmExe install $ServiceName $WaitressExe "--host=127.0.0.1 --port=5000 app:app"

# Service settings
& $NssmExe set $ServiceName AppDirectory      $AppDir
& $NssmExe set $ServiceName DisplayName       "Flask Web App"
& $NssmExe set $ServiceName Description       "Flask web application served by Waitress"
& $NssmExe set $ServiceName Start             SERVICE_AUTO_START

# Stdout / Stderr logs (NSSM rotates these)
& $NssmExe set $ServiceName AppStdout         "$LogDir\stdout.log"
& $NssmExe set $ServiceName AppStderr         "$LogDir\stderr.log"
& $NssmExe set $ServiceName AppRotateFiles    1
& $NssmExe set $ServiceName AppRotateBytes    5242880   # 5 MB per file

# Auto-restart on crash: restart after 3 s, then 10 s, then 30 s
& $NssmExe set $ServiceName AppThrottle       3000
& $NssmExe set $ServiceName AppRestartDelay   3000

# ── 5. Start the service ──────────────────────────────────────────────────────
Write-Host "[5/5] Starting service..."
& $NssmExe start $ServiceName

Start-Sleep -Seconds 3
$svc = Get-Service -Name $ServiceName
Write-Host ""
Write-Host "========================================"
Write-Host "Service status: $($svc.Status)"
Write-Host "Open http://127.0.0.1:5000 to verify."
Write-Host ""
Write-Host "Useful commands:"
Write-Host "  Stop:    nssm stop FlaskWebApp"
Write-Host "  Start:   nssm start FlaskWebApp"
Write-Host "  Restart: nssm restart FlaskWebApp"
Write-Host "  Remove:  nssm remove FlaskWebApp confirm"
Write-Host "  Logs:    $LogDir"
Write-Host "========================================"
