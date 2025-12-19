<#
.SYNOPSIS
    Deploy-Office365.ps1 - Uninstall existing Office and install fresh from Azure Storage
.DESCRIPTION
    1. Uninstalls existing Office 365 (if installed)
    2. Downloads Office 365 installation files from Azure
    3. Installs Office 365
    Self-elevates if not running as admin
#>

# ============================================
# CONFIGURATION
# ============================================
$azureStorageBaseUrl = "https://membranding.blob.core.windows.net/branding"
$setupExeUrl = "$azureStorageBaseUrl/Office365/setup.exe"
$configXmlUrl = "$azureStorageBaseUrl/Office365/Configuration.xml"
$installDir = "C:\Office365Install"
$scriptUrl = "https://raw.githubusercontent.com/markorr321/Office_Removal_Re-Install_On_Demmand/main/Deploy-Office365.ps1"

# ============================================
# Self-elevate if not admin
# ============================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Elevating to administrator..." -ForegroundColor Yellow
    Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -NoExit -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm '$scriptUrl' | iex`"" -Verb RunAs
    exit
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Continue"
$ProgressPreference = 'SilentlyContinue'
$totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Clear-Host
Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║      OFFICE 365 DEPLOYMENT TOOL           ║" -ForegroundColor Cyan
Write-Host "  ║   Uninstall + Download + Install          ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ============================================
# STEP 1: Download Office 365 files from Azure
# ============================================
Write-Host "  [1/3] Downloading Office 365 files to $installDir..." -ForegroundColor Yellow

if (Test-Path $installDir) { Remove-Item -Path $installDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -Path $installDir -ItemType Directory -Force | Out-Null

try {
    Invoke-WebRequest -Uri $setupExeUrl -OutFile "$installDir\setup.exe" -UseBasicParsing
    Invoke-WebRequest -Uri $configXmlUrl -OutFile "$installDir\Configuration.xml" -UseBasicParsing
    Write-Host "  [1/3] Download complete" -ForegroundColor Green
} catch {
    Write-Host "  [1/3] ERROR: Failed to download files - $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ============================================
# STEP 2: Uninstall existing Office 365 (if installed)
# ============================================
$officeInstalled = $false
$officeRegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\O365*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office*"
)

foreach ($path in $officeRegPaths) {
    if (Get-Item -Path $path -ErrorAction SilentlyContinue) {
        $officeInstalled = $true
        break
    }
}

if ($officeInstalled) {
    Write-Host "  [2/3] Uninstalling existing Office 365..." -ForegroundColor Yellow

    # Create uninstall config
@"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@ | Out-File -FilePath "$installDir\uninstall.xml" -Encoding UTF8

    $uninstallProcess = Start-Process -FilePath "$installDir\setup.exe" -ArgumentList "/configure `"$installDir\uninstall.xml`"" -PassThru -WindowStyle Hidden

    $spinner = @('|', '/', '-', '\')
    $i = 0
    $elapsed = 0

    while ($uninstallProcess -and -not $uninstallProcess.HasExited) {
        $elapsed++
        $minutes = [math]::Floor($elapsed / 60)
        $seconds = $elapsed % 60
        $spin = $spinner[$i % 4]
        Write-Host "`r        $spin  Removing Office... [$($minutes.ToString('00')):$($seconds.ToString('00'))] " -NoNewline -ForegroundColor White
        $i++
        Start-Sleep -Seconds 1
    }

    Write-Host ""
    Write-Host ""
    Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║      Office 365 Uninstall Complete        ║" -ForegroundColor Green
    Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Start-Sleep -Seconds 2
} else {
    Write-Host "  [2/3] No existing Office installation found - skipping uninstall" -ForegroundColor Green
}

# ============================================
# STEP 3: Install Office 365
# ============================================
Write-Host "  [3/3] Installing Office 365..." -ForegroundColor Yellow

$installProcess = Start-Process -FilePath "$installDir\setup.exe" -ArgumentList "/configure `"$installDir\Configuration.xml`"" -PassThru -WindowStyle Hidden

$spinner = @('|', '/', '-', '\')
$i = 0
$elapsed = 0

while ($installProcess -and -not $installProcess.HasExited) {
    $elapsed++
    $minutes = [math]::Floor($elapsed / 60)
    $seconds = $elapsed % 60
    $spin = $spinner[$i % 4]
    Write-Host "`r        $spin  Installing Office... [$($minutes.ToString('00')):$($seconds.ToString('00'))] " -NoNewline -ForegroundColor White
    $i++
    Start-Sleep -Seconds 1
}

Write-Host ""
Write-Host ""

$totalStopwatch.Stop()
$totalMinutes = [math]::Floor($totalStopwatch.Elapsed.TotalMinutes)
$totalSeconds = $totalStopwatch.Elapsed.Seconds

if ($installProcess.ExitCode -eq 0) {
    Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║      SUCCESS: Office 365 Deployed         ║" -ForegroundColor Green
    Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Green
} else {
    Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "  ║  WARNING: Deployment may need attention   ║" -ForegroundColor Yellow
    Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  Install files saved to: $installDir" -ForegroundColor Gray
Write-Host "  Total time: $($totalMinutes.ToString('00')):$($totalSeconds.ToString('00'))" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
