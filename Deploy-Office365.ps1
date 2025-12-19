<#
.SYNOPSIS
    Deploy-Office365.ps1 - Uninstall existing Office and install fresh from Azure Storage
.DESCRIPTION
    1. Downloads ODT and uninstalls existing Office 365
    2. Downloads Office 365 installation files to a known location
    3. Installs Office 365 from the local files
    All steps run seamlessly without user interaction until completion
    Self-elevates if not running as admin
.NOTES
    Update the Azure Storage URLs below before use
#>

# ============================================
# CONFIGURATION - Azure Storage URLs
# ============================================
$azureStorageBaseUrl = "https://membranding.blob.core.windows.net/branding"
$configXmlUrl = "$azureStorageBaseUrl/Office365/Configuration.xml"

# Known location for Office 365 files
$installDir = "C:\Office365Install"

# ODT download URL
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"

# GitHub URL for this script
$scriptUrl = "https://raw.githubusercontent.com/markorr321/Office_Removal_Re-Install_On_Demmand/main/Deploy-Office365.ps1"

# ============================================
# Self-elevate if not admin
# ============================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Elevating to administrator..." -ForegroundColor Yellow
    Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -Command `"irm '$scriptUrl' | iex`"" -Verb RunAs
    exit
}

$ErrorActionPreference = "Continue"
$ProgressPreference = 'SilentlyContinue'

# Start total elapsed time stopwatch
$totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Clear-Host
Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║      OFFICE 365 DEPLOYMENT TOOL           ║" -ForegroundColor Cyan
Write-Host "  ║   Uninstall + Download + Install          ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

$tempDir = "$env:TEMP\Office365Deploy"
if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force }
New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

# ============================================
# STEP 1: Download ODT for uninstall
# ============================================
Write-Host "  [1/4] Downloading Office Deployment Tool..." -ForegroundColor Yellow

try {
    Invoke-WebRequest -Uri $odtUrl -OutFile "$tempDir\ODT.exe" -UseBasicParsing
    Write-Host "  [1/4] ODT downloaded" -ForegroundColor Green
} catch {
    Write-Host "  [1/4] ERROR: Failed to download ODT - $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Extract ODT
Start-Process -FilePath "$tempDir\ODT.exe" -ArgumentList "/quiet /extract:$tempDir" -Wait

# ============================================
# STEP 2: Uninstall existing Office 365 (if installed)
# ============================================
# Check if Office is installed
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
    Write-Host "  [2/4] Uninstalling existing Office 365..." -ForegroundColor Yellow

    try {
        # Create uninstall config
@"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@ | Out-File -FilePath "$tempDir\uninstall.xml" -Encoding UTF8

        $uninstallProcess = Start-Process -FilePath "$tempDir\setup.exe" -ArgumentList "/configure `"$tempDir\uninstall.xml`"" -PassThru -WindowStyle Hidden

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
    } catch {
        Write-Host ""
        Write-Host "  [2/4] WARNING: Uninstall encountered an issue - continuing..." -ForegroundColor Yellow
        Write-Host "        $($_.Exception.Message)" -ForegroundColor Gray
    }
} else {
    Write-Host "  [2/4] No existing Office installation found - skipping uninstall" -ForegroundColor Green
}

# ============================================
# STEP 3: Download Office 365 files from Azure
# ============================================
Write-Host "  [3/4] Downloading Office 365 files to $installDir..." -ForegroundColor Yellow

if (Test-Path $installDir) {
    Remove-Item -Path $installDir -Recurse -Force
}
New-Item -Path $installDir -ItemType Directory -Force | Out-Null

try {
    Invoke-WebRequest -Uri $setupExeUrl -OutFile "$installDir\setup.exe" -UseBasicParsing
    Invoke-WebRequest -Uri $configXmlUrl -OutFile "$installDir\Configuration.xml" -UseBasicParsing
    Write-Host "  [3/4] Download complete" -ForegroundColor Green
} catch {
    Write-Host "  [3/4] ERROR: Failed to download files - $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ============================================
# STEP 4: Install Office 365
# ============================================
Write-Host "  [4/4] Installing Office 365..." -ForegroundColor Yellow

$installProcess = Start-Process -FilePath "$installDir\setup.exe" -ArgumentList "/configure `"$installDir\Configuration.xml`"" -PassThru -WindowStyle Hidden

$spinner = @('|', '/', '-', '\')
$i = 0
$elapsed = 0

while (-not $installProcess.HasExited) {
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

if ($installProcess.ExitCode -eq 0) {
    Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║      SUCCESS: Office 365 Deployed         ║" -ForegroundColor Green
    Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Green
} else {
    Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "  ║  WARNING: Deployment may need attention   ║" -ForegroundColor Yellow
    Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Yellow
}

# Stop the stopwatch and display total time
$totalStopwatch.Stop()
$totalMinutes = [math]::Floor($totalStopwatch.Elapsed.TotalMinutes)
$totalSeconds = $totalStopwatch.Elapsed.Seconds

Write-Host ""
Write-Host "  Install files saved to: $installDir" -ForegroundColor Gray
Write-Host "  Total time: $($totalMinutes.ToString('00')):$($totalSeconds.ToString('00'))" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
