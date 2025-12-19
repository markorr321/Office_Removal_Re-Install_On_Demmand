# Office 365 Removal & Re-Install On Demand

A PowerShell script that automates the uninstallation and reinstallation of Microsoft 365 / Office 365.

## Quick Start

Run this command in PowerShell:

```powershell
irm bit.ly/O365Deploy | iex
```

Or the full URL:

```powershell
irm https://raw.githubusercontent.com/markorr321/Office_Removal_Re-Install_On_Demmand/main/Deploy-Office365.ps1 | iex
```

## What It Does

1. **Downloads** the Office Deployment Tool (ODT) from Microsoft
2. **Uninstalls** existing Office 365 (if installed) - displays completion notification
3. **Downloads** installation files from Azure Storage
4. **Installs** Office 365 with your custom configuration

The script automatically elevates to administrator if needed.

## Configuration

The script installs **Microsoft 365 Apps for Business (O365BusinessRetail)** with:

- **Architecture:** 64-bit
- **Update Channel:** Current
- **Included Apps:** Word, Excel, PowerPoint, OneNote
- **Excluded Apps:** Access, Groove, Lync/Skype, OneDrive, Outlook, Publisher

## Verify Installation

After installation, verify the correct version with PowerShell:

```powershell
# Check installed Office product
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like "*Microsoft 365*" -or $_.DisplayName -like "*Office*" } | Select-Object DisplayName, DisplayVersion
```

```powershell
# Check Office configuration
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" | Select-Object ProductReleaseIds, Platform, UpdateChannel
```

Expected output:
- **ProductReleaseIds:** O365BusinessRetail
- **Platform:** x64

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- Internet connection
- Administrator privileges (script self-elevates)

## Files

- `Deploy-Office365.ps1` - Main deployment script
- `Configuration.xml` - Office installation configuration (hosted on Azure Storage)

## Customization

To modify the Office configuration (add/remove apps, change update channel, etc.), edit the `Configuration.xml` file in Azure Storage.

Use the [Office Customization Tool](https://config.office.com/) to generate a new configuration.
