<#  Repair-AppxFrameworks.ps1
    - Re-register VCLibs/.NET Native frameworks
    - Optionally re-register Store + App Installer
    - Clear Store cache
    - Kick Intune IME to re-check assignments
    Run in elevated PowerShell.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 1) (Optional) Verify WindowsApps owner (should be NT SERVICE\TrustedInstaller)
try {
    (Get-Acl 'C:\Program Files\WindowsApps').Owner | Out-Host
} catch { Write-Verbose "Couldn't read WindowsApps ACL: $_" }

# Helper to enumerate one or more -Name patterns safely
function Get-AppxByNames {
    param(
        [Parameter(Mandatory)]
        [string[]]$Names,
        [switch]$AllUsers
    )
    foreach ($n in $Names) {
        if ($AllUsers) {
            Get-AppxPackage -AllUsers -Name $n -ErrorAction SilentlyContinue
        } else {
            Get-AppxPackage -Name $n -ErrorAction SilentlyContinue
        }
    }
}

# Targets
$FrameworkNames = @(
    'Microsoft.VCLibs.140.00',
    'Microsoft.NET.Native.Framework*',
    'Microsoft.NET.Native.Runtime*'
)

# 2) Validate install paths for the frameworks/runtimes
Get-AppxByNames -Names $FrameworkNames -AllUsers |
    Select-Object Name, Architecture, Version, InstallLocation, @{
        n = 'PathOk'; e = { Test-Path $_.InstallLocation }
    } | Format-Table -AutoSize | Out-Host

# 3) Re-register VCLibs + .NET Native (all users)
Get-AppxByNames -Names $FrameworkNames -AllUsers |
    ForEach-Object {
        $man = Join-Path $_.InstallLocation 'AppxManifest.xml'
        if (Test-Path $man) {
            Write-Host "Re-registering $($_.Name) $($_.Architecture) $($_.Version)"
            Add-AppxPackage -Register $man -DisableDevelopmentMode -ForceApplicationShutdown
        } else {
            Write-Warning "Manifest not found for $($_.Name) at $($_.InstallLocation)"
        }
    }

# 4) (Optional) Re-register Microsoft Store + App Installer (if present)
$store = Get-AppxByNames -Names 'Microsoft.WindowsStore' -AllUsers | Select-Object -First 1
if ($store) {
    $man = Join-Path $store.InstallLocation 'AppxManifest.xml'
    if (Test-Path $man) {
        Write-Host "Re-registering Microsoft Store"
        Add-AppxPackage -Register $man -DisableDevelopmentMode -ForceApplicationShutdown
    }
}

$appInstaller = Get-AppxByNames -Names 'Microsoft.DesktopAppInstaller' -AllUsers | Select-Object -First 1
if ($appInstaller) {
    $man = Join-Path $appInstaller.InstallLocation 'AppxManifest.xml'
    if (Test-Path $man) {
        Write-Host "Re-registering App Installer"
        Add-AppxPackage -Register $man -DisableDevelopmentMode -ForceApplicationShutdown
    }
}

# 5) Clear Store cache (best-effort)
try { Start-Process -FilePath 'wsreset.exe' -Wait } catch { Write-Verbose "wsreset failed: $_" }

# 6) Trigger Intune IME to re-check assignments/detection
# Preferred: restart the service
try {
    Write-Host "Restarting Intune Management Extension service..."
    Restart-Service -Name 'IntuneManagementExtension' -Force -ErrorAction Stop
} catch {
    Write-Warning "Service restart failed: $_"
}

# Also attempt to run IME client with -c (if present)
$imePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft Intune Management Extension\client\IntuneManagementExtension.exe",
    "${env:ProgramFiles}\Microsoft Intune Management Extension\client\IntuneManagementExtension.exe"
)
$ime = $imePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($ime) {
    Write-Host "Launching IME client: $ime -c"
    Start-Process -FilePath $ime -ArgumentList '-c'
} else {
    Write-Warning "IntuneManagementExtension.exe not found in expected locations."
}
