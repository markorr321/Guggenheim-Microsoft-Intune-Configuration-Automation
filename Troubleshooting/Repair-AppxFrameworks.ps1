# 1) Verify WindowsApps owner (should be NT SERVICE\TrustedInstaller)
(Get-Acl 'C:\Program Files\WindowsApps').Owner

# 2) Validate install paths for the frameworks/runtimes
Get-AppxPackage -AllUsers -Name Microsoft.VCLibs.140.00, Microsoft.NET.Native.Framework*, Microsoft.NET.Native.Runtime* |
  Select Name, Architecture, Version, InstallLocation, @{
    n='PathOk'; e={ Test-Path $_.InstallLocation }
  }

# 3) Re-register VCLibs + .NET Native (all users)
Get-AppxPackage -AllUsers -Name Microsoft.VCLibs.140.00, Microsoft.NET.Native.Framework*, Microsoft.NET.Native.Runtime* |
  ForEach-Object {
    $man = Join-Path $_.InstallLocation 'AppxManifest.xml'
    if (Test-Path $man) {
      Write-Host "Re-registering $($_.Name) $($_.Architecture) $($_.Version)"
      Add-AppxPackage -Register $man -DisableDevelopmentMode -ForceApplicationShutdown
    }
  }

# 4) Re-register Microsoft Store plumbing (common after migrations)
$store = Get-AppxPackage -AllUsers -Name Microsoft.WindowsStore
if ($store) {
  Add-AppxPackage -Register (Join-Path $store.InstallLocation 'AppxManifest.xml') -DisableDevelopmentMode -ForceApplicationShutdown
}
$appInstaller = Get-AppxPackage -AllUsers -Name Microsoft.DesktopAppInstaller
if ($appInstaller) {
  Add-AppxPackage -Register (Join-Path $appInstaller.InstallLocation 'AppxManifest.xml') -DisableDevelopmentMode -ForceApplicationShutdown
}

# 5) Optional: clear Store cache
Start-Process -FilePath wsreset.exe -Wait

# 6) Retry Company Portal install/detection
Start-Process "IntuneManagementExtension.exe" -ArgumentList "-c"
