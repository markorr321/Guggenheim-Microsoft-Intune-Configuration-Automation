# Collect-CompanyPortalLogs.ps1
[CmdletBinding()]
param(
  [string]$OutRoot = 'C:\Temp\CPortalLogs',
  [string]$AppFilter = 'CompanyPortal'   # pattern used in message filters and package search
)

New-Item -Path $OutRoot -ItemType Directory -Force | Out-Null

# Ensure relevant logs are enabled (no-op if already enabled)
$logs = @(
  "Microsoft-Windows-AppXDeploymentServer/Operational",
  "Microsoft-Windows-AppXDeployment/Operational",
  "Microsoft-Windows-AppReadiness/Admin",
  "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin"
)
foreach ($ln in $logs) { wevtutil set-log $ln /e:true | Out-Null }

# Export raw event logs
wevtutil epl "Microsoft-Windows-AppXDeploymentServer/Operational" "$OutRoot\AppXDeploymentServer-Operational.evtx" /ow:true
wevtutil epl "Microsoft-Windows-AppXDeployment/Operational"       "$OutRoot\AppXDeployment-Operational.evtx"       /ow:true
wevtutil epl "Microsoft-Windows-AppReadiness/Admin"               "$OutRoot\AppReadiness-Admin.evtx"               /ow:true
wevtutil epl "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin" "$OutRoot\DMClient-Admin.evtx" /ow:true

# Build summaries (wider filter catches both short & full names)
$summary = Join-Path $OutRoot '_Summary.txt'
"==== AppXDeploymentServer ====" | Out-File $summary -Force

Get-WinEvent -LogName "Microsoft-Windows-AppXDeploymentServer/Operational" -ErrorAction SilentlyContinue |
  Where-Object { $_.Message -match $AppFilter -or $_.Message -match 'Microsoft\.CompanyPortal' } |
  Sort-Object TimeCreated |
  Format-Table TimeCreated, Id, LevelDisplayName, Message -Wrap |
  Out-File -Append $summary

"==== AppXDeployment ====" | Out-File $summary -Append
Get-WinEvent -LogName "Microsoft-Windows-AppXDeployment/Operational" -ErrorAction SilentlyContinue |
  Where-Object { $_.Message -match $AppFilter -or $_.Message -match 'Microsoft\.CompanyPortal' } |
  Sort-Object TimeCreated |
  Format-Table TimeCreated, Id, LevelDisplayName, Message -Wrap |
  Out-File -Append $summary

"==== AppReadiness ====" | Out-File $summary -Append
Get-WinEvent -LogName "Microsoft-Windows-AppReadiness/Admin" -ErrorAction SilentlyContinue |
  Where-Object { $_.Message -match $AppFilter -or $_.Message -match 'Microsoft\.CompanyPortal' } |
  Sort-Object TimeCreated |
  Format-Table TimeCreated, Id, LevelDisplayName, Message -Wrap |
  Out-File -Append $summary

"==== DMClient/Admin (Intune MDM) ====" | Out-File $summary -Append
Get-WinEvent -LogName "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin" -ErrorAction SilentlyContinue |
  Where-Object { $_.Message -match $AppFilter -or $_.Message -match 'Microsoft\.CompanyPortal' } |
  Sort-Object TimeCreated |
  Format-Table TimeCreated, Id, LevelDisplayName, Message -Wrap |
  Out-File -Append $summary

# Grab Company Portal diag folders from ALL local profiles (if present)
$profileRoots = Get-ChildItem C:\Users -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -notin 'Public','Default','Default User','All Users','WDAGUtilityAccount' }

foreach ($pr in $profileRoots) {
  $packages = Join-Path $pr.FullName 'AppData\Local\Packages'
  if (-not (Test-Path $packages)) { continue }

  $cpPkgs = Get-ChildItem $packages -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like 'Microsoft.CompanyPortal_*' }

  foreach ($pkg in $cpPkgs) {
    $diag = Join-Path $pkg.FullName 'LocalState\DiagOutputDir'
    if (Test-Path $diag) {
      $target = Join-Path $OutRoot ("CompanyPortal-DiagOutputDir-" + $pr.Name)
      New-Item -ItemType Directory -Path $target -Force | Out-Null
      Copy-Item "$diag\*" -Destination $target -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

# Zip for easy sharing
$zipPath = Join-Path (Split-Path $OutRoot -Parent) 'CPortalLogs.zip'
Compress-Archive -Path "$OutRoot\*" -DestinationPath $zipPath -Force
Write-Host "Collected logs and summaries in $OutRoot and $zipPath"
