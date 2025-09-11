# Ensure the script is running with Administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as an administrator." -ForegroundColor Red
    exit
}

# Persistent flag location
$MarkerFile = "C:\ProgramData\TeamsAndOutlookRemoval.flag"

# Check if the script has already run
if (Test-Path $MarkerFile) {
    Write-Host "Removal script has already been executed. Exiting..." -ForegroundColor Green
    exit 0
}

Write-Host "Starting enhanced removal of New Microsoft Teams and New Outlook..." -ForegroundColor Green

# Function to Take Ownership and Grant Access
function Grant-WindowsAppsAccess {
    param (
        [string]$Path
    )

    if (Test-Path $Path) {
        Write-Host "Taking ownership of $Path..." -ForegroundColor Yellow
        try {
            Start-Process -FilePath "takeown" -ArgumentList "/f `"$Path`" /r /d y" -Wait -NoNewWindow
            Start-Process -FilePath "icacls" -ArgumentList "`"$Path`" /grant Administrators:F /t" -Wait -NoNewWindow
            Write-Host "Ownership and permissions granted for $Path" -ForegroundColor Green
        } catch {
            Write-Host "Failed to take ownership of $Path" -ForegroundColor Red
        }
    } else {
        Write-Host "Path not found: $Path" -ForegroundColor Yellow
    }
}

# Function to Remove Files or Folders
function Remove-ItemForce {
    param (
        [string]$Path
    )

    if (Test-Path $Path) {
        Write-Host "Removing $Path..." -ForegroundColor Yellow
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Removed: $Path" -ForegroundColor Green
        } catch {
            Write-Host "Failed to remove: $Path" -ForegroundColor Red
        }
    } else {
        Write-Host "Path not found: $Path" -ForegroundColor Yellow
    }
}

# Function to Remove Appx Packages
function Remove-Appx {
    param (
        [string]$AppName
    )

    Write-Host "Removing Appx packages for $AppName..." -ForegroundColor Yellow
    Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*$AppName*" } | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$AppName*" } | ForEach-Object {
        Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue
        Write-Host "Removed provisioned package: $($_.DisplayName)" -ForegroundColor Green
    }
}

# Function to Clean Registry
function Remove-Registry {
    param (
        [array]$RegistryKeys
    )

    foreach ($key in $RegistryKeys) {
        if (Test-Path $key) {
            try {
                Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Removed registry key: $key" -ForegroundColor Green
            } catch {
                Write-Host "Failed to remove registry key: $key" -ForegroundColor Red
            }
        }
    }
}

# Main Script Logic

# 1. Remove Microsoft Teams Appx Packages
Write-Host "Removing New Microsoft Teams Appx packages..." -ForegroundColor Green
Remove-Appx -AppName "Teams"

# 2. Remove New Outlook Appx Packages
Write-Host "Removing New Outlook Appx packages..." -ForegroundColor Green
Remove-Appx -AppName "Outlook"

# 3. Remove Teams Folders in WindowsApps
$windowsAppsPath = "C:\Program Files\WindowsApps"
$teamsFolders = Get-ChildItem -Path $windowsAppsPath -Filter "MSTeams_*" -ErrorAction SilentlyContinue
foreach ($folder in $teamsFolders) {
    if (Test-Path $folder.FullName) {
        Grant-WindowsAppsAccess -Path $folder.FullName
        Remove-ItemForce -Path $folder.FullName
    }
}

# 4. Remove Specific Files in WindowsApps
$teamsFiles = @(
    "$windowsAppsPath\MSTeams_*_x64__*\msteams_autostarter",
    "$windowsAppsPath\MSTeams_*_x64__*\msteams_canary"
)
foreach ($filePath in $teamsFiles) {
    if (Test-Path $filePath) {
        Grant-WindowsAppsAccess -Path $filePath
        Remove-ItemForce -Path $filePath
    }
}

# 5. Remove User-Specific Teams Cache and Data
Write-Host "Removing Teams residual files in user profiles..." -ForegroundColor Yellow
$profiles = Get-WmiObject Win32_UserProfile | Where-Object { $_.Special -eq $false }
foreach ($profile in $profiles) {
    $appDataPath = Join-Path -Path $profile.LocalPath -ChildPath "AppData\Local\Packages"
    $teamsCache = Get-ChildItem -Path $appDataPath -Filter "*Teams*" -Recurse -ErrorAction SilentlyContinue
    foreach ($cacheItem in $teamsCache) {
        Remove-ItemForce -Path $cacheItem.FullName
    }
}

# 6. Remove Teams-Related Registry Keys
Write-Host "Removing Teams-related registry keys..." -ForegroundColor Yellow
$teamsRegistryKeys = @(
    "HKCU:\Software\Microsoft\Office\Teams",
    "HKCU:\Software\Microsoft\Teams",
    "HKLM:\Software\Microsoft\Office\Teams",
    "HKLM:\Software\Microsoft\Teams"
)
Remove-Registry -RegistryKeys $teamsRegistryKeys

# 7. Block Reinstallation of Unified Teams and New Outlook
Write-Host "Blocking Unified Teams and New Outlook from reinstallation..." -ForegroundColor Green

# Block Unified Teams
New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\ApplicationManagement" -Name "PreventUnifiedTeamsInstall" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\ApplicationManagement\PreventUnifiedTeamsInstall" -Name "Value" -Value 1

# Block New Outlook
New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\ApplicationManagement" -Name "PreventNewOutlookInstall" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\ApplicationManagement\PreventNewOutlookInstall" -Name "Value" -Value 1

# 8. Final Detection for Remnants
Write-Host "Checking for any remaining Teams files or folders..." -ForegroundColor Yellow
$remainingTeamsFolders = Get-ChildItem -Path $windowsAppsPath -Filter "MSTeams_*" -ErrorAction SilentlyContinue

if ($remainingTeamsFolders.Count -eq 0) {
    Write-Host "No Microsoft Teams remnants found in WindowsApps." -ForegroundColor Green
} else {
    Write-Host "Remaining Teams folders or remnants detected:" -ForegroundColor Red
    $remainingTeamsFolders | ForEach-Object { Write-Host $_.FullName }
}

# 9. Create Marker File to Prevent Re-Execution
Write-Host "Creating marker file to prevent re-execution..." -ForegroundColor Yellow
New-Item -ItemType File -Path $MarkerFile -Force | Out-Null
Write-Host "Marker file created: $MarkerFile" -ForegroundColor Green

Write-Host "Script execution completed." -ForegroundColor Green
exit 0
