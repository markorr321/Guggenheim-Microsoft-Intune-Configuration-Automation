<#
.SYNOPSIS
    Sets the Windows time zone automatically based on the public IP address using IANA to Windows time zone conversion.

.DESCRIPTION
    This script queries http://ipinfo.io/json to determine the device's geographic IANA time zone, 
    maps it to the correct Windows time zone using a custom XML source, and sets the local time zone accordingly.

.NOTES
    Author:        Florian Salzmann (http://elFlorian.ch)
    Version:       1.1
    Last Updated:  2025-09-11
    Run As:        SYSTEM (not user)
    Architecture:  Must run in 64-bit context

#>

$PackageName = "Set-TimeZoneByIPAddress"
Start-Transcript -Path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$PackageName-script.log" -Force -Append

try {
    # --- Registry marker for versioning ---
    $Path = "HKLM:\SOFTWARE\scloud\Set-TimeZoneByIPAddress"
    $Key = "Version"
    $KeyFormat = "DWord"
    $Value = 1

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    if (-not (Get-ItemProperty -Path $Path -Name $Key -ErrorAction SilentlyContinue)) {
        # Create new property if missing
        New-ItemProperty -Path $Path -Name $Key -Value $Value -PropertyType $KeyFormat -Force | Out-Null
    } else {
        # Update existing property
        Set-ItemProperty -Path $Path -Name $Key -Value $Value
    }

    Write-Output "Registry key [$Path] set: $Key = $Value ($KeyFormat)"

    # --- Time zone detection and mapping ---
    Write-Output "Fetching IANA time zone from ipinfo.io..."
    $ianaTz = (Invoke-RestMethod -Uri "http://ipinfo.io/json").timezone
    if (-not $ianaTz) {
        throw "Could not retrieve IANA time zone from ipinfo.io."
    }
    Write-Output "Detected IANA Time Zone: $ianaTz"

    Write-Output "Downloading custom XML mapping..."
    $xmlUrl = "https://raw.githubusercontent.com/FlorianSLZ/scloud/refs/heads/main/scripts/Set-TimeZoneByIPAddress/windowsZones.xml"
    [xml]$windowsZones = Invoke-RestMethod -Uri $xmlUrl
    if (-not $windowsZones) {
        throw "Failed to download or parse the XML mapping file."
    }

    Write-Output "Searching for matching mapping..."
    $mapping = $windowsZones.supplementalData.windowsZones.mapTimezones.mapZone | Where-Object {
        $_.type -split ' ' -contains $ianaTz
    }

    if (-not $mapping) {
        throw "No mapping found for IANA time zone: $ianaTz"
    }

    $windowsTZ = $mapping.other | Select-Object -First 1
    Write-Output "Mapped to Windows Time Zone: $windowsTZ"

    try {
        Write-Output "Setting Windows time zone using Set-TimeZone..."
        Set-TimeZone -Id $windowsTZ
        Write-Output "Successfully set Windows Time Zone: $windowsTZ"
    } catch {
        Write-Error "Set-TimeZone failed. $_"
    }

} catch {
    Write-Error "Failed to set Windows time zone: $_"
}

Stop-Transcript
