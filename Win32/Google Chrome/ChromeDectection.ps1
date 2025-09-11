# =========================
# Intune Detection: Google Chrome minimum version across MSI, per-machine EXE, per-user EXE
# Exit 0 = Chrome is installed AND version >= RequiredVersion
# Exit 1 = Not installed or version < RequiredVersion
# =========================

# >>> EDIT THESE TWO LINES <<<
$RequiredVersion = "140.0.7339.81"                          # baseline you want to enforce
$MsiProductCode  = "{66793499-0B07-380D-8FDE-467BB6263225}" # MSI ProductCode for your packaged Chrome (optional but recommended)

function As-Version($s) {
    try { return [version]$s } catch { return $null }
}

$req = As-Version $RequiredVersion
if (-not $req) {
    Write-Output "Invalid RequiredVersion: $RequiredVersion"
    exit 1
}

$foundVersions = @()

function Add-FindResult($versionString, $source) {
    $v = As-Version $versionString
    if ($v) {
        $script:foundVersions += [pscustomobject]@{ Version=$v; Source=$source }
    }
}

# --- 1) MSI install check (enterprise installs) ---
if ($MsiProductCode) {
    $uninstallRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$MsiProductCode",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$MsiProductCode"
    )
    foreach ($root in $uninstallRoots) {
        if (Test-Path $root) {
            try {
                $p = Get-ItemProperty $root
                if ($p.DisplayVersion) {
                    Add-FindResult -versionString $p.DisplayVersion -source "MSI:$root"
                }
            } catch {}
        }
    }
}

# --- 2) Per-machine EXE via Google Update (HKLM) ---
# Stable channel AppID:
$ChromeStableAppId = "{8A69D345-D564-463C-AFF1-A69D9E530F96}"
$clientsKey = "HKLM:\SOFTWARE\Google\Update\Clients\$ChromeStableAppId"
if (Test-Path $clientsKey) {
    try {
        $pv = (Get-ItemProperty $clientsKey).pv  # product version
        if ($pv) { Add-FindResult $pv "HKLM Update\Clients pv" }
    } catch {}
}

# --- 3) Per-user EXE via BLBeacon (HKU for all loaded user hives) ---
# Intune detection runs as SYSTEM, so enumerate HKEY_USERS SIDs to simulate HKCU for each user
$sidExclude = @(
    "S-1-5-18", "S-1-5-19", "S-1-5-20"  # LocalSystem, LocalService, NetworkService
)

try {
    $hku = Get-ChildItem "Registry::HKEY_USERS" -ErrorAction SilentlyContinue
    foreach ($sid in $hku) {
        if ($sid.Name -match "S-1-5-21-" -and -not ($sidExclude -contains ($sid.PSChildName))) {
            $bl = "Registry::HKEY_USERS\$($sid.PSChildName)\Software\Google\Chrome\BLBeacon"
            if (Test-Path $bl) {
                try {
                    $ver = (Get-ItemProperty $bl).version
                    if ($ver) { Add-FindResult $ver "HKU:\$($sid.PSChildName)\Chrome\BLBeacon" }
                } catch {}
            }
        }
    }
} catch {}

# --- 4) File version checks (per-machine and common per-user locations) ---
$exePaths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
)

# Add per-user LocalAppData chrome.exe for each profile (best-effort)
try {
    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -notin @("All Users","Default","Default User","Public","WDAGUtilityAccount") -and
        -not $_.Attributes.ToString().Contains("ReparsePoint")
    }
    foreach ($profile in $userProfiles) {
        $exePaths += Join-Path $profile.FullName "AppData\Local\Google\Chrome\Application\chrome.exe"
    }
} catch {}

foreach ($exe in $exePaths | Select-Object -Unique) {
    if (Test-Path $exe) {
        try {
            $fv = (Get-Item $exe).VersionInfo.FileVersion
            if ($fv) { Add-FindResult $fv "File:$exe" }
        } catch {}
    }
}

# --- Decide result ---
if ($foundVersions.Count -gt 0) {
    $best = $foundVersions | Sort-Object Version -Descending | Select-Object -First 1
    if ($best.Version -ge $req) {
        Write-Output ("Chrome detected {0} (>= {1}) via {2}" -f $best.Version, $req, $best.Source)
        exit 0
    } else {
        Write-Output ("Chrome detected {0} (< {1}); sources: {2}" -f $best.Version, $req,
            ($foundVersions | ForEach-Object { "$($_.Version) from $($_.Source)" } -join "; "))
        exit 1
    }
} else {
    Write-Output "Chrome not detected at any known MSI/registry/file locations"
    exit 1
}
