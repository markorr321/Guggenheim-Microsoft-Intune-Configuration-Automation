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
    try {
        $msi = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$MsiProductCode" -ErrorAction Stop
        if ($msi.DisplayVersion) { Add-FindResult $msi.DisplayVersion "MSI" }
    } catch {}
    try {
        $msiWow = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$MsiProductCode" -ErrorAction Stop
        if ($msiWow.DisplayVersion) { Add-FindResult $msiWow.DisplayVersion "MSI-WOW6432Node" }
    } catch {}
}

# --- 2) Registry-based version checks (all install types) ---
$regPaths = @(
    "HKLM:\SOFTWARE\Google\Update\Clients",
    "HKLM:\SOFTWARE\WOW6432Node\Google\Update\Clients",
    "HKCU:\SOFTWARE\Google\Update\Clients"
)

foreach ($rp in $regPaths) {
    try {
        Get-ChildItem $rp -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $dv = (Get-ItemProperty $_.PsPath -ErrorAction Stop).pv
                if ($dv) { Add-FindResult $dv "Reg:$rp\$($_.PSChildName)" }
            } catch {}
        }
    } catch {}
}

# --- 3) File system checks for chrome.exe (machine & user) ---
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
        $sources = ($foundVersions | ForEach-Object { "$($_.Version) from $($_.Source)" }) -join "; "
        Write-Output ("Chrome detected {0} (< {1}); sources: {2}" -f $best.Version, $req, $sources)
        exit 1
    }
} else {
    Write-Output "Chrome not detected at any known MSI/registry/file locations"
    exit 1
}
