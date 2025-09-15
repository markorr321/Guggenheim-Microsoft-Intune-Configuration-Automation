# Detection.ps1
# Always trigger Remediation (exit 1), while logging what we detect.

function Test-IsLaptop {
    try {
        $enc = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop
        if ($enc.ChassisTypes -match '^(8|9|10|14)$') { return $true }  # 8=Portable, 9=Laptop, 10=Notebook, 14=Sub-Notebook
    } catch {}

    try {
        $bat = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($bat) { return $true }
    } catch {}

    return $false
}

$IsLaptop = Test-IsLaptop
Write-Output ("Detected: {0}" -f ($IsLaptop ? "Laptop" : "Desktop"))
exit 1  # Force Remediation on all targeted devices
