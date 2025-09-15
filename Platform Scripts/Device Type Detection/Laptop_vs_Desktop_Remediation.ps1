# Remediation.ps1
# Updates Entra device extensionAttributes:
#   Laptops  -> ext7="Laptop",  ext8=""
#   Desktops -> ext7="",        ext8="Desktop"

$ErrorActionPreference = 'Stop'

# ====== CONFIG ======
$TenantId     = '<YOUR_TENANT_ID>'
$ClientId     = '<YOUR_APP_CLIENT_ID>'
$ClientSecret = '<YOUR_APP_CLIENT_SECRET>'   # Prefer secure retrieval (Key Vault, etc.)
# ====================

function Test-IsLaptop {
    try {
        $enc = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop
        if ($enc.ChassisTypes -match '^(8|9|10|14)$') { return $true }
    } catch {}

    try {
        $bat = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($bat) { return $true }
    } catch {}

    return $false
}

function Get-AzureAdDeviceId {
    $out = dsregcmd /status | Out-String
    if ($out -match 'AzureAdDeviceId\s*:\s*([0-9a-fA-F-]{36})') { return $Matches[1] }
    return $null
}

$IsLaptop = Test-IsLaptop

# 1) Get device ID
$aadDeviceId = Get-AzureAdDeviceId
if (-not $aadDeviceId) {
    Write-Warning "No AzureAdDeviceId (device not AADJ/HAADJ?); skipping Entra update."
    exit 0
}

# 2) Acquire token
$tokenResp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
    client_id     = $ClientId
    scope         = 'https://graph.microsoft.com/.default'
    client_secret = $ClientSecret
    grant_type    = 'client_credentials'
}
$accessToken = $tokenResp.access_token
if (-not $accessToken) {
    Write-Warning "Token acquisition failed; skipping Entra update."
    exit 0
}
$headers = @{ Authorization = "Bearer $accessToken" }

# 3) Resolve Entra device object
$dev = Invoke-RestMethod -Headers $headers -Method Get -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadDeviceId'"
if (-not $dev.value -or $dev.value.Count -eq 0) {
    Write-Warning "No Graph device for deviceId=$aadDeviceId; skipping."
    exit 0
}
$deviceObjectId = $dev.value[0].id

# 4) Decide extension attribute values
if ($IsLaptop) {
    $ext7 = 'Laptop'
    $ext8 = ''         # clear desktop flag
} else {
    $ext7 = ''         # clear laptop flag
    $ext8 = 'Desktop'
}

$body = @{
    extensionAttributes = @{
        extensionAttribute7 = $ext7
        extensionAttribute8 = $ext8
    }
} | ConvertTo-Json -Depth 5

# 5) Patch Entra device
Invoke-RestMethod -Headers $headers -Method Patch -Uri "https://graph.microsoft.com/v1.0/devices/$deviceObjectId" -Body $body -ContentType 'application/json'
Write-Output ("Updated: extensionAttribute7='{0}', extensionAttribute8='{1}' on device {2}" -f $ext7, $ext8, $deviceObjectId)

exit 0
