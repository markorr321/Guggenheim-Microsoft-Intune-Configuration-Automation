<# 
Export-ConditionalAccessPolicies.ps1
- Exports all Entra ID Conditional Access policies to JSON
- Requires: Graph delegated permission Policy.Read.All
- Works in PowerShell 7+
#>

param(
  [string]$OutRoot = 'C:\reports\CA',
  [string]$TenantId = ''   # optional: set if you want to force a tenant
)

# ---- Prep output folder (dated) ----
$stamp   = (Get-Date).ToString('yyyy-MM-dd')
$OutDir  = Join-Path $OutRoot $stamp
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

Write-Host "Output: $OutDir" -ForegroundColor Cyan

# ---- Connect to Microsoft Graph ----
try {
  # Check if already connected with required scope
  $ctx = Get-MgContext
  if ($ctx -and $ctx.Scopes -contains 'Policy.Read.All') {
    Write-Host "Already connected to Microsoft Graph with required permissions." -ForegroundColor Green
    Write-Host "Connected as $($ctx.Account) in tenant $($ctx.TenantId)" -ForegroundColor Green
  } else {
    # Connect with required scope
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    if ([string]::IsNullOrWhiteSpace($TenantId)) {
      Connect-MgGraph -Scopes "Policy.Read.All" -NoWelcome
    } else {
      Connect-MgGraph -TenantId $TenantId -Scopes "Policy.Read.All" -NoWelcome
    }
    
    # Verify connection and scope
    $ctx = Get-MgContext
    if (-not $ctx -or $ctx.Scopes -notcontains 'Policy.Read.All') {
      throw "Failed to connect with Policy.Read.All scope. Please ensure you have the required permissions."
    }
    Write-Host "Connected as $($ctx.Account) in tenant $($ctx.TenantId)" -ForegroundColor Green
  }
} catch {
  throw "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
}

# ---- Helper: safe filename from display name ----
function Get-SafeName([string]$name, [string]$id) {
  $safe = ($name -replace '[^\w\- ]','_').Trim()
  if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "Policy_$id" }
  # keep filenames short-ish; append 8 chars of id to avoid collisions
  $suffix = $id -replace '[^0-9a-fA-F-]',''
  $suffix = $suffix.Substring(0,[Math]::Min(8,$suffix.Length))
  return "{0}__{1}.json" -f $safe, $suffix
}

# ---- Fetch all CA policies ----
Write-Host "Fetching Conditional Access policies..." -ForegroundColor Cyan
$policies = @()
try {
  $policies = Get-MgIdentityConditionalAccessPolicy -All
} catch {
  throw "Get-MgIdentityConditionalAccessPolicy failed: $($_.Exception.Message)"
}

if (-not $policies) {
  Write-Warning "No Conditional Access policies were returned. (Role/scopes/tenant?)"
  return
}

# ---- Export each policy to its own JSON ----
$index = @()
$i = 0
foreach ($p in $policies) {
  $i++
  $file = Join-Path $OutDir (Get-SafeName -name $p.DisplayName -id $p.Id)
  try {
    # Re-get by ID to ensure full object & avoid any paging nuances
    $full = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $p.Id
    $full | ConvertTo-Json -Depth 100 | Set-Content -Path $file -Encoding UTF8

    $index += [PSCustomObject]@{
      DisplayName = $p.DisplayName
      Id          = $p.Id
      State       = $p.State
      File        = $file
    }
    Write-Host ("[{0}/{1}] {2}" -f $i,$policies.Count,$p.DisplayName)
  } catch {
    Write-Warning ("Failed to export '{0}' ({1}): {2}" -f $p.DisplayName,$p.Id,$_.Exception.Message)
  }
}

# ---- Write combined JSON & CSV index ----
$combinedPath = Join-Path $OutDir 'all-policies.json'
$indexCsv     = Join-Path $OutDir 'index.csv'

$policies | ConvertTo-Json -Depth 100 | Set-Content -Path $combinedPath -Encoding UTF8
$index    | Sort-Object DisplayName | Export-Csv -NoTypeInformation -Path $indexCsv -Encoding UTF8

Write-Host "`nExport complete." -ForegroundColor Green
Write-Host "Per-policy JSON:   $OutDir\*.json"
Write-Host "Combined JSON:     $combinedPath"
Write-Host "Index CSV:         $indexCsv"
