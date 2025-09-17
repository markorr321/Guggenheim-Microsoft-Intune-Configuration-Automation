# Critical Microsoft Store endpoints (from Microsoft docs)
$Endpoints = @(
    "displaycatalog.mp.microsoft.com",   # Store catalog
    "purchase.md.mp.microsoft.com",      # Purchase API
    "licensing.mp.microsoft.com",        # App licensing
    "storeedgefd.dsx.mp.microsoft.com",  # Store metadata/CDN
    "cdn.storeedgefd.dsx.mp.microsoft.com" # Microsoft-hosted Win32 app fallback cache
)

Write-Host "Testing connectivity to Microsoft Store endpoints (Ports 80 & 443)..." -ForegroundColor Cyan

$Results = foreach ($ep in $Endpoints) {
    foreach ($port in 80,443) {
        try {
            $test = Test-NetConnection -ComputerName $ep -Port $port -WarningAction SilentlyContinue
            [PSCustomObject]@{
                Endpoint  = $ep
                Port      = $port
                Reachable = $test.TcpTestSucceeded
            }
        }
        catch {
            [PSCustomObject]@{
                Endpoint  = $ep
                Port      = $port
                Reachable = $false
            }
        }
    }
}

$Results | Format-Table -AutoSize
