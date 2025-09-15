$appId = "1ddb4e76-6d86-4804-a715-36e60b237541"

$intuneLogList = Get-ChildItem -Path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs" -Filter "appworkload*.log" -File | sort LastWriteTime -Descending | select -ExpandProperty FullName

if (!$intuneLogList) {
Write-Error "Unable to find any Intune log files. Redeploy will probably not work as expected."
return
}

foreach ($intuneLog in $intuneLogList) {
$appMatch = Select-String -Path $intuneLog -Pattern "\[Win32App\]\[GRSManager\] App with id: $appId is not expired." -Context 0, 1
if ($appMatch) {
foreach ($match in $appMatch) {
$Hash = “”
$LineNumber = 0
$LineNumber = $match.LineNumber
$Hash = Get-Content $intuneLog | Select-Object -Skip $LineNumber -First 1
if ($hash) {
$hash = $hash.Replace('+','\+')
return $hash
}
}
}
}
