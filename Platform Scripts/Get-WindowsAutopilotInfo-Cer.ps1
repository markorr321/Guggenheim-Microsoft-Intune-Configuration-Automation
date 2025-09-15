[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
 
Install-PackageProvider -Name Nuget -MinimumVersion 2.8.5.201 -Force
Install-Module -Name Microsoft.Graph -Force
Install-Script -Name Get-WindowsAutopilotInfo -Force
# Define tenant info and app registration
$tenantId = "b5a5c657-a4c9-4b11-9b2a-a71eb7ca9433"
$appId = "16176c07-b2b2-4ad7-a4e4-82a8430cf553"
 
# Thumbprint of certificate (from cert store)
$certThumbprint = "13D20512C83E0B1EEEB02DA0297BF82D4C292A48"
 
$password = ConvertTo-SecureString -String "ENTER Passwrod" -Force -AsPlainText
$LoadCertificate = Import-PfxCertificate -FilePath "D:\autopilot\Guggenheim Securities intune app reg.pfx" -CertStoreLocation Cert:\LocalMachine\My -Password $password
 
#$cert = Get-ChildItem -Path Cert:\LocalMachine\My\$certThumbprint
Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $certThumbprint -NoWelcome
 
Get-WindowsAutopilotInfo -Online -Grouptag Hybrid -appId $appId -TenantId $tenantId
