Get-AppxPackage -AllUsers Microsoft.CompanyPortal | Foreach {
    Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppxManifest.xml"
}
