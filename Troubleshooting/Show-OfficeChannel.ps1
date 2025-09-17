Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" |
   Select-Object CDNBaseUrl, UpdateChannel, VersionToReport | ogv
