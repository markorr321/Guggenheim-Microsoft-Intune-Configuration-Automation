# Map Office channel URL → friendly name
$ChannelMap = @{
  'http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise'
  'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current'
  'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current (Preview)'
  'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise'
  'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise (Preview)'
  'http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta'
}

$conf = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue

# Prefer UpdateChannel; fall back to CDNBaseUrl if UpdateChannel is empty
$rawChannel = if ($conf.UpdateChannel) { $conf.UpdateChannel } else { $conf.CDNBaseUrl }
$channelName = $ChannelMap[$rawChannel]
if (-not $channelName) { $channelName = 'Unknown (not in map)' }

[pscustomobject]@{
  ProductReleaseIds = $conf.ProductReleaseIds
  CDNBaseUrl        = $conf.CDNBaseUrl
  UpdateChannel     = $conf.UpdateChannel
  ChannelName       = $channelName
  VersionToReport   = $conf.VersionToReport
} | Format-List
