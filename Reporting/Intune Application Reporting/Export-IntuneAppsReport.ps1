# =========================================
# Script: Intune Apps Report with Group Assignments & Membership
# Author: Mark Orr
# Date:   09/12/2025
#
# Description:
# Generates a comprehensive Intune apps report with group assignments and detailed group membership.
# Features interactive Out-GridView selection and exports to Excel workbook with multiple sheets:
# - Main app assignments data with columns: AppName, Id, AppType, Platform, Publisher,
#   CreatedDateTime, LastModifiedDateTime, TargetType, GroupName, GroupMode
# - Group Summary sheet showing member counts and app assignments per group
# - Individual sheets for each group displaying all members (users/devices) with details
#
# Requirements: Microsoft.Graph and ImportExcel PowerShell modules (auto-installed)
# =========================================

$ShowGrid   = $true           # set $false to skip OGV
$DoExport   = $false          # set $true to also export a single CSV
$ExportExcel = $true          # set $true to export Excel workbook with group member sheets
$ExportGroupCSVs = $false     # set $true to export separate CSV files for each group
$OutputFile = Join-Path (Get-Location) 'IntuneAppsAndAssignments.csv'
$ExcelFile  = Join-Path (Get-Location) 'IntuneAppsAndAssignments.xlsx'
$GroupCSVFolder = Join-Path (Get-Location) 'GroupMembers'

# Ensure Graph SDK & ImportExcel module & connect
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
if ($ExportExcel -and -not (Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -SkipPublisherCheck
    } catch {
        Write-Host "Failed to install ImportExcel module. Trying alternative method..." -ForegroundColor Yellow
        try {
            Install-Module ImportExcel -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
        } catch {
            Write-Host "Could not install ImportExcel module. Disabling Excel export." -ForegroundColor Red
            $ExportExcel = $false
        }
    }
}
Connect-MgGraph -NoWelcome -Scopes "DeviceManagementApps.Read.All","Group.Read.All","GroupMember.Read.All","User.Read.All","Device.Read.All"

# Helpers
function Get-AnyProp {
    param($obj, [string[]]$Keys)
    foreach ($k in $Keys) {
        if ($null -ne $obj.$k -and $obj.$k -ne '') { return $obj.$k }
        if ($obj.AdditionalProperties -and $obj.AdditionalProperties.ContainsKey($k)) {
            $v = $obj.AdditionalProperties[$k]
            if ($null -ne $v -and $v -ne '') { return $v }
        }
    }
    return $null
}
function Get-PlatformTypeFromOData {
    param([string]$odataType)
    $t = ($odataType -replace '^#microsoft\.graph\.', '')
    $platform = switch -regex ($t) {
        'win32|windows' { 'Windows' }
        'macOS'         { 'macOS' }
        'ios'           { 'iOS/iPadOS' }
        'android'       { 'Android' }
        'webApp'        { 'Web' }
        default         { 'Unknown' }
    }
    [pscustomobject]@{ Platform = $platform; AppType = $t }
}
$GroupNameCache = @{}
function Resolve-GroupName {
    param([string]$GroupId)
    if ([string]::IsNullOrWhiteSpace($GroupId)) { return $null }
    if ($GroupNameCache.ContainsKey($GroupId)) { return $GroupNameCache[$GroupId] }
    try {
        $g = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        $GroupNameCache[$GroupId] = $g.DisplayName
        return $g.DisplayName
    } catch {
        $GroupNameCache[$GroupId] = "(group not found: $GroupId)"
        return $GroupNameCache[$GroupId]
    }
}

function Show-GroupMembers {
    param([string]$GroupName, [array]$AllRows)
    
    # Find the group ID from our data
    $groupRow = $AllRows | Where-Object { $_.GroupName -eq $GroupName -and $_.TargetType -eq 'Group' } | Select-Object -First 1
    if (-not $groupRow) {
        Write-Host "Group '$GroupName' not found in the current data." -ForegroundColor Red
        return
    }
    
    # Get the group ID by reverse lookup
    $groupId = $null
    foreach ($kvp in $GroupNameCache.GetEnumerator()) {
        if ($kvp.Value -eq $GroupName) {
            $groupId = $kvp.Key
            break
        }
    }
    
    if (-not $groupId) {
        Write-Host "Could not find Group ID for '$GroupName'." -ForegroundColor Red
        return
    }
    
    try {
        Write-Host "Fetching members for group: $GroupName" -ForegroundColor Cyan
        $members = Get-MgGroupMember -GroupId $groupId -All
        
        if (-not $members) {
            Write-Host "No members found in group '$GroupName'." -ForegroundColor Yellow
            return
        }
        
        $memberDetails = foreach ($member in $members) {
            $memberType = ($member.AdditionalProperties['@odata.type'] -replace '^#microsoft\.graph\.', '')
            
            switch ($memberType) {
                'user' {
                    [pscustomobject]@{
                        Type = 'User'
                        Name = $member.AdditionalProperties.displayName
                        Email = $member.AdditionalProperties.userPrincipalName
                        Id = $member.Id
                    }
                }
                'device' {
                    [pscustomobject]@{
                        Type = 'Device'
                        Name = $member.AdditionalProperties.displayName
                        Email = 'N/A'
                        Id = $member.Id
                    }
                }
                default {
                    [pscustomobject]@{
                        Type = $memberType
                        Name = $member.AdditionalProperties.displayName
                        Email = 'N/A'
                        Id = $member.Id
                    }
                }
            }
        }
        
        $memberDetails | Sort-Object Type, Name | Out-GridView -Title "Members of Group: $GroupName" -Wait
        
    } catch {
        Write-Host "Error retrieving members for group '$GroupName': $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-GroupMembersData {
    param([string]$GroupName)
    
    # Get the group ID by reverse lookup
    $groupId = $null
    foreach ($kvp in $GroupNameCache.GetEnumerator()) {
        if ($kvp.Value -eq $GroupName) {
            $groupId = $kvp.Key
            break
        }
    }
    
    if (-not $groupId) {
        return @()
    }
    
    try {
        $members = Get-MgGroupMember -GroupId $groupId -All -ErrorAction SilentlyContinue
        
        if (-not $members) {
            return @()
        }
        
        $memberDetails = foreach ($member in $members) {
            $memberType = ($member.AdditionalProperties['@odata.type'] -replace '^#microsoft\.graph\.', '')
            
            switch ($memberType) {
                'user' {
                    [pscustomobject]@{
                        Type = 'User'
                        Name = $member.AdditionalProperties.displayName
                        Email = $member.AdditionalProperties.userPrincipalName
                        Id = $member.Id
                    }
                }
                'device' {
                    [pscustomobject]@{
                        Type = 'Device'
                        Name = $member.AdditionalProperties.displayName
                        Email = 'N/A'
                        Id = $member.Id
                    }
                }
                default {
                    [pscustomobject]@{
                        Type = $memberType
                        Name = $member.AdditionalProperties.displayName
                        Email = 'N/A'
                        Id = $member.Id
                    }
                }
            }
        }
        
        return ($memberDetails | Sort-Object Type, Name)
        
    } catch {
        return @()
    }
}

# Fetch apps
Write-Host "Fetching Intune mobile apps..." -ForegroundColor Cyan
$apps = Get-MgDeviceAppManagementMobileApp -All

# Build ONE unified table (one row per assignment; unassigned => single (none) row)
Write-Host "Building unified rows..." -ForegroundColor Cyan
$rows = foreach ($app in $apps) {
    $odata = Get-AnyProp $app '@odata.type'
    $meta  = Get-PlatformTypeFromOData $odata

    $assignments = Get-MgDeviceAppManagementMobileAppAssignment -MobileAppId $app.Id -ErrorAction SilentlyContinue

    if (-not $assignments) {
        [pscustomobject]@{
            AppName              = $app.DisplayName
            Id                   = $app.Id
            AppType              = $meta.AppType
            Platform             = $meta.Platform
            Publisher            = $app.Publisher
            CreatedDateTime      = $app.CreatedDateTime
            LastModifiedDateTime = $app.LastModifiedDateTime
            TargetType           = '(none)'
            GroupName            = ''
            GroupMode            = ''
        }
        continue
    }

    foreach ($as in $assignments) {
        $target = $as.Target
        $tType  = (Get-AnyProp $target @('@odata.type')) -replace '^#microsoft\.graph\.', ''
        $targetType = switch ($tType) {
            'groupAssignmentTarget'             { 'Group' }
            'allDevicesAssignmentTarget'        { 'All Devices' }
            'allLicensedUsersAssignmentTarget'  { 'All Users' }
            default                             { $tType }
        }
        $groupId    = if ($targetType -eq 'Group') { Get-AnyProp $target @('groupId') } else { $null }
        $targetName = switch ($targetType) {
            'Group'       { Resolve-GroupName $groupId }
            'All Users'   { 'All Users' }
            'All Devices' { 'All Devices' }
            default       { '' }
        }

        [pscustomobject]@{
            AppName              = $app.DisplayName
            Id                   = $app.Id
            AppType              = $meta.AppType
            Platform             = $meta.Platform
            Publisher            = $app.Publisher
            CreatedDateTime      = $app.CreatedDateTime
            LastModifiedDateTime = $app.LastModifiedDateTime
            TargetType           = $targetType
            GroupName            = $targetName
            GroupMode            = $as.Intent
        }
    }
}

# Show ONE OGV of the final data (exact columns/order) with selection capability
$selectedRows = $null
if ($ShowGrid) {
    Write-Host "Select the rows you want to export in the grid view, then click OK..." -ForegroundColor Yellow
    $selectedRows = $rows |
      Select-Object AppName, Id, AppType, Platform, Publisher, CreatedDateTime, LastModifiedDateTime, TargetType, GroupName, GroupMode |
      Sort-Object AppName, TargetType, GroupName |
      Out-GridView -Title 'Intune Apps & Assignments - SELECT ROWS TO EXPORT' -PassThru
    
    if ($selectedRows) {
        Write-Host "You selected $($selectedRows.Count) row(s) for export." -ForegroundColor Green
    } else {
        Write-Host "No rows were selected. Nothing will be exported." -ForegroundColor Yellow
    }
}

# Export selected rows to CSV (automatically enabled when rows are selected)
if ($selectedRows -and $selectedRows.Count -gt 0) {
    $selectedRows | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $($selectedRows.Count) selected rows to: $OutputFile" -ForegroundColor Green
} elseif ($DoExport -and -not $ShowGrid) {
    # Fallback: export all rows if $DoExport is true but no grid selection was made
    $rows |
      Select-Object AppName, Id, AppType, Platform, Publisher, CreatedDateTime, LastModifiedDateTime, TargetType, GroupName, GroupMode |
      Sort-Object AppName, TargetType, GroupName |
      Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Exported all $($rows.Count) rows to: $OutputFile" -ForegroundColor Green
}

# Export separate CSV files for each group
if ($ExportGroupCSVs -and $rows) {
    Write-Host "Creating separate CSV files for each group..." -ForegroundColor Cyan
    
    # Create GroupMembers folder if it doesn't exist
    if (-not (Test-Path $GroupCSVFolder)) {
        New-Item -ItemType Directory -Path $GroupCSVFolder -Force | Out-Null
    }
    
    # Get unique groups for creating CSV files
    $uniqueGroups = $rows | Where-Object { $_.TargetType -eq 'Group' -and $_.GroupName -ne '' } | 
                    Select-Object -ExpandProperty GroupName -Unique | Sort-Object
    
    # Create a summary file with group names and their CSV file paths
    $groupSummary = foreach ($groupName in $uniqueGroups) {
        $safeFileName = $groupName -replace '[\\\/\?\*\[\]<>|:"]', '_'
        $csvFileName = "GroupMembers_$safeFileName.csv"
        $csvPath = Join-Path $GroupCSVFolder $csvFileName
        
        Write-Host "  Creating CSV for group: $groupName" -ForegroundColor Gray
        
        $groupMembers = Get-GroupMembersData -GroupName $groupName
        
        if ($groupMembers -and $groupMembers.Count -gt 0) {
            $groupMembers | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            
            [pscustomobject]@{
                GroupName = $groupName
                MemberCount = $groupMembers.Count
                CSVFile = $csvFileName
                FullPath = $csvPath
            }
        } else {
            # Create empty CSV with message
            $emptyData = [pscustomobject]@{
                Type = "INFO"
                Name = "No members found in this group"
                Email = "This group has no members or could not be accessed"
                Id = ""
            }
            $emptyData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            
            [pscustomobject]@{
                GroupName = $groupName
                MemberCount = 0
                CSVFile = $csvFileName
                FullPath = $csvPath
            }
        }
    }
    
    # Export the group summary
    $summaryPath = Join-Path $GroupCSVFolder "GroupSummary.csv"
    $groupSummary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "Group member CSV files created in: $GroupCSVFolder" -ForegroundColor Green
    Write-Host "Group summary file: $summaryPath" -ForegroundColor Yellow
    Write-Host "To view group members, open the corresponding CSV file from the GroupMembers folder." -ForegroundColor Cyan
}

# Export to Excel workbook with group member sheets
if ($ExportExcel -and $rows) {
    Write-Host "Creating Excel workbook with group member sheets..." -ForegroundColor Cyan
    
    # Remove existing Excel file if it exists
    if (Test-Path $ExcelFile) {
        Remove-Item $ExcelFile -Force
    }
    
    # Get unique groups for creating worksheets
    $uniqueGroups = $rows | Where-Object { $_.TargetType -eq 'Group' -and $_.GroupName -ne '' } | 
                    Select-Object -ExpandProperty GroupName -Unique | Sort-Object
    
    # Create main data sheet
    $mainData = if ($selectedRows -and $selectedRows.Count -gt 0) { $selectedRows } else { 
        $rows | Select-Object AppName, Id, AppType, Platform, Publisher, CreatedDateTime, LastModifiedDateTime, TargetType, GroupName, GroupMode |
        Sort-Object AppName, TargetType, GroupName 
    }
    
    # Create all sheets in one operation using a hashtable
    $allSheets = @{}
    
    # Add main data sheet
    $allSheets["App Assignments"] = $mainData
    
    # Create group summary data
    $groupSummaryData = foreach ($groupName in $uniqueGroups) {
        $groupMembers = Get-GroupMembersData -GroupName $groupName
        $memberCount = if ($groupMembers) { $groupMembers.Count } else { 0 }
        
        # Count how many apps are assigned to this group
        $appCount = ($mainData | Where-Object { $_.GroupName -eq $groupName }).Count
        
        [pscustomobject]@{
            GroupName = $groupName
            MemberCount = $memberCount
            AssignedApps = $appCount
            SheetName = ($groupName -replace '[\\\/\?\*\[\]<>|:"]', '_').Substring(0, [Math]::Min(31, ($groupName -replace '[\\\/\?\*\[\]<>|:"]', '_').Length))
        }
    }
    
    # Add group summary sheet
    $allSheets["Group Summary"] = $groupSummaryData
    
    # Add group member sheets
    foreach ($groupName in $uniqueGroups) {
        Write-Host "  Preparing sheet for group: $groupName" -ForegroundColor Gray
        
        $safeSheetName = $groupName -replace '[\\\/\?\*\[\]<>|:"]', '_'
        if ($safeSheetName.Length -gt 31) {
            $safeSheetName = $safeSheetName.Substring(0, 31)
        }
        
        $groupMembers = Get-GroupMembersData -GroupName $groupName
        
        if ($groupMembers -and $groupMembers.Count -gt 0) {
            # Add header row with group info
            $headerInfo = [pscustomobject]@{
                Type = "GROUP INFO"
                Name = $groupName
                Email = "Members: $($groupMembers.Count)"
                Id = ""
            }
            
            $sheetData = @($headerInfo) + @($groupMembers)
            $allSheets[$safeSheetName] = $sheetData
        } else {
            # Create empty sheet with message
            $emptyData = [pscustomobject]@{
                Type = "INFO"
                Name = "No members found"
                Email = "This group has no members or could not be accessed"
                Id = ""
            }
            $allSheets[$safeSheetName] = @($emptyData)
        }
    }
    
    # Create workbook with sheets in correct order
    # 1. App Assignments sheet first
    $allSheets["App Assignments"] | Export-Excel -Path $ExcelFile -WorksheetName "App Assignments" -AutoSize -FreezeTopRow -BoldTopRow
    
    # 2. Group Summary sheet second
    $allSheets["Group Summary"] | Export-Excel -Path $ExcelFile -WorksheetName "Group Summary" -AutoSize -FreezeTopRow -BoldTopRow -Append
    
    # 3. Individual group sheets
    foreach ($sheetName in $allSheets.Keys) {
        if ($sheetName -ne "App Assignments" -and $sheetName -ne "Group Summary") {
            $allSheets[$sheetName] | Export-Excel -Path $ExcelFile -WorksheetName $sheetName -AutoSize -FreezeTopRow -BoldTopRow -Append
        }
    }
    
    Write-Host "Excel workbook created: $ExcelFile" -ForegroundColor Green
    Write-Host "The workbook contains:" -ForegroundColor Yellow
    Write-Host "  - 'App Assignments' sheet with your main data" -ForegroundColor Yellow
    Write-Host "  - 'Group Summary' sheet with group overview" -ForegroundColor Yellow
    Write-Host "  - Individual sheets for each group showing members" -ForegroundColor Yellow
}

