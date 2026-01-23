<#
====================================================================================
Script Name: Get-M365GroupMembershipsReport.ps1
Description: Comprehensive Microsoft 365 Groups and membership analysis
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Microsoft 365 Groups and Teams
• Lists all members and owners for each group
• Identifies groups without owners
• Shows external/guest members in groups
• Calculates group size and activity metrics
• Supports filtering by group type and membership
• Exports detailed hierarchical CSV reports
• MFA-compatible Microsoft Graph authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$GroupDisplayName,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSecurityGroups,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowGuestMembersOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$OrphanedGroupsOnly,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumMemberCount,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Group_Memberships_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Group Memberships Report Generator" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph.Groups", "Microsoft.Graph.Users")

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Required module '$module' is not installed." -ForegroundColor Yellow
        $install = Read-Host "Would you like to install it now? (Y/N)"
        
        if ($install -eq 'Y' -or $install -eq 'y') {
            try {
                Write-Host "Installing $module..." -ForegroundColor Cyan
                Install-Module -Name $module -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
                Write-Host "$module installed successfully.`n" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to install $module. Error: $_" -ForegroundColor Red
                exit
            }
        }
        else {
            Write-Host "Module installation declined. Script cannot continue." -ForegroundColor Red
            exit
        }
    }
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "Group.Read.All", "GroupMember.Read.All", "User.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Build filter
$filter = "mailEnabled eq true and securityEnabled eq false"
if ($IncludeSecurityGroups) {
    $filter = $null
}

# Retrieve groups
Write-Host "Retrieving Microsoft 365 Groups..." -ForegroundColor Cyan
$results = @()
$orphanedCount = 0
$guestMemberCount = 0

try {
    if ($GroupDisplayName) {
        $groups = Get-MgGroup -Filter "displayName eq '$GroupDisplayName'" -All
    }
    elseif ($filter) {
        $groups = Get-MgGroup -Filter $filter -All
    }
    else {
        $groups = Get-MgGroup -All
    }
    
    Write-Host "Found $($groups.Count) group(s). Retrieving membership details...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($group in $groups) {
        $progressCounter++
        Write-Progress -Activity "Processing Groups" -Status "Group $progressCounter of $($groups.Count): $($group.DisplayName)" -PercentComplete (($progressCounter / $groups.Count) * 100)
        
        try {
            # Get members
            $members = Get-MgGroupMember -GroupId $group.Id -All
            $memberCount = $members.Count
            
            # Get owners
            $owners = Get-MgGroupOwner -GroupId $group.Id -All
            $ownerCount = $owners.Count
            
            # Check for orphaned groups
            $isOrphaned = ($ownerCount -eq 0)
            if ($isOrphaned) {
                $orphanedCount++
            }
            
            # Skip if filters don't match
            if ($OrphanedGroupsOnly -and -not $isOrphaned) {
                continue
            }
            
            if ($MinimumMemberCount -and $memberCount -lt $MinimumMemberCount) {
                continue
            }
            
            # Process members
            $guestMembers = @()
            $memberDetails = @()
            
            foreach ($member in $members) {
                try {
                    $memberUser = Get-MgUser -UserId $member.Id -Property DisplayName, UserPrincipalName, UserType -ErrorAction SilentlyContinue
                    
                    if ($memberUser) {
                        $memberDetails += "$($memberUser.DisplayName) ($($memberUser.UserPrincipalName))"
                        
                        if ($memberUser.UserType -eq "Guest") {
                            $guestMembers += $memberUser.UserPrincipalName
                            $guestMemberCount++
                        }
                    }
                }
                catch {
                    # Member might not be a user (could be another group or service principal)
                    $memberDetails += $member.Id
                }
            }
            
            # Skip if showing guest members only
            if ($ShowGuestMembersOnly -and $guestMembers.Count -eq 0) {
                continue
            }
            
            # Process owners
            $ownerDetails = @()
            foreach ($owner in $owners) {
                try {
                    $ownerUser = Get-MgUser -UserId $owner.Id -Property DisplayName, UserPrincipalName -ErrorAction SilentlyContinue
                    if ($ownerUser) {
                        $ownerDetails += "$($ownerUser.DisplayName) ($($ownerUser.UserPrincipalName))"
                    }
                }
                catch {
                    $ownerDetails += $owner.Id
                }
            }
            
            $obj = [PSCustomObject]@{
                GroupName = $group.DisplayName
                GroupEmail = $group.Mail
                GroupType = if ($group.MailEnabled -and -not $group.SecurityEnabled) { "Microsoft 365" } 
                           elseif ($group.SecurityEnabled -and -not $group.MailEnabled) { "Security" }
                           else { "Mail-Enabled Security" }
                Description = $group.Description
                MemberCount = $memberCount
                OwnerCount = $ownerCount
                GuestMemberCount = $guestMembers.Count
                IsOrphaned = $isOrphaned
                Owners = ($ownerDetails -join "; ")
                Members = ($memberDetails -join "; ")
                GuestMembers = ($guestMembers -join "; ")
                CreatedDateTime = $group.CreatedDateTime
                Visibility = $group.Visibility
                GroupId = $group.Id
            }
            
            $results += $obj
        }
        catch {
            Write-Warning "Error processing group $($group.DisplayName): $_"
        }
    }
    
    Write-Progress -Activity "Processing Groups" -Completed
}
catch {
    Write-Host "Error retrieving groups: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Group Membership Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total Groups Analyzed: $($results.Count)" -ForegroundColor White
    Write-Host "  Total Members Across All Groups: $(($results | Measure-Object -Property MemberCount -Sum).Sum)" -ForegroundColor White
    Write-Host "  Groups Without Owners (Orphaned): $orphanedCount" -ForegroundColor Yellow
    Write-Host "  Total Guest Members: $(($results | Measure-Object -Property GuestMemberCount -Sum).Sum)" -ForegroundColor White
    Write-Host "  Average Members Per Group: $([math]::Round((($results | Measure-Object -Property MemberCount -Average).Average), 2))" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($orphanedCount -gt 0) {
        Write-Host "WARNING: $orphanedCount orphaned group(s) detected without owners!" -ForegroundColor Red
    }
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table GroupName, GroupType, MemberCount, OwnerCount, GuestMemberCount, IsOrphaned -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No groups found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
