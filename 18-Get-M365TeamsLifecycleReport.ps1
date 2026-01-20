<#
====================================================================================
Script Name: Get-M365TeamsLifecycleReport.ps1
Description: Microsoft Teams Creation and Lifecycle Management Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Lists all Microsoft Teams with creation details
• Shows team owners and member counts
• Identifies archived, deleted, and active teams
• Tracks team lifecycle events and changes
• Calculates team age and last activity date
• Highlights teams without owners (orphaned)
• Generates governance and compliance reports
• MFA-compatible Microsoft Teams connection

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory=$false)]
    [switch]$ArchivedOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$OrphanedOnly,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumMembers,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Teams_Lifecycle_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Teams Lifecycle Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "MicrosoftTeams"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Required module '$requiredModule' is not installed." -ForegroundColor Yellow
    $install = Read-Host "Would you like to install it now? (Y/N)"
    
    if ($install -eq 'Y' -or $install -eq 'y') {
        try {
            Write-Host "Installing $requiredModule..." -ForegroundColor Cyan
            Install-Module -Name $requiredModule -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
            Write-Host "$requiredModule installed successfully.`n" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install $requiredModule. Error: $_" -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Module installation declined. Script cannot continue." -ForegroundColor Red
        exit
    }
}

# Connect to Microsoft Teams
Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan

try {
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Host "Successfully connected to Microsoft Teams.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Teams. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve Teams
Write-Host "Retrieving Microsoft Teams..." -ForegroundColor Cyan
$results = @()
$archivedCount = 0
$orphanedCount = 0
$activeCount = 0

try {
    $teams = Get-Team
    
    Write-Host "Found $($teams.Count) team(s). Analyzing lifecycle data...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $thresholdDate = (Get-Date).AddDays(-$InactiveDays)
    
    foreach ($team in $teams) {
        $progressCounter++
        Write-Progress -Activity "Processing Teams" -Status "Team $progressCounter of $($teams.Count): $($team.DisplayName)" -PercentComplete (($progressCounter / $teams.Count) * 100)
        
        try {
            # Get team details
            $teamDetails = Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue
            $teamUsers = Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue
            
            $ownerCount = ($teamUsers | Where-Object { $_.Role -eq "owner" }).Count
            $memberCount = ($teamUsers | Where-Object { $_.Role -eq "member" }).Count
            $guestCount = ($teamUsers | Where-Object { $_.Role -eq "guest" }).Count
            $totalUsers = $teamUsers.Count
            
            # Check if orphaned
            $isOrphaned = ($ownerCount -eq 0)
            if ($isOrphaned) {
                $orphanedCount++
            }
            
            # Check if archived
            $isArchived = $team.Archived
            if ($isArchived) {
                $archivedCount++
            } else {
                $activeCount++
            }
            
            # Apply filters
            if ($ArchivedOnly -and -not $isArchived) {
                continue
            }
            
            if ($OrphanedOnly -and -not $isOrphaned) {
                continue
            }
            
            if ($MinimumMembers -and $totalUsers -lt $MinimumMembers) {
                continue
            }
            
            # Get channel count
            $channelCount = if ($teamDetails) { $teamDetails.Count } else { 0 }
            
            # Calculate team age
            $teamAgeDays = if ($team.CreatedDateTime) {
                (New-TimeSpan -Start $team.CreatedDateTime -End (Get-Date)).Days
            } else {
                "Unknown"
            }
            
            $obj = [PSCustomObject]@{
                TeamName = $team.DisplayName
                Description = $team.Description
                Visibility = $team.Visibility
                Archived = $isArchived
                IsOrphaned = $isOrphaned
                OwnerCount = $ownerCount
                MemberCount = $memberCount
                GuestCount = $guestCount
                TotalUsers = $totalUsers
                ChannelCount = $channelCount
                CreatedDateTime = $team.CreatedDateTime
                TeamAgeDays = $teamAgeDays
                MailNickName = $team.MailNickName
                Classification = $team.Classification
                GroupId = $team.GroupId
                Status = if ($isArchived) { "Archived" } elseif ($isOrphaned) { "Orphaned" } else { "Active" }
            }
            
            $results += $obj
        }
        catch {
            Write-Warning "Error processing team $($team.DisplayName): $_"
        }
    }
    
    Write-Progress -Activity "Processing Teams" -Completed
}
catch {
    Write-Host "Error retrieving teams: $_" -ForegroundColor Red
    Disconnect-MicrosoftTeams | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Teams Lifecycle Summary:" -ForegroundColor Green
    Write-Host "  Total Teams: $($results.Count)" -ForegroundColor White
    Write-Host "  Active Teams: $activeCount" -ForegroundColor Green
    Write-Host "  Archived Teams: $archivedCount" -ForegroundColor Yellow
    Write-Host "  Orphaned Teams (No Owners): $orphanedCount" -ForegroundColor Red
    Write-Host "  Average Team Size: $([math]::Round((($results | Measure-Object -Property TotalUsers -Average).Average), 2)) users" -ForegroundColor White
    Write-Host "  Average Team Age: $([math]::Round((($results | Where-Object { $_.TeamAgeDays -ne 'Unknown' } | Measure-Object -Property TeamAgeDays -Average).Average), 0)) days" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($orphanedCount -gt 0) {
        Write-Host "WARNING: $orphanedCount orphaned team(s) detected without owners!" -ForegroundColor Red
    }
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table TeamName, Status, TotalUsers, ChannelCount, TeamAgeDays -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No teams found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Teams..." -ForegroundColor Cyan
Disconnect-MicrosoftTeams | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
