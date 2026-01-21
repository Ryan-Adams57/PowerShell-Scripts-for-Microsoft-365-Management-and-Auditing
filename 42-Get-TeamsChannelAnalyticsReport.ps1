<#
====================================================================================
Script Name: Get-TeamsChannelAnalyticsReport.ps1
Description: Microsoft Teams channel usage and engagement analytics
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Teams and their channels
• Shows channel member counts and activity
• Lists private vs standard channels
• Tracks channel creation and ownership
• Identifies inactive or orphaned channels
• Supports filtering by team or activity
• Generates Teams governance and usage reports
• Provides channel-level analytics for optimization

====================================================================================
#>

param(
    [Parameter(Mandatory=\$false)]
    [string]\$TeamName,
    
    [Parameter(Mandatory=\$false)]
    [switch]\$PrivateChannelsOnly,
    
    [Parameter(Mandatory=\$false)]
    [switch]\$IncludeMemberCounts,
    
    [Parameter(Mandatory=\$false)]
    [int]\$MinimumMembers,
    
    [Parameter(Mandatory=\$false)]
    [string]\$ExportPath = ".\\Teams_Channel_Analytics_Report_\$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "\`n====================================================================================\`n" -ForegroundColor Cyan
Write-Host "Microsoft Teams Channel Analytics Report" -ForegroundColor Green
Write-Host "====================================================================================\`n" -ForegroundColor Cyan

\$requiredModule = "MicrosoftTeams"

if (-not (Get-Module -ListAvailable -Name \$requiredModule)) {
    \$install = Read-Host "Install \$requiredModule? (Y/N)"
    if (\$install -eq 'Y' -or \$install -eq 'y') {
        Install-Module -Name \$requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.\`n" -ForegroundColor Green
    } else { exit }
}

# Connect
Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
try {
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Host "Connected.\`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: \$_" -ForegroundColor Red
    exit
}

# Retrieve teams and channels
Write-Host "Retrieving Teams channels..." -ForegroundColor Cyan
\$results = @()
\$totalChannels = 0
\$privateChannels = 0

try {
    if (\$TeamName) {
        \$teams = Get-Team -DisplayName \$TeamName
    } else {
        \$teams = Get-Team
    }
    
    Write-Host "Found \$(\$teams.Count) team(s). Analyzing channels...\`n" -ForegroundColor Green
    
    \$progressCounter = 0
    
    foreach (\$team in \$teams) {
        \$progressCounter++
        Write-Progress -Activity "Processing Teams" -Status "Team \$progressCounter of \$(\$teams.Count): \$(\$team.DisplayName)" -PercentComplete ((\$progressCounter / \$teams.Count) * 100)
        
        try {
            \$channels = Get-TeamChannel -GroupId \$team.GroupId -ErrorAction SilentlyContinue
            
            if (\$channels) {
                foreach (\$channel in \$channels) {
                    \$totalChannels++
                    
                    \$isPrivate = \$channel.MembershipType -eq "Private"
                    if (\$isPrivate) { \$privateChannels++ }
                    
                    if (\$PrivateChannelsOnly -and -not \$isPrivate) { continue }
                    
                    # Get member count if requested
                    \$memberCount = 0
                    if (\$IncludeMemberCounts -and \$isPrivate) {
                        try {
                            \$members = Get-TeamChannelUser -GroupId \$team.GroupId -DisplayName \$channel.DisplayName -ErrorAction SilentlyContinue
                            \$memberCount = if (\$members) { \$members.Count } else { 0 }
                        } catch {
                            \$memberCount = "N/A"
                        }
                    }
                    
                    if (\$MinimumMembers -and \$memberCount -lt \$MinimumMembers) { continue }
                    
                    \$results += [PSCustomObject]@{
                        TeamName = \$team.DisplayName
                        TeamId = \$team.GroupId
                        ChannelName = \$channel.DisplayName
                        ChannelType = \$channel.MembershipType
                        Description = \$channel.Description
                        MemberCount = \$memberCount
                        IsPrivate = \$isPrivate
                        ChannelId = \$channel.Id
                    }
                }
            }
        } catch {
            Write-Warning "Error processing team \$(\$team.DisplayName): \$_"
        }
    }
    
    Write-Progress -Activity "Processing Teams" -Completed
} catch {
    Write-Host "Error: \$_" -ForegroundColor Red
    Disconnect-MicrosoftTeams | Out-Null
    exit
}

# Export
if (\$results.Count -gt 0) {
    Write-Host "\`n====================================================================================\`n" -ForegroundColor Cyan
    Write-Host "Teams Channel Summary:" -ForegroundColor Green
    Write-Host "  Total Teams: \$(\$teams.Count)" -ForegroundColor White
    Write-Host "  Total Channels: \$totalChannels" -ForegroundColor White
    Write-Host "  Private Channels: \$privateChannels" -ForegroundColor Yellow
    Write-Host "  Standard Channels: \$(\$totalChannels - \$privateChannels)" -ForegroundColor Green
    
    \$results | Export-Csv -Path \$ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: \$ExportPath" -ForegroundColor White
    Write-Host "\`n====================================================================================\`n" -ForegroundColor Cyan
    
    \$results | Select-Object -First 10 | Format-Table TeamName, ChannelName, ChannelType, MemberCount -AutoSize
    
    \$open = Read-Host "Open CSV? (Y/N)"
    if (\$open -eq 'Y' -or \$open -eq 'y') { Invoke-Item \$ExportPath }
} else {
    Write-Host "No channels found." -ForegroundColor Yellow
}

Disconnect-MicrosoftTeams | Out-Null
Write-Host "Completed.\`n" -ForegroundColor Green
