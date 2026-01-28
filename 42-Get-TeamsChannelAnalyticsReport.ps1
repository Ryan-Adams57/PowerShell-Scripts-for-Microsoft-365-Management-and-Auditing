# ====================================================================================
# M365 POWERSHELL REPORTING SCRIPT - PRODUCTION READY
# Enterprise-grade reporting with comprehensive error handling
# ====================================================================================
#
<#
====================================================================================
Script Name: 42-Get-TeamsChannelAnalyticsReport.ps1
Description: Microsoft Teams channel usage and engagement analytics
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

REQUIREMENTS:
• PowerShell 5.1 or higher
• Appropriate M365 administrator permissions
• Required modules (validated at runtime)

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$TeamName,
    
    [Parameter(Mandatory=$false)]
    [switch]$PrivateChannelsOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeMemberCounts,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumMembers,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\\Teams_Channel_Analytics_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest

# Comprehensive logging function
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info","Warning","Error","Success")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Info" { "Cyan" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] $Level: $Message" -ForegroundColor $color
    if ($Level -eq "Error") {
        $script:ErrorCount++
    }
}

# Parameter validation function
function Test-ScriptParameters {
    Write-Log "Validating script parameters..." -Level Info
    if ($ExportPath -and -not (Test-Path (Split-Path $ExportPath -Parent) -ErrorAction SilentlyContinue)) {
        Write-Log "Export path directory does not exist. Creating..." -Level Warning
        try {
            New-Item -Path (Split-Path $ExportPath -Parent) -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Log "Directory created successfully" -Level Success
        } catch {
            Write-Log "Failed to create directory: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
    Write-Log "Parameter validation complete" -Level Success
    return $true
}

$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:Results = @()
$script:ErrorCount = 0

Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "M365 Reporting Script - Production Ready" -ForegroundColor Green
Write-Host "Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
Write-Host "\n====================================================================================\n" -ForegroundColor Cyan


Set-StrictMode -Version Latest

# Comprehensive logging function
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info","Warning","Error","Success")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Info" { "Cyan" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] $Level: $Message" -ForegroundColor $color
    if ($Level -eq "Error") {
        $script:ErrorCount++
    }
}

# Parameter validation function
function Test-ScriptParameters {
    Write-Log "Validating script parameters..." -Level Info
    if ($ExportPath -and -not (Test-Path (Split-Path $ExportPath -Parent) -ErrorAction SilentlyContinue)) {
        Write-Log "Export path directory does not exist. Creating..." -Level Warning
        try {
            New-Item -Path (Split-Path $ExportPath -Parent) -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Log "Directory created successfully" -Level Success
        } catch {
            Write-Log "Failed to create directory: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
    Write-Log "Parameter validation complete" -Level Success
    return $true
}

$ErrorActionPreference = "Stop"

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft Teams Channel Analytics Report" -ForegroundColor Green
Write-Host "====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "MicrosoftTeams"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.`n" -ForegroundColor Green
    } else { exit }
}

# Connect
Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
try {
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve teams and channels
Write-Host "Retrieving Teams channels..." -ForegroundColor Cyan
$script:Results = @()
$totalChannels = 0
$privateChannels = 0

try {
    if ($TeamName) {
        $teams = Get-Team -DisplayName $TeamName
    } else {
        $teams = Get-Team
    }
    
    Write-Host "Found $($teams.Count) team(s). Analyzing channels...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($team in $teams) {
        $progressCounter++
        Write-Progress -Activity "Processing Teams" -Status "Team $progressCounter of $($teams.Count): $($team.DisplayName)" -PercentComplete (($progressCounter / $teams.Count) * 100)
        
        try {
            $channels = Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue
            
            if ($channels) {
                foreach ($channel in $channels) {
                    $totalChannels++
                    
                    $isPrivate = $channel.MembershipType -eq "Private"
                    if ($isPrivate) { $privateChannels++ }
                    
                    if ($PrivateChannelsOnly -and -not $isPrivate) { continue }
                    
                    # Get member count if requested
                    $memberCount = 0
                    if ($IncludeMemberCounts -and $isPrivate) {
                        try {
                            $members = Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName -ErrorAction SilentlyContinue
                            $memberCount = if ($members) { $members.Count } else { 0 }
                        } catch {
                            $memberCount = "N/A"
                        }
                    }
                    
                    if ($MinimumMembers -and $memberCount -lt $MinimumMembers) { continue }
                    
                    $script:Results += [PSCustomObject]@{
                        TeamName = $team.DisplayName
                        TeamId = $team.GroupId
                        ChannelName = $channel.DisplayName
                        ChannelType = $channel.MembershipType
                        Description = $channel.Description
                        MemberCount = $memberCount
                        IsPrivate = $isPrivate
                        ChannelId = $channel.Id
                    }
                }
            }
        } catch {
            Write-Warning "Error processing team $($team.DisplayName): $_"
        }
    }
    
    Write-Progress -Activity "Processing Teams" -Completed
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MicrosoftTeams | Out-Null
    exit
}

# Export
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Teams Channel Summary:" -ForegroundColor Green
    Write-Host "  Total Teams: $($teams.Count)" -ForegroundColor White
    Write-Host "  Total Channels: $totalChannels" -ForegroundColor White
    Write-Host "  Private Channels: $privateChannels" -ForegroundColor Yellow
    Write-Host "  Standard Channels: $($totalChannels - $privateChannels)" -ForegroundColor Green
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table TeamName, ChannelName, ChannelType, MemberCount -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No channels found." -ForegroundColor Yellow
}

Disconnect-MicrosoftTeams | Out-Null

# Comprehensive cleanup and summary
Write-Log "Performing cleanup operations..." -Level Info

$script:EndTime = Get-Date
$script:Duration = $script:EndTime - $script:StartTime

try {
    # Disconnect from services
    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    if (Get-Command Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Log "Disconnected from services" -Level Success
} catch {
    Write-Log "Disconnect completed with warnings" -Level Warning
}

Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "Execution Summary:" -ForegroundColor Green
Write-Host "  Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor White
Write-Host "  Duration: $($script:Duration.TotalSeconds) seconds" -ForegroundColor White
Write-Host "  Results Collected: $($script:Results.Count)" -ForegroundColor White
Write-Host "  Errors Encountered: $script:ErrorCount" -ForegroundColor White
if ($ExportPath -and (Test-Path $ExportPath)) {
    Write-Host "  Report Location: $ExportPath" -ForegroundColor Cyan
}
Write-Host "\n====================================================================================\n" -ForegroundColor Cyan

if ($script:ErrorCount -eq 0) {
    Write-Log "Script completed successfully!" -Level Success
    exit 0
} else {
    Write-Log "Script completed with $script:ErrorCount error(s)" -Level Warning
    exit 1
}
Write-Host "Completed.`n" -ForegroundColor Green
