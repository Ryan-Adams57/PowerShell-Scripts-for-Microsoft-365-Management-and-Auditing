<#
====================================================================================
Script Name: 19-Get-M365TeamsExternalAccessReport.ps1
Description: Production-ready M365 reporting script
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

REQUIREMENTS:
• PowerShell 5.1 or higher
• Appropriate M365 administrator permissions
• Required modules (will be validated at runtime)

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]

param(
    [Parameter(Mandatory=$false)]
    [switch]$TeamsWithGuestsOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowPolicyDetails,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Teams_External_Access_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Teams External Access Report" -ForegroundColor Green
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

Write-Host "Retrieving Teams external access configuration..." -ForegroundColor Cyan
$results = @()

try {
    # Get external access policy
    $externalAccessConfig = Get-CsTenantFederationConfiguration -ErrorAction SilentlyContinue
    $guestConfig = Get-CsTeamsGuestMessagingConfiguration -ErrorAction SilentlyContinue
    $guestCallingConfig = Get-CsTeamsGuestCallingConfiguration -ErrorAction SilentlyContinue
    $guestMeetingConfig = Get-CsTeamsGuestMeetingConfiguration -ErrorAction SilentlyContinue
    
    Write-Host "Retrieved tenant-level policies.`n" -ForegroundColor Green
    
    # Get all teams
    $teams = Get-Team
    
    Write-Host "Found $($teams.Count) team(s). Checking for guest members...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $teamsWithGuests = 0
    
    foreach ($team in $teams) {
        $progressCounter++
        Write-Progress -Activity "Analyzing Teams" -Status "Team $progressCounter of $($teams.Count): $($team.DisplayName)" -PercentComplete (($progressCounter / $teams.Count) * 100)
        
        try {
            $teamUsers = Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue
            $guestUsers = $teamUsers | Where-Object { $_.Role -eq "guest" }
            $guestCount = if ($guestUsers) { $guestUsers.Count } else { 0 }
            
            if ($guestCount -gt 0) {
                $teamsWithGuests++
            }
            
            # Filter if only showing teams with guests
            if ($TeamsWithGuestsOnly -and $guestCount -eq 0) {
                continue
            }
            
            $guestEmails = if ($guestUsers) {
                ($guestUsers | ForEach-Object { $_.User }) -join "; "
            } else {
                "None"
            }
            
            $obj = [PSCustomObject]@{
                TeamName = $team.DisplayName
                TeamVisibility = $team.Visibility
                GuestCount = $guestCount
                HasGuests = ($guestCount -gt 0)
                GuestUsers = $guestEmails
                TotalMembers = $teamUsers.Count
                ExternalAccessEnabled = $externalAccessConfig.AllowFederatedUsers
                AllowedDomains = if ($externalAccessConfig.AllowedDomains) { ($externalAccessConfig.AllowedDomains.AllowedDomain.Domain -join "; ") } else { "All" }
                BlockedDomains = if ($externalAccessConfig.BlockedDomains) { ($externalAccessConfig.BlockedDomains.BlockedDomain.Domain -join "; ") } else { "None" }
                GuestMessagingEnabled = $guestConfig.AllowUserChat
                GuestCallingEnabled = $guestCallingConfig.AllowPrivateCalling
                GuestMeetingEnabled = $guestMeetingConfig.AllowMeetNow
                GroupId = $team.GroupId
            }
            
            $results += $obj
        }
        catch {
            Write-Warning "Error processing team $($team.DisplayName): $_"
        }
    }
    
    Write-Progress -Activity "Analyzing Teams" -Completed
}
catch {
    Write-Host "Error retrieving Teams data: $_" -ForegroundColor Red
    Disconnect-MicrosoftTeams | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Teams External Access Summary:" -ForegroundColor Green
    Write-Host "  Total Teams Analyzed: $($results.Count)" -ForegroundColor White
    Write-Host "  Teams with Guest Members: $teamsWithGuests" -ForegroundColor Yellow
    Write-Host "  Total Guest Users: $(($results | Measure-Object -Property GuestCount -Sum).Sum)" -ForegroundColor White
    
    if ($ShowPolicyDetails) {
        Write-Host "`n  Tenant-Level External Access Settings:" -ForegroundColor Cyan
        Write-Host "    External Access Enabled: $($externalAccessConfig.AllowFederatedUsers)" -ForegroundColor White
        Write-Host "    Guest Messaging Enabled: $($guestConfig.AllowUserChat)" -ForegroundColor White
        Write-Host "    Guest Calling Enabled: $($guestCallingConfig.AllowPrivateCalling)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table TeamName, GuestCount, HasGuests, ExternalAccessEnabled -AutoSize
    
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
