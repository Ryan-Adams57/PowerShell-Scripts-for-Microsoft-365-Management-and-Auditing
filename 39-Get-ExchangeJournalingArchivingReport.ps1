<#
====================================================================================
Script Name: 39-Get-ExchangeJournalingArchivingReport.ps1
Description: Exchange Online journaling rules and in-place archiving configuration
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
    [switch]$IncludeArchiveMailboxes,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowJournalingRules,
    
    [Parameter(Mandatory=$false)]
    [switch]$ArchiveEnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("All","Internal","External","Global")]
    [string]$JournalScope = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Exchange_Journaling_Archiving_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Exchange Online Journaling and Archiving Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"

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

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

try {
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online. Error: $_" -ForegroundColor Red
    exit
}

# Initialize results
$script:Results = @()
$journalingRuleCount = 0
$archiveEnabledCount = 0
$autoExpandingCount = 0

# Retrieve journaling rules
if ($ShowJournalingRules) {
    Write-Host "Retrieving journaling rules..." -ForegroundColor Cyan
    
    try {
        $journalRules = Get-JournalRule -ErrorAction Stop
        $journalingRuleCount = $journalRules.Count
        
        Write-Host "Found $journalingRuleCount journaling rule(s).`n" -ForegroundColor Green
        
        foreach ($rule in $journalRules) {
            # Filter by scope
            if ($JournalScope -ne "All" -and $rule.Scope -ne $JournalScope) {
                continue
            }
            
            $obj = [PSCustomObject]@{
                Type = "JournalingRule"
                Name = $rule.Name
                Enabled = $rule.Enabled
                Scope = $rule.Scope
                Recipient = $rule.Recipient
                JournalEmailAddress = $rule.JournalEmailAddress
                FullReport = $rule.FullReport
                LawfulInterception = $rule.LawfulInterception
                ExpiryDate = $rule.ExpiryDate
                RuleIdentity = $rule.Identity
                ArchiveEnabled = "N/A"
                ArchiveStatus = "N/A"
                ArchiveQuota = "N/A"
                AutoExpandingArchive = "N/A"
            }
            
            $script:Results += $obj
        }
    }
    catch {
        Write-Host "Error retrieving journaling rules: $_" -ForegroundColor Red
    }
}

# Retrieve archive mailbox information
if ($IncludeArchiveMailboxes) {
    Write-Host "Retrieving mailboxes with archive information..." -ForegroundColor Cyan
    
    try {
        # Get mailboxes
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        
        Write-Host "Found $($mailboxes.Count) mailbox(es). Checking archive status...`n" -ForegroundColor Green
        
        $progressCounter = 0
        
        foreach ($mailbox in $mailboxes) {
            $progressCounter++
            Write-Progress -Activity "Processing Mailboxes" -Status "Mailbox $progressCounter of $($mailboxes.Count): $($mailbox.DisplayName)" -PercentComplete (($progressCounter / $mailboxes.Count) * 100)
            
            $archiveEnabled = $mailbox.ArchiveStatus -ne "None"
            
            # Filter for archive-enabled only if specified
            if ($ArchiveEnabledOnly -and -not $archiveEnabled) {
                continue
            }
            
            if ($archiveEnabled) {
                $archiveEnabledCount++
            }
            
            # Get archive statistics if archive is enabled
            $archiveSize = "N/A"
            $archiveQuota = "N/A"
            $archiveWarningQuota = "N/A"
            
            if ($archiveEnabled) {
                try {
                    $archiveStats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName -Archive -ErrorAction SilentlyContinue
                    
                    if ($archiveStats) {
                        $archiveSize = $archiveStats.TotalItemSize.ToString()
                        $archiveQuota = $mailbox.ArchiveQuota
                        $archiveWarningQuota = $mailbox.ArchiveWarningQuota
                    }
                }
                catch {
                    Write-Warning "Could not retrieve archive statistics for $($mailbox.DisplayName)"
                }
            }
            
            # Check auto-expanding archive
            $autoExpanding = $mailbox.AutoExpandingArchiveEnabled
            if ($autoExpanding) {
                $autoExpandingCount++
            }
            
            $obj = [PSCustomObject]@{
                Type = "ArchiveMailbox"
                Name = $mailbox.DisplayName
                UserPrincipalName = $mailbox.UserPrincipalName
                Enabled = "N/A"
                Scope = "N/A"
                Recipient = "N/A"
                JournalEmailAddress = "N/A"
                FullReport = "N/A"
                LawfulInterception = "N/A"
                ExpiryDate = "N/A"
                RuleIdentity = "N/A"
                ArchiveEnabled = $archiveEnabled
                ArchiveStatus = $mailbox.ArchiveStatus
                ArchiveSize = $archiveSize
                ArchiveQuota = $archiveQuota
                ArchiveWarningQuota = $archiveWarningQuota
                AutoExpandingArchive = $autoExpanding
                MailboxType = $mailbox.RecipientTypeDetails
            }
            
            $script:Results += $obj
        }
        
        Write-Progress -Activity "Processing Mailboxes" -Completed
    }
    catch {
        Write-Host "Error retrieving mailbox archive information: $_" -ForegroundColor Red
    }
}

# If neither option selected, get basic journaling info
if (-not $ShowJournalingRules -and -not $IncludeArchiveMailboxes) {
    Write-Host "Retrieving basic journaling configuration..." -ForegroundColor Cyan
    
    try {
        $journalRules = Get-JournalRule -ErrorAction SilentlyContinue
        
        if ($journalRules) {
            $journalingRuleCount = $journalRules.Count
            
            foreach ($rule in $journalRules) {
                $obj = [PSCustomObject]@{
                    Type = "JournalingRule"
                    Name = $rule.Name
                    Enabled = $rule.Enabled
                    Scope = $rule.Scope
                    JournalEmailAddress = $rule.JournalEmailAddress
                }
                
                $script:Results += $obj
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve journaling rules"
    }
}

# Export and display results
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Journaling and Archiving Summary:" -ForegroundColor Green
    
    if ($ShowJournalingRules) {
        Write-Host "  Total Journaling Rules: $journalingRuleCount" -ForegroundColor White
        Write-Host "  Enabled Rules: $(($script:Results | Where-Object { $_.Type -eq 'JournalingRule' -and $_.Enabled -eq $true }).Count)" -ForegroundColor Green
        Write-Host "  Disabled Rules: $(($script:Results | Where-Object { $_.Type -eq 'JournalingRule' -and $_.Enabled -eq $false }).Count)" -ForegroundColor Yellow
    }
    
    if ($IncludeArchiveMailboxes) {
        Write-Host "  Total Mailboxes Analyzed: $(($script:Results | Where-Object { $_.Type -eq 'ArchiveMailbox' }).Count)" -ForegroundColor White
        Write-Host "  Archive Enabled Mailboxes: $archiveEnabledCount" -ForegroundColor Green
        Write-Host "  Auto-Expanding Archives: $autoExpandingCount" -ForegroundColor Cyan
    }
    
    # Export to CSV
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "COMPLIANCE NOTE:" -ForegroundColor Cyan
    Write-Host "Journaling and archiving are critical for regulatory compliance.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $script:Results | Select-Object -First 10 | Format-Table Type, Name, Enabled, ArchiveEnabled, AutoExpandingArchive -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No journaling rules or archive mailboxes found." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
