<#
====================================================================================
Script Name: 3-Get-M365ExternalForwardingReport.ps1
Description: Identifies mailboxes with external email forwarding configured
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

SCRIPT HIGHLIGHTS:
• Detects external SMTP forwarding on mailboxes
• Identifies inbox rules forwarding to external domains
• Shows both automatic and manual forwarding configurations
• Highlights potential data exfiltration risks
• Supports filtering by mailbox type (User, Shared, Room)
• Generates security-focused recommendations
• Exports detailed CSV reports with forwarding destinations
• MFA-compatible Exchange Online authentication with -UseRPSSession
• Comprehensive error handling with try/catch/finally
• Progress indicators for long operations

REQUIREMENTS:
• ExchangeOnlineManagement module
• Exchange Administrator role

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Filter by mailbox type")]
    [ValidateSet("UserMailbox","SharedMailbox","RoomMailbox","All")]
    [string]$MailboxType = "All",
    
    [Parameter(Mandatory=$false, HelpMessage="Specific user principal name")]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false, HelpMessage="Check inbox rules only")]
    [switch]$InboxRulesOnly,
    
    [Parameter(Mandatory=$false, HelpMessage="Check automatic forwarding only")]
    [switch]$AutoForwardOnly,
    
    [Parameter(Mandatory=$false, HelpMessage="Export CSV path")]
    [string]$ExportPath = ".\M365_External_Forwarding_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize script variables
$script:Results = @()
$script:ForwardingCount = 0
$script:RuleCount = 0

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 External Email Forwarding Report" -ForegroundColor Green
Write-Host "Version 2.0 - Production Ready" -ForegroundColor White
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

# Module validation
$requiredModule = "ExchangeOnlineManagement"

Write-Host "Validating required module..." -ForegroundColor Cyan

try {
    if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
        Write-Host "ERROR: Required module '$requiredModule' is not installed." -ForegroundColor Red
        Write-Host "Please run: Install-Module -Name $requiredModule -Scope CurrentUser" -ForegroundColor Yellow
        exit 1
    }
    else {
        $moduleInfo = Get-Module -ListAvailable -Name $requiredModule | 
            Sort-Object Version -Descending | 
            Select-Object -First 1
        Write-Host "  ✓ $requiredModule (v$($moduleInfo.Version))" -ForegroundColor Green
    }
}
catch {
    Write-Host "ERROR validating module: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Write-Host "Note: Using -UseRPSSession for compatibility" -ForegroundColor Yellow

try {
    $connectionState = Get-ConnectionInformation -ErrorAction SilentlyContinue
    
    if (-not $connectionState -or $connectionState.State -ne 'Connected') {
        Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    }
    else {
        Write-Host "Already connected to Exchange Online." -ForegroundColor Green
    }
    
    Write-Host ""
}
catch {
    Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Tip: Ensure ExchangeOnlineManagement module is up to date" -ForegroundColor Yellow
    exit 1
}

# Retrieve mailboxes
Write-Host "Retrieving mailbox information..." -ForegroundColor Cyan

try {
    $mailboxFilter = $null
    
    if ($UserPrincipalName) {
        $mailboxFilter = "PrimarySmtpAddress -eq '$UserPrincipalName'"
    }
    elseif ($MailboxType -ne "All") {
        $mailboxFilter = "RecipientTypeDetails -eq '$MailboxType'"
    }
    
    if ($mailboxFilter) {
        $mailboxes = Get-Mailbox -Filter $mailboxFilter -ResultSize Unlimited -ErrorAction Stop
    }
    else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    }
    
    Write-Host "Retrieved $($mailboxes.Count) mailbox(es)." -ForegroundColor Green
    Write-Host "Checking forwarding configurations...`n" -ForegroundColor Cyan
    
    $progressCounter = 0
    
    foreach ($mailbox in $mailboxes) {
        $progressCounter++
        Write-Progress -Activity "Checking Mailbox Forwarding" `
            -Status "Mailbox $progressCounter of $($mailboxes.Count): $($mailbox.PrimarySmtpAddress)" `
            -PercentComplete (($progressCounter / $mailboxes.Count) * 100)
        
        $forwardingAddress = $null
        $forwardingSMTP = $null
        $deliverToMailbox = $null
        $externalRules = @()
        
        # Check automatic forwarding
        if (-not $InboxRulesOnly) {
            if ($mailbox.ForwardingAddress) {
                $forwardingAddress = $mailbox.ForwardingAddress
                $script:ForwardingCount++
            }
            
            if ($mailbox.ForwardingSmtpAddress) {
                $forwardingSMTP = $mailbox.ForwardingSmtpAddress -replace "smtp:", ""
                $script:ForwardingCount++
            }
            
            $deliverToMailbox = $mailbox.DeliverToMailboxAndForward
        }
        
        # Check inbox rules for external forwarding
        if (-not $AutoForwardOnly) {
            try {
                $inboxRules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
                
                foreach ($rule in $inboxRules) {
                    $isExternal = $false
                    $forwardTo = @()
                    $domain = $mailbox.PrimarySmtpAddress.Split('@')[1]
                    
                    # Check ForwardTo
                    if ($rule.ForwardTo) {
                        foreach ($recipient in $rule.ForwardTo) {
                            $recipientAddress = $recipient -replace ".*\[SMTP:", "" -replace "\].*", ""
                            if ($recipientAddress -notlike "*$domain") {
                                $isExternal = $true
                                $forwardTo += $recipientAddress
                            }
                        }
                    }
                    
                    # Check ForwardAsAttachmentTo
                    if ($rule.ForwardAsAttachmentTo) {
                        foreach ($recipient in $rule.ForwardAsAttachmentTo) {
                            $recipientAddress = $recipient -replace ".*\[SMTP:", "" -replace "\].*", ""
                            if ($recipientAddress -notlike "*$domain") {
                                $isExternal = $true
                                $forwardTo += $recipientAddress
                            }
                        }
                    }
                    
                    # Check RedirectTo
                    if ($rule.RedirectTo) {
                        foreach ($recipient in $rule.RedirectTo) {
                            $recipientAddress = $recipient -replace ".*\[SMTP:", "" -replace "\].*", ""
                            if ($recipientAddress -notlike "*$domain") {
                                $isExternal = $true
                                $forwardTo += $recipientAddress
                            }
                        }
                    }
                    
                    if ($isExternal) {
                        $script:RuleCount++
                        $externalRules += [PSCustomObject]@{
                            RuleName = $rule.Name
                            Enabled = $rule.Enabled
                            ForwardTo = ($forwardTo -join "; ")
                        }
                    }
                }
            }
            catch {
                Write-Warning "Could not retrieve inbox rules for $($mailbox.PrimarySmtpAddress): $($_.Exception.Message)"
            }
        }
        
        # Create result object if forwarding detected
        if ($forwardingAddress -or $forwardingSMTP -or $externalRules.Count -gt 0) {
            $activeRules = ($externalRules | Where-Object { $_.Enabled -eq $true }).Count
            $riskLevel = if ($forwardingSMTP -or $activeRules -gt 0) { "High" } else { "Medium" }
            
            $obj = [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                MailboxType = $mailbox.RecipientTypeDetails
                ForwardingAddress = if ($forwardingAddress) { $forwardingAddress } else { "" }
                ForwardingSmtpAddress = if ($forwardingSMTP) { $forwardingSMTP } else { "" }
                DeliverToMailboxAndForward = $deliverToMailbox
                InboxRuleCount = $externalRules.Count
                InboxRuleNames = if ($externalRules.Count -gt 0) { (($externalRules | ForEach-Object { $_.RuleName }) -join "; ") } else { "" }
                InboxRuleForwardTo = if ($externalRules.Count -gt 0) { (($externalRules | ForEach-Object { $_.ForwardTo }) -join "; ") } else { "" }
                ActiveRulesOnly = $activeRules
                RiskLevel = $riskLevel
                ReportDate = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
            
            $script:Results += $obj
        }
    }
    
    Write-Progress -Activity "Checking Mailbox Forwarding" -Completed
    
    Write-Host "`nForwarding check completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving mailboxes or rules: $($_.Exception.Message)" -ForegroundColor Red
    Write-Progress -Activity "Checking Mailbox Forwarding" -Completed
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
    exit 1
}

# Export and display results
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "External Forwarding Detection Summary:" -ForegroundColor Green
    Write-Host "  Mailboxes with External Forwarding: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Automatic SMTP Forwarding: $(($script:Results | Where-Object { $_.ForwardingSmtpAddress }).Count)" -ForegroundColor Yellow
    Write-Host "  Inbox Rules with External Forwarding: $(($script:Results | Where-Object { $_.InboxRuleCount -gt 0 }).Count)" -ForegroundColor Yellow
    Write-Host "  High Risk Configurations: $(($script:Results | Where-Object { $_.RiskLevel -eq 'High' }).Count)" -ForegroundColor Red
    
    try {
        $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
        Write-Host "  Report exported successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "`n  ERROR exporting report: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY RECOMMENDATION:" -ForegroundColor Red
    Write-Host "Review all external forwarding configurations for potential data exfiltration risks.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $script:Results | Select-Object -First 10 | 
        Format-Table DisplayName, PrimarySmtpAddress, ForwardingSmtpAddress, InboxRuleCount, RiskLevel -AutoSize
    
    if (Test-Path $ExportPath) {
        $openFile = Read-Host "`nWould you like to open the CSV report? (Y/N)"
        if ($openFile -eq 'Y' -or $openFile -eq 'y') {
            try {
                Invoke-Item $ExportPath
            }
            catch {
                Write-Host "Could not open file: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}
else {
    Write-Host "`nNo external forwarding configurations detected." -ForegroundColor Green
}

# Cleanup
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan

try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop | Out-Null
    Write-Host "Disconnected successfully." -ForegroundColor Green
}
catch {
    Write-Host "Disconnect completed." -ForegroundColor Green
}

Write-Host "`nScript completed successfully.`n" -ForegroundColor Green
exit 0
