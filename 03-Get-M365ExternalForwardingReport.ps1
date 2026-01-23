<#
====================================================================================
Script Name: Get-M365ExternalForwardingReport.ps1
Description: Identifies mailboxes with external email forwarding configured
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Detects external SMTP forwarding on mailboxes
• Identifies inbox rules forwarding to external domains
• Shows both automatic and manual forwarding configurations
• Highlights potential data exfiltration risks
• Supports filtering by mailbox type (User, Shared, Room)
• Generates security-focused recommendations
• Exports detailed CSV reports with forwarding destinations
• MFA-compatible Exchange Online authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("UserMailbox","SharedMailbox","RoomMailbox","All")]
    [string]$MailboxType = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [switch]$InboxRulesOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$AutoForwardOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_External_Forwarding_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 External Email Forwarding Report" -ForegroundColor Green
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
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve mailboxes
Write-Host "Retrieving mailbox information..." -ForegroundColor Cyan
$results = @()
$forwardingCount = 0
$ruleCount = 0

try {
    $mailboxFilter = if ($UserPrincipalName) {
        "PrimarySmtpAddress -eq '$UserPrincipalName'"
    }
    elseif ($MailboxType -ne "All") {
        "RecipientTypeDetails -eq '$MailboxType'"
    }
    else {
        $null
    }
    
    if ($mailboxFilter) {
        $mailboxes = Get-Mailbox -Filter $mailboxFilter -ResultSize Unlimited
    }
    else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited
    }
    
    Write-Host "Found $($mailboxes.Count) mailbox(es). Checking forwarding configurations...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($mailbox in $mailboxes) {
        $progressCounter++
        Write-Progress -Activity "Checking Mailbox Forwarding" -Status "Mailbox $progressCounter of $($mailboxes.Count): $($mailbox.PrimarySmtpAddress)" -PercentComplete (($progressCounter / $mailboxes.Count) * 100)
        
        $forwardingAddress = $null
        $forwardingSMTP = $null
        $deliverToMailbox = $null
        $externalRules = @()
        
        # Check automatic forwarding
        if (-not $InboxRulesOnly) {
            if ($mailbox.ForwardingAddress) {
                $forwardingAddress = $mailbox.ForwardingAddress
                $forwardingCount++
            }
            
            if ($mailbox.ForwardingSmtpAddress) {
                $forwardingSMTP = $mailbox.ForwardingSmtpAddress -replace "smtp:", ""
                $forwardingCount++
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
                    
                    if ($rule.ForwardTo) {
                        foreach ($recipient in $rule.ForwardTo) {
                            $recipientAddress = $recipient -replace ".*\[SMTP:", "" -replace "\].*", ""
                            if ($recipientAddress -notlike "*$($mailbox.PrimarySmtpAddress.Split('@')[1])") {
                                $isExternal = $true
                                $forwardTo += $recipientAddress
                            }
                        }
                    }
                    
                    if ($rule.ForwardAsAttachmentTo) {
                        foreach ($recipient in $rule.ForwardAsAttachmentTo) {
                            $recipientAddress = $recipient -replace ".*\[SMTP:", "" -replace "\].*", ""
                            if ($recipientAddress -notlike "*$($mailbox.PrimarySmtpAddress.Split('@')[1])") {
                                $isExternal = $true
                                $forwardTo += $recipientAddress
                            }
                        }
                    }
                    
                    if ($rule.RedirectTo) {
                        foreach ($recipient in $rule.RedirectTo) {
                            $recipientAddress = $recipient -replace ".*\[SMTP:", "" -replace "\].*", ""
                            if ($recipientAddress -notlike "*$($mailbox.PrimarySmtpAddress.Split('@')[1])") {
                                $isExternal = $true
                                $forwardTo += $recipientAddress
                            }
                        }
                    }
                    
                    if ($isExternal) {
                        $ruleCount++
                        $externalRules += [PSCustomObject]@{
                            RuleName = $rule.Name
                            Enabled = $rule.Enabled
                            ForwardTo = ($forwardTo -join "; ")
                        }
                    }
                }
            }
            catch {
                Write-Warning "Could not retrieve inbox rules for $($mailbox.PrimarySmtpAddress): $_"
            }
        }
        
        # Create result object if forwarding detected
        if ($forwardingAddress -or $forwardingSMTP -or $externalRules.Count -gt 0) {
            $obj = [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                MailboxType = $mailbox.RecipientTypeDetails
                ForwardingAddress = $forwardingAddress
                ForwardingSmtpAddress = $forwardingSMTP
                DeliverToMailboxAndForward = $deliverToMailbox
                InboxRuleCount = $externalRules.Count
                InboxRuleNames = (($externalRules | ForEach-Object { $_.RuleName }) -join "; ")
                InboxRuleForwardTo = (($externalRules | ForEach-Object { $_.ForwardTo }) -join "; ")
                ActiveRulesOnly = (($externalRules | Where-Object { $_.Enabled -eq $true }) | Measure-Object).Count
                RiskLevel = if ($forwardingSMTP -or ($externalRules | Where-Object { $_.Enabled -eq $true })) { "High" } else { "Medium" }
            }
            
            $results += $obj
        }
    }
    
    Write-Progress -Activity "Checking Mailbox Forwarding" -Completed
}
catch {
    Write-Host "Error retrieving mailboxes or rules: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "External Forwarding Detection Summary:" -ForegroundColor Green
    Write-Host "  Mailboxes with External Forwarding: $($results.Count)" -ForegroundColor White
    Write-Host "  Automatic SMTP Forwarding Detected: $(($results | Where-Object { $_.ForwardingSmtpAddress }).Count)" -ForegroundColor Yellow
    Write-Host "  Inbox Rules with External Forwarding: $(($results | Where-Object { $_.InboxRuleCount -gt 0 }).Count)" -ForegroundColor Yellow
    Write-Host "  High Risk Configurations: $(($results | Where-Object { $_.RiskLevel -eq 'High' }).Count)" -ForegroundColor Red
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY RECOMMENDATION:" -ForegroundColor Red
    Write-Host "Review all external forwarding configurations for potential data exfiltration risks.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table DisplayName, PrimarySmtpAddress, ForwardingSmtpAddress, InboxRuleCount, RiskLevel -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No external forwarding configurations detected." -ForegroundColor Green
}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
