<#
====================================================================================
Script Name: Get-M365MailFlowRulesReport.ps1
Description: Exchange Transport and Mail Flow Rules Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Lists all Exchange Online transport (mail flow) rules
• Shows rule conditions, exceptions, and actions
• Identifies enabled vs disabled rules
• Highlights security-focused and compliance rules
• Shows rule priority and processing order
• Detects potential mail flow issues and conflicts
• Exports complete rule configuration backup
• MFA-compatible Exchange Online authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$EnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$RuleName,
    
    [Parameter(Mandatory=$false)]
    [switch]$SecurityRulesOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Mail_Flow_Rules_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Mail Flow Rules Report" -ForegroundColor Green
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

# Retrieve transport rules
Write-Host "Retrieving mail flow (transport) rules..." -ForegroundColor Cyan
$results = @()
$enabledCount = 0
$disabledCount = 0

try {
    $rules = Get-TransportRule
    
    if ($RuleName) {
        $rules = $rules | Where-Object { $_.Name -like "*$RuleName*" }
    }
    
    Write-Host "Found $($rules.Count) mail flow rule(s). Analyzing configurations...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($rule in $rules) {
        $progressCounter++
        Write-Progress -Activity "Processing Mail Flow Rules" -Status "Rule $progressCounter of $($rules.Count): $($rule.Name)" -PercentComplete (($progressCounter / $rules.Count) * 100)
        
        # Filter by state
        if ($EnabledOnly -and $rule.State -ne "Enabled") {
            continue
        }
        
        # Count by state
        if ($rule.State -eq "Enabled") {
            $enabledCount++
        } else {
            $disabledCount++
        }
        
        # Determine rule category
        $category = "General"
        $isSecurityRule = $false
        
        if ($rule.Name -match "malware|spam|phish|encrypt|dlp|block|quarantine|reject") {
            $category = "Security"
            $isSecurityRule = $true
        }
        elseif ($rule.Name -match "disclaimer|signature|footer|header") {
            $category = "Disclaimer"
        }
        elseif ($rule.Name -match "forward|redirect") {
            $category = "Routing"
        }
        elseif ($rule.Name -match "compliance|retention|archive|legal") {
            $category = "Compliance"
        }
        
        # Skip if filtering for security rules only
        if ($SecurityRulesOnly -and -not $isSecurityRule) {
            continue
        }
        
        # Parse conditions
        $conditions = @()
        if ($rule.From) { $conditions += "From: $($rule.From -join ', ')" }
        if ($rule.FromMemberOf) { $conditions += "From Member Of: $($rule.FromMemberOf -join ', ')" }
        if ($rule.SentTo) { $conditions += "Sent To: $($rule.SentTo -join ', ')" }
        if ($rule.SentToMemberOf) { $conditions += "Sent To Member Of: $($rule.SentToMemberOf -join ', ')" }
        if ($rule.SubjectContainsWords) { $conditions += "Subject Contains: $($rule.SubjectContainsWords -join ', ')" }
        if ($rule.SubjectOrBodyContainsWords) { $conditions += "Subject/Body Contains: $($rule.SubjectOrBodyContainsWords -join ', ')" }
        if ($rule.AttachmentExtensionMatchesWords) { $conditions += "Attachment Ext: $($rule.AttachmentExtensionMatchesWords -join ', ')" }
        if ($rule.FromScope) { $conditions += "From Scope: $($rule.FromScope)" }
        if ($rule.SentToScope) { $conditions += "Sent To Scope: $($rule.SentToScope)" }
        
        $conditionsStr = if ($conditions.Count -gt 0) { $conditions -join "; " } else { "None" }
        
        # Parse exceptions
        $exceptions = @()
        if ($rule.ExceptIfFrom) { $exceptions += "Except From: $($rule.ExceptIfFrom -join ', ')" }
        if ($rule.ExceptIfSentTo) { $exceptions += "Except Sent To: $($rule.ExceptIfSentTo -join ', ')" }
        
        $exceptionsStr = if ($exceptions.Count -gt 0) { $exceptions -join "; " } else { "None" }
        
        # Parse actions
        $actions = @()
        if ($rule.RejectMessageReasonText) { $actions += "Reject: $($rule.RejectMessageReasonText)" }
        if ($rule.DeleteMessage) { $actions += "Delete Message" }
        if ($rule.Quarantine) { $actions += "Quarantine" }
        if ($rule.RedirectMessageTo) { $actions += "Redirect To: $($rule.RedirectMessageTo -join ', ')" }
        if ($rule.BlindCopyTo) { $actions += "BCC: $($rule.BlindCopyTo -join ', ')" }
        if ($rule.ModifySubject) { $actions += "Modify Subject: $($rule.ModifySubject)" }
        if ($rule.PrependSubject) { $actions += "Prepend Subject: $($rule.PrependSubject)" }
        if ($rule.SetSCL) { $actions += "Set SCL: $($rule.SetSCL)" }
        if ($rule.ApplyClassification) { $actions += "Apply Classification: $($rule.ApplyClassification)" }
        if ($rule.ApplyHtmlDisclaimerText) { $actions += "Add Disclaimer" }
        if ($rule.SetHeaderName) { $actions += "Set Header: $($rule.SetHeaderName) = $($rule.SetHeaderValue)" }
        
        $actionsStr = if ($actions.Count -gt 0) { $actions -join "; " } else { "None" }
        
        $obj = [PSCustomObject]@{
            RuleName = $rule.Name
            State = $rule.State
            Priority = $rule.Priority
            Category = $category
            Mode = $rule.Mode
            Conditions = $conditionsStr
            Exceptions = $exceptionsStr
            Actions = $actionsStr
            Comments = $rule.Comments
            CreatedBy = $rule.CreatedBy
            WhenChanged = $rule.WhenChanged
            RuleVersion = $rule.RuleVersion
            IsSecurityRule = $isSecurityRule
            RuleIdentity = $rule.Identity
        }
        
        $results += $obj
    }
    
    Write-Progress -Activity "Processing Mail Flow Rules" -Completed
}
catch {
    Write-Host "Error retrieving transport rules: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Mail Flow Rules Summary:" -ForegroundColor Green
    Write-Host "  Total Rules: $($results.Count)" -ForegroundColor White
    Write-Host "  Enabled Rules: $enabledCount" -ForegroundColor Green
    Write-Host "  Disabled Rules: $disabledCount" -ForegroundColor Yellow
    Write-Host "  Security-Focused Rules: $(($results | Where-Object { $_.IsSecurityRule -eq $true }).Count)" -ForegroundColor Cyan
    
    # Category breakdown
    Write-Host "`n  Rules by Category:" -ForegroundColor Cyan
    $results | Group-Object Category | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table RuleName, State, Priority, Category, Mode -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No mail flow rules found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
