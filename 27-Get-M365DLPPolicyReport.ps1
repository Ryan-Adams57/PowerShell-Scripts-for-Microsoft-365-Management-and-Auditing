<#
====================================================================================
Script Name: 27-Get-M365DLPPolicyReport.ps1
Description: Data Loss Prevention policy configuration and incidents report
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
    [ValidateSet("Exchange","SharePoint","OneDriveForBusiness","Teams","All")]
    [string]$Workload = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$EnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeIncidents,
    
    [Parameter(Mandatory=$false)]
    [datetime]$IncidentStartDate = (Get-Date).AddDays(-7),
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_DLP_Policy_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Data Loss Prevention Policy Report" -ForegroundColor Green
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

# Connect to Security & Compliance Center
Write-Host "Connecting to Security & Compliance Center..." -ForegroundColor Cyan

try {
    Connect-IPPSSession -ErrorAction Stop
    Write-Host "Successfully connected to Security & Compliance Center.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect. Error: $_" -ForegroundColor Red
    Write-Host "Attempting Exchange Online connection instead..." -ForegroundColor Yellow
    
    try {
        Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
        Write-Host "Connected to Exchange Online.`n" -ForegroundColor Green
    }
    catch {
        Write-Host "Connection failed. Error: $_" -ForegroundColor Red
        exit
    }
}

# Retrieve DLP policies
Write-Host "Retrieving DLP policies..." -ForegroundColor Cyan
$script:Results = @()
$enabledCount = 0
$disabledCount = 0

try {
    $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction Stop
    
    Write-Host "Found $($dlpPolicies.Count) DLP policy/policies. Analyzing configurations...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($policy in $dlpPolicies) {
        $progressCounter++
        Write-Progress -Activity "Processing DLP Policies" -Status "Policy $progressCounter of $($dlpPolicies.Count): $($policy.Name)" -PercentComplete (($progressCounter / $dlpPolicies.Count) * 100)
        
        # Filter by status
        if ($EnabledOnly -and $policy.Enabled -ne $true) {
            continue
        }
        
        # Count by status
        if ($policy.Enabled) {
            $enabledCount++
        } else {
            $disabledCount++
        }
        
        # Get policy rules
        $policyRules = Get-DlpComplianceRule -Policy $policy.Name -ErrorAction SilentlyContinue
        $ruleCount = if ($policyRules) { $policyRules.Count } else { 0 }
        
        # Parse workload locations
        $locations = @()
        if ($policy.ExchangeLocation) { $locations += "Exchange" }
        if ($policy.SharePointLocation) { $locations += "SharePoint" }
        if ($policy.OneDriveLocation) { $locations += "OneDrive" }
        if ($policy.TeamsLocation) { $locations += "Teams" }
        
        $locationsStr = $locations -join ", "
        
        # Filter by workload
        if ($Workload -ne "All") {
            if ($Workload -eq "OneDriveForBusiness") {
                if (-not $policy.OneDriveLocation) { continue }
            }
            elseif (-not $locations -contains $Workload) {
                continue
            }
        }
        
        # Get sensitive information types
        $sensitiveTypes = @()
        foreach ($rule in $policyRules) {
            if ($rule.ContentContainsSensitiveInformation) {
                foreach ($sensitiveInfo in $rule.ContentContainsSensitiveInformation) {
                    $sensitiveTypes += $sensitiveInfo.Name
                }
            }
        }
        $sensitiveTypesStr = ($sensitiveTypes | Select-Object -Unique) -join "; "
        
        # Parse actions
        $actions = @()
        foreach ($rule in $policyRules) {
            if ($rule.BlockAccess) { $actions += "BlockAccess" }
            if ($rule.NotifyUser) { $actions += "NotifyUser" }
            if ($rule.GenerateIncidentReport) { $actions += "GenerateIncidentReport" }
            if ($rule.GenerateAlert) { $actions += "GenerateAlert" }
        }
        $actionsStr = ($actions | Select-Object -Unique) -join "; "
        
        $obj = [PSCustomObject]@{
            PolicyName = $policy.Name
            Enabled = $policy.Enabled
            Mode = $policy.Mode
            Workloads = $locationsStr
            RuleCount = $ruleCount
            SensitiveInfoTypes = $sensitiveTypesStr
            Actions = $actionsStr
            Priority = $policy.Priority
            CreatedBy = $policy.CreatedBy
            ModifiedBy = $policy.ModifiedBy
            WhenCreated = $policy.WhenCreatedUTC
            WhenChanged = $policy.WhenChangedUTC
            Comment = $policy.Comment
            PolicyGuid = $policy.Guid
        }
        
        $script:Results += $obj
    }
    
    Write-Progress -Activity "Processing DLP Policies" -Completed
    
    # Get DLP incidents if requested
    if ($IncludeIncidents) {
        Write-Host "`nRetrieving DLP policy match incidents..." -ForegroundColor Cyan
        
        try {
            $incidents = Get-DlpDetailReport -StartDate $IncidentStartDate -EndDate (Get-Date) -ErrorAction SilentlyContinue
            
            if ($incidents) {
                Write-Host "Found $($incidents.Count) DLP incident(s) since $($IncidentStartDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor Green
                
                foreach ($incident in $incidents) {
                    $incidentObj = [PSCustomObject]@{
                        PolicyName = $incident.PolicyName
                        RuleName = $incident.RuleName
                        IncidentDate = $incident.Date
                        User = $incident.User
                        Workload = $incident.Workload
                        FileName = $incident.FileName
                        Location = $incident.Location
                        SensitiveType = $incident.SensitiveType
                        Action = $incident.Action
                        IncidentId = $incident.IncidentId
                    }
                    
                    $script:Results += $incidentObj
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve DLP incidents: $_"
        }
    }
}
catch {
    Write-Host "Error retrieving DLP policies: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit
}

# Export and display results
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "DLP Policy Summary:" -ForegroundColor Green
    Write-Host "  Total Policies: $($dlpPolicies.Count)" -ForegroundColor White
    Write-Host "  Enabled Policies: $enabledCount" -ForegroundColor Green
    Write-Host "  Disabled Policies: $disabledCount" -ForegroundColor Yellow
    
    if ($IncludeIncidents) {
        Write-Host "  DLP Incidents Included: Yes" -ForegroundColor Cyan
    }
    
    # Workload distribution
    Write-Host "`n  Policies by Workload:" -ForegroundColor Cyan
    $script:Results | Where-Object { $_.PolicyName } | Group-Object { 
        if ($_.Workloads -like "*Exchange*") { "Exchange" } 
        elseif ($_.Workloads -like "*SharePoint*") { "SharePoint" }
        elseif ($_.Workloads -like "*OneDrive*") { "OneDrive" }
        elseif ($_.Workloads -like "*Teams*") { "Teams" }
        else { "Other" }
    } | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "COMPLIANCE NOTE:" -ForegroundColor Red
    Write-Host "Review DLP policies regularly to ensure data protection compliance.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $script:Results | Select-Object -First 10 | Format-Table PolicyName, Enabled, Mode, Workloads, RuleCount -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No DLP policies found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from services..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
