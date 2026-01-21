<#
====================================================================================
Script Name: Get-M365RetentionPoliciesReport.ps1
Description: Retention policies and retention labels configuration report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all retention policies and labels
• Shows policy scope (Exchange, SharePoint, OneDrive, Teams)
• Lists retention periods and disposition actions
• Identifies policy assignments and exclusions
• Tracks adaptive vs static policy scopes
• Supports filtering by workload
• Generates compliance and governance reports
• Requires Compliance Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Exchange","SharePoint","OneDrive","Teams","All")]
    [string]$Workload = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeLabels,
    
    [Parameter(Mandatory=$false)]
    [switch]$EnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Retention_Policies_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "M365 Retention Policies and Labels Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
    } else { exit }
}

Write-Host "Connecting to Security & Compliance Center..." -ForegroundColor Cyan
try {
    Connect-IPPSSession -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

$results = @()

try {
    $policies = Get-RetentionCompliancePolicy -ErrorAction Stop
    
    foreach ($policy in $policies) {
        if ($EnabledOnly -and -not $policy.Enabled) { continue }
        
        $locations = @()
        if ($policy.ExchangeLocation) { $locations += "Exchange" }
        if ($policy.SharePointLocation) { $locations += "SharePoint" }
        if ($policy.OneDriveLocation) { $locations += "OneDrive" }
        if ($policy.TeamsChannelLocation) { $locations += "Teams" }
        
        $locationsStr = $locations -join ", "
        
        if ($Workload -ne "All" -and -not ($locations -contains $Workload)) { continue }
        
        $rules = Get-RetentionComplianceRule -Policy $policy.Name -ErrorAction SilentlyContinue
        
        $retentionDuration = "Not Set"
        $retentionAction = "Not Set"
        
        if ($rules) {
            $retentionDuration = $rules[0].RetentionDuration
            $retentionAction = $rules[0].RetentionComplianceAction
        }
        
        $results += [PSCustomObject]@{
            PolicyName = $policy.Name
            Enabled = $policy.Enabled
            Mode = $policy.Mode
            Workloads = $locationsStr
            RetentionDuration = $retentionDuration
            RetentionAction = $retentionAction
            IsAdaptive = $policy.IsAdaptiveScope
            CreatedBy = $policy.CreatedBy
            WhenCreated = $policy.WhenCreatedUTC
            WhenChanged = $policy.WhenChangedUTC
        }
    }
    
    if ($IncludeLabels) {
        $labels = Get-ComplianceTag -ErrorAction SilentlyContinue
        
        foreach ($label in $labels) {
            $results += [PSCustomObject]@{
                PolicyName = $label.Name
                Enabled = "N/A"
                Mode = "Label"
                Workloads = "Label"
                RetentionDuration = $label.RetentionDuration
                RetentionAction = $label.RetentionAction
                IsAdaptive = $false
                CreatedBy = $label.CreatedBy
                WhenCreated = $label.WhenCreatedUTC
                WhenChanged = $label.WhenChangedUTC
            }
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Retention Summary:" -ForegroundColor Green
    Write-Host "  Total Policies: $(($results | Where-Object { $_.Mode -ne 'Label' }).Count)" -ForegroundColor White
    if ($IncludeLabels) {
        Write-Host "  Total Labels: $(($results | Where-Object { $_.Mode -eq 'Label' }).Count)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table PolicyName, Enabled, Workloads, RetentionDuration -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No retention policies found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
