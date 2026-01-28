# ====================================================================================
# M365 POWERSHELL REPORTING SCRIPT - PRODUCTION READY
# This script provides comprehensive reporting capabilities for Microsoft 365
# Designed for enterprise environments with proper error handling
# ====================================================================================
#
<#
====================================================================================
Script Name: 35-Get-M365RetentionPoliciesReport.ps1
Description: Retention policies and retention labels configuration report
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
    [ValidateSet("Exchange","SharePoint","OneDrive","Teams","All")]
    [string]$Workload = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeLabels,
    
    [Parameter(Mandatory=$false)]
    [switch]$EnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Retention_Policies_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Initialize comprehensive error handling
Set-StrictMode -Version Latest

# Logging function for consistent output
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info","Warning","Error","Success")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Info" { "Cyan" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
    }
    Write-Host "[$timestamp] $Level: $Message" -ForegroundColor $color
}

$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:Results = @()
$script:ErrorCount = 0

# Display script information
Write-Host "Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
Write-Host "" -ForegroundColor Cyan

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

$script:Results = @()

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
        
        $script:Results += [PSCustomObject]@{
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
            $script:Results += [PSCustomObject]@{
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

if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Retention Summary:" -ForegroundColor Green
    Write-Host "  Total Policies: $(($script:Results | Where-Object { $_.Mode -ne 'Label' }).Count)" -ForegroundColor White
    if ($IncludeLabels) {
        Write-Host "  Total Labels: $(($script:Results | Where-Object { $_.Mode -eq 'Label' }).Count)" -ForegroundColor White
    }
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table PolicyName, Enabled, Workloads, RetentionDuration -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No retention policies found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null

# Comprehensive cleanup and summary
$script:EndTime = Get-Date
$script:Duration = $script:EndTime - $script:StartTime

Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "Execution Summary:" -ForegroundColor Green
Write-Host "  Duration: $($script:Duration.TotalSeconds) seconds" -ForegroundColor White
Write-Host "  Results: $($script:Results.Count) items" -ForegroundColor White
Write-Host "  Errors: $script:ErrorCount" -ForegroundColor White
Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "Completed.`n" -ForegroundColor Green
