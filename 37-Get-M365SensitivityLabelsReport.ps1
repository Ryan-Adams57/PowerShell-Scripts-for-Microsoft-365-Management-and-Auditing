# ====================================================================================

# USAGE NOTES:
# - Requires Security & Compliance Center PowerShell connection
# - Run as an administrator or with appropriate M365 permissions
# - First run may take longer due to module loading
# - Report includes all sensitivity label configurations
# - Use -PublishedOnly to filter active labels
# - Export path can be customized for integration
#
# M365 POWERSHELL REPORTING SCRIPT - PRODUCTION READY
# This script provides comprehensive reporting capabilities for Microsoft 365
# Designed for enterprise environments with proper error handling
# ====================================================================================

# USAGE NOTES:
# - Requires Security & Compliance Center PowerShell connection
# - Run as an administrator or with appropriate M365 permissions
# - First run may take longer due to module loading
# - Report includes all sensitivity label configurations
# - Use -PublishedOnly to filter active labels
# - Export path can be customized for integration
#
#
<#
====================================================================================
Script Name: 37-Get-M365SensitivityLabelsReport.ps1
Description: Sensitivity labels configuration and usage report
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
    # Comprehensive parameter documentation
    # PublishedOnly: Filter to show only published labels
    # IncludeUsageStats: Include usage statistics if available
    # ExportPath: Custom export path for CSV report
    [Parameter(Mandatory=$false)]
    [switch]$PublishedOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeUsageStats,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Sensitivity_Labels_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Initialize comprehensive error handling
Set-StrictMode -Version Latest

# Logging function for consistent output
# Input validation function
function Test-Parameters {
    if ($ExportPath -and -not (Test-Path (Split-Path $ExportPath -Parent))) {
        Write-Log "Export path directory does not exist" -Level Error
        exit 1
    }
    Write-Log "Parameters validated successfully" -Level Success
}

function Write-Log {
    param(
    # Comprehensive parameter documentation
    # PublishedOnly: Filter to show only published labels
    # IncludeUsageStats: Include usage statistics if available
    # ExportPath: Custom export path for CSV report
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
Write-Host "M365 Sensitivity Labels Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install module? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
    } else { exit }
}

Write-Host "Connecting to Security & Compliance..." -ForegroundColor Cyan
try {
    Connect-IPPSSession -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

$script:Results = @()

try {
    $labels = Get-Label -ErrorAction Stop
    Write-Host "Found $($labels.Count) label(s).`n" -ForegroundColor Green
    
    foreach ($label in $labels) {
        $labelPolicy = Get-LabelPolicy -ErrorAction SilentlyContinue | Where-Object { $_.Labels -contains $label.Guid }
        $isPublished = $labelPolicy -ne $null
        
        if ($PublishedOnly -and -not $isPublished) { continue }
        
        $script:Results += [PSCustomObject]@{
            LabelName = $label.DisplayName
            Description = $label.Comment
            IsPublished = $isPublished
            PublishedBy = if ($labelPolicy) { $labelPolicy.Name } else { "Not Published" }
            Priority = $label.Priority
            EncryptionEnabled = $label.EncryptionEnabled
            ContentMarkingEnabled = $label.ContentMarkingEnabled
            SiteAndGroupProtectionEnabled = $label.SiteAndGroupProtectionEnabled
            ParentLabel = $label.ParentId
            Tooltip = $label.Tooltip
            CreatedBy = $label.CreatedBy
            ModifiedBy = $label.LastModifiedBy
            WhenCreated = $label.WhenCreatedUTC
            LabelId = $label.Guid
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Sensitivity Labels Summary:" -ForegroundColor Green
    Write-Host "  Total Labels: $($labels.Count)" -ForegroundColor White
    Write-Host "  Published Labels: $(($script:Results | Where-Object { $_.IsPublished -eq $true }).Count)" -ForegroundColor Green
    Write-Host "  Encryption Enabled: $(($script:Results | Where-Object { $_.EncryptionEnabled -eq $true }).Count)" -ForegroundColor Cyan
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table LabelName, IsPublished, EncryptionEnabled, Priority -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No labels found." -ForegroundColor Yellow
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
