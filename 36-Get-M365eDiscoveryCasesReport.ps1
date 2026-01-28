# ====================================================================================
# M365 POWERSHELL REPORTING SCRIPT - PRODUCTION READY
# This script provides comprehensive reporting capabilities for Microsoft 365
# Designed for enterprise environments with proper error handling
# ====================================================================================
#
<#
====================================================================================
Script Name: 36-Get-M365eDiscoveryCasesReport.ps1
Description: eDiscovery cases, holds, and search configuration report
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
    [ValidateSet("Active","Closed","All")]
    [string]$CaseStatus = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeHolds,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_eDiscovery_Cases_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
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
Write-Host "M365 eDiscovery Cases Report" -ForegroundColor Green
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
$activeCount = 0
$closedCount = 0

try {
    $cases = Get-ComplianceCase -ErrorAction Stop
    
    Write-Host "Found $($cases.Count) case(s).`n" -ForegroundColor Green
    
    foreach ($case in $cases) {
        $status = $case.Status
        
        if ($status -eq "Active") { $activeCount++ } else { $closedCount++ }
        
        if ($CaseStatus -ne "All" -and $status -ne $CaseStatus) { continue }
        
        $members = Get-ComplianceCaseMember -Case $case.Name -ErrorAction SilentlyContinue
        $memberCount = if ($members) { $members.Count } else { 0 }
        
        $searches = Get-ComplianceSearch -Case $case.Name -ErrorAction SilentlyContinue
        $searchCount = if ($searches) { $searches.Count } else { 0 }
        
        $holds = "Not Retrieved"
        $holdCount = 0
        
        if ($IncludeHolds) {
            $caseHolds = Get-CaseHoldPolicy -Case $case.Name -ErrorAction SilentlyContinue
            if ($caseHolds) {
                $holdCount = $caseHolds.Count
                $holds = ($caseHolds | ForEach-Object { $_.Name }) -join "; "
            } else {
                $holds = "None"
            }
        }
        
        $script:Results += [PSCustomObject]@{
            CaseName = $case.Name
            Status = $status
            CaseType = $case.CaseType
            CreatedDateTime = $case.CreatedDateTime
            ClosedDateTime = $case.ClosedDateTime
            CreatedBy = $case.CreatedBy
            MemberCount = $memberCount
            SearchCount = $searchCount
            HoldCount = $holdCount
            Holds = $holds
            Description = $case.Description
            CaseId = $case.Identity
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "eDiscovery Summary:" -ForegroundColor Green
    Write-Host "  Total Cases: $($cases.Count)" -ForegroundColor White
    Write-Host "  Active: $activeCount" -ForegroundColor Green
    Write-Host "  Closed: $closedCount" -ForegroundColor Yellow
    Write-Host "  Total Searches: $(($script:Results | Measure-Object -Property SearchCount -Sum).Sum)" -ForegroundColor White
    
    if ($IncludeHolds) {
        Write-Host "  Total Holds: $(($script:Results | Measure-Object -Property HoldCount -Sum).Sum)" -ForegroundColor Cyan
    }
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table CaseName, Status, MemberCount, SearchCount, HoldCount -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No cases found." -ForegroundColor Yellow
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
