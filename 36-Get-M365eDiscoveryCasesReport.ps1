<#
====================================================================================
Script Name: Get-M365eDiscoveryCasesReport.ps1
Description: eDiscovery cases, holds, and search configuration report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all eDiscovery cases (Standard and Advanced)
• Shows case status, members, and custodians
• Lists content searches and hold policies
• Identifies active vs closed cases
• Tracks case creation and modification dates
• Supports filtering by case status
• Generates legal and compliance audit reports
• Requires eDiscovery Manager or Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Active","Closed","All")]
    [string]$CaseStatus = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeHolds,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_eDiscovery_Cases_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

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

$results = @()
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
        
        $results += [PSCustomObject]@{
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

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "eDiscovery Summary:" -ForegroundColor Green
    Write-Host "  Total Cases: $($cases.Count)" -ForegroundColor White
    Write-Host "  Active: $activeCount" -ForegroundColor Green
    Write-Host "  Closed: $closedCount" -ForegroundColor Yellow
    Write-Host "  Total Searches: $(($results | Measure-Object -Property SearchCount -Sum).Sum)" -ForegroundColor White
    
    if ($IncludeHolds) {
        Write-Host "  Total Holds: $(($results | Measure-Object -Property HoldCount -Sum).Sum)" -ForegroundColor Cyan
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table CaseName, Status, MemberCount, SearchCount, HoldCount -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No cases found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
