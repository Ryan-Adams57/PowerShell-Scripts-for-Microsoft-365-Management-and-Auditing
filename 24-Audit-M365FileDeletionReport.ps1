<#
====================================================================================
Script Name: 24-Audit-M365FileDeletionReport.ps1
Description: File and Document Deletion Audit Report
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
    [datetime]$StartDate = (Get-Date).AddDays(-7),
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date),
    
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_File_Deletion_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Initialize comprehensive error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:Results = @()
$script:ErrorCount = 0

# Display script information
Write-Host "Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
Write-Host "" -ForegroundColor Cyan


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 File Deletion Audit Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Module not installed." -ForegroundColor Yellow
    $install = Read-Host "Install? (Y/N)"
    
    if ($install -eq 'Y' -or $install -eq 'y') {
        try {
            Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
            Write-Host "Installed.`n" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed: $_" -ForegroundColor Red
            exit
        }
    }
    else {
        exit
    }
}

# Validate dates
if ($StartDate -gt $EndDate) {
    Write-Host "ERROR: Start date after end date." -ForegroundColor Red
    exit
}

$dateRange = (New-TimeSpan -Start $StartDate -End $EndDate).Days
if ($dateRange -gt 90) {
    Write-Host "WARNING: Adjusting to 90-day limit." -ForegroundColor Yellow
    $StartDate = (Get-Date).AddDays(-90)
}

Write-Host "Range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor Cyan

# Connect
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Search
Write-Host "Searching for file deletion events..." -ForegroundColor Cyan
$script:Results = @()

try {
    $searchParams = @{
        StartDate = $StartDate
        EndDate = $EndDate
        Operations = "FileDeleted,FileRecycled"
        RecordType = "SharePointFileOperation"
        ResultSize = 5000
    }
    
    if ($UserPrincipalName) {
        $searchParams.Add("UserIds", $UserPrincipalName)
    }
    
    if ($SiteUrl) {
        $searchParams.Add("SiteIds", $SiteUrl)
    }
    
    $auditRecords = Search-UnifiedAuditLog @searchParams
    
    Write-Host "Found $($auditRecords.Count) deletion record(s).`n" -ForegroundColor Green
    
    foreach ($record in $auditRecords) {
        try {
            $auditData = $record.AuditData | ConvertFrom-Json
            
            $script:Results += [PSCustomObject]@{
                TimeStamp = $record.CreationDate
                User = $record.UserIds
                Operation = $record.Operations
                FileName = $auditData.SourceFileName
                FileExtension = $auditData.SourceFileExtension
                SiteUrl = $auditData.SiteUrl
                SourceRelativeUrl = $auditData.SourceRelativeUrl
                ItemType = $auditData.ItemType
                ClientIP = $auditData.ClientIP
                UserAgent = $auditData.UserAgent
                Workload = $auditData.Workload
            }
        }
        catch {
            Write-Warning "Parse error: $_"
        }
    }
}
catch {
    Write-Host "Search error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Deletion Audit Summary:" -ForegroundColor Green
    Write-Host "  Total Deletions: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Unique Users: $(($script:Results | Select-Object -Unique User).Count)" -ForegroundColor White
    Write-Host "  Unique Sites: $(($script:Results | Select-Object -Unique SiteUrl).Count)" -ForegroundColor White
    
    Write-Host "`n  Top Deleters:" -ForegroundColor Cyan
    $script:Results | Group-Object User | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count) files" -ForegroundColor White
    }
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table TimeStamp, User, FileName, SiteUrl -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No file deletions found." -ForegroundColor Yellow
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
