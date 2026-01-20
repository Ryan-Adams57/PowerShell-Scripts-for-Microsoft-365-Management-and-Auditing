<#
====================================================================================
Script Name: Audit-M365FileDeletionReport.ps1
Description: File and Document Deletion Audit Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Searches Unified Audit Log for file deletion events
• Shows who deleted files, when, and from where
• Lists deleted file names and SharePoint/OneDrive locations
• Identifies bulk file deletion activities
• Supports date range and user filtering
• Generates forensic-ready audit reports
• Exports deletion history and evidence CSV
• Critical for data loss investigations and compliance

====================================================================================
#>

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
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Search
Write-Host "Searching for file deletion events..." -ForegroundColor Cyan
$results = @()

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
            
            $results += [PSCustomObject]@{
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
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Deletion Audit Summary:" -ForegroundColor Green
    Write-Host "  Total Deletions: $($results.Count)" -ForegroundColor White
    Write-Host "  Unique Users: $(($results | Select-Object -Unique User).Count)" -ForegroundColor White
    Write-Host "  Unique Sites: $(($results | Select-Object -Unique SiteUrl).Count)" -ForegroundColor White
    
    Write-Host "`n  Top Deleters:" -ForegroundColor Cyan
    $results | Group-Object User | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count) files" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table TimeStamp, User, FileName, SiteUrl -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No file deletions found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
