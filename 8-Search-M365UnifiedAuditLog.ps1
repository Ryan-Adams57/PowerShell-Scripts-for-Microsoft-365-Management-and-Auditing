<#
====================================================================================
Script Name: 8-Search-M365UnifiedAuditLog.ps1
Description: Production-ready M365 reporting script
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

REQUIREMENTS:
• PowerShell 5.1 or higher
• Appropriate M365 administrator permissions
• Required modules (will be validated at runtime)

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
    [string]$UserIds,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Exchange","SharePoint","OneDrive","AzureActiveDirectory","Yammer","Teams","All")]
    [string]$Workload = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$Operations,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("AzureActiveDirectory","Exchange","SharePoint","OneDrive","Skype","Yammer","SecurityComplianceCenter","ThreatIntelligence","All")]
    [string]$RecordType,
    
    [Parameter(Mandatory=$false)]
    [int]$ResultSize = 5000,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Audit_Log_Search_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Unified Audit Log Search Tool" -ForegroundColor Green
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

# Validate date range
$maxRetentionDays = 90
$dateRangeDays = (New-TimeSpan -Start $StartDate -End $EndDate).Days

if ($dateRangeDays -gt $maxRetentionDays) {
    Write-Host "WARNING: Date range exceeds maximum audit log retention ($maxRetentionDays days)." -ForegroundColor Yellow
    Write-Host "Adjusting start date to $maxRetentionDays days ago.`n" -ForegroundColor Yellow
    $StartDate = (Get-Date).AddDays(-$maxRetentionDays)
}

if ($StartDate -gt $EndDate) {
    Write-Host "ERROR: Start date cannot be after end date." -ForegroundColor Red
    exit
}

Write-Host "Search Parameters:" -ForegroundColor Cyan
Write-Host "  Start Date: $($StartDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "  End Date: $($EndDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "  Date Range: $dateRangeDays days`n" -ForegroundColor White

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online (required for Unified Audit Log access)..." -ForegroundColor Cyan

try {
    # UseRPSSession avoids authentication broker issues
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online. Error: $_" -ForegroundColor Red
    Write-Host "Tip: Ensure ExchangeOnlineManagement module is up to date: Update-Module ExchangeOnlineManagement" -ForegroundColor Yellow
    exit
}

# Build search parameters
$searchParams = @{
    StartDate = $StartDate
    EndDate = $EndDate
    ResultSize = if ($ResultSize -gt 5000) { 5000 } else { $ResultSize }
}

if ($UserIds) {
    $searchParams.Add("UserIds", $UserIds)
}

if ($Operations) {
    $searchParams.Add("Operations", $Operations)
}

if ($RecordType -and $RecordType -ne "All") {
    $searchParams.Add("RecordType", $RecordType)
}

# Execute search
Write-Host "Searching Unified Audit Log..." -ForegroundColor Cyan
Write-Host "Note: Large searches may take several minutes.`n" -ForegroundColor Yellow

$results = @()
$sessionId = [Guid]::NewGuid().ToString()
$searchParams.Add("SessionId", $sessionId)
$searchParams.Add("SessionCommand", "ReturnLargeSet")

try {
    $totalRecords = 0
    $batchNumber = 1
    
    do {
        Write-Progress -Activity "Searching Audit Logs" -Status "Retrieving batch $batchNumber (Total records: $totalRecords)" -PercentComplete -1
        
        $auditData = Search-UnifiedAuditLog @searchParams
        
        if ($auditData) {
            $totalRecords += $auditData.Count
            
            foreach ($record in $auditData) {
                try {
                    $auditInfo = $record.AuditData | ConvertFrom-Json
                    
                    $obj = [PSCustomObject]@{
                        CreationDate = $record.CreationDate
                        UserIds = $record.UserIds
                        Operations = $record.Operations
                        RecordType = $record.RecordType
                        Workload = $auditInfo.Workload
                        ObjectId = $auditInfo.ObjectId
                        UserId = $auditInfo.UserId
                        ClientIP = $auditInfo.ClientIP
                        UserAgent = $auditInfo.UserAgent
                        SiteUrl = $auditInfo.SiteUrl
                        SourceFileName = $auditInfo.SourceFileName
                        DestinationFileName = $auditInfo.DestinationFileName
                        ItemType = $auditInfo.ItemType
                        EventSource = $auditInfo.EventSource
                        ResultStatus = $auditInfo.ResultStatus
                        AuditDataJson = $record.AuditData
                    }
                    
                    # Apply workload filter
                    if ($Workload -eq "All" -or $obj.Workload -eq $Workload) {
                        $results += $obj
                    }
                }
                catch {
                    Write-Warning "Error parsing audit record: $_"
                }
            }
            
            $batchNumber++
        }
        
    } while ($auditData -and $auditData.Count -eq $searchParams.ResultSize -and $totalRecords -lt 50000)
    
    Write-Progress -Activity "Searching Audit Logs" -Completed
    
    Write-Host "Search completed. Found $totalRecords total audit records.`n" -ForegroundColor Green
}
catch {
    Write-Host "Error searching audit log: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Audit Log Search Summary:" -ForegroundColor Green
    Write-Host "  Total Records Retrieved: $($results.Count)" -ForegroundColor White
    Write-Host "  Unique Users: $(($results | Select-Object -Unique UserId).Count)" -ForegroundColor White
    Write-Host "  Unique Operations: $(($results | Select-Object -Unique Operations).Count)" -ForegroundColor White
    Write-Host "  Date Range Searched: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
    
    # Workload breakdown
    Write-Host "`n  Records by Workload:" -ForegroundColor Cyan
    $results | Group-Object Workload | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Sample Audit Records (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table CreationDate, UserIds, Operations, Workload -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No audit records found matching the search criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
