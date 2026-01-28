<#
====================================================================================
Script Name: 21-Audit-M365LicenseAssignmentChanges.ps1
Description: License Assignment Change Tracking and Audit Report
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
    [datetime]$StartDate = (Get-Date).AddDays(-30),
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date),
    
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$AdminUser,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_License_Changes_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
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
Write-Host "Microsoft 365 License Assignment Changes Audit" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Module '$requiredModule' not installed." -ForegroundColor Yellow
    $install = Read-Host "Install now? (Y/N)"
    
    if ($install -eq 'Y' -or $install -eq 'y') {
        try {
            Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
            Write-Host "Installed successfully.`n" -ForegroundColor Green
        }
        catch {
            Write-Host "Installation failed: $_" -ForegroundColor Red
            exit
        }
    }
    else {
        exit
    }
}

# Validate dates
if ($StartDate -gt $EndDate) {
    Write-Host "ERROR: Start date cannot be after end date." -ForegroundColor Red
    exit
}

$dateRange = (New-TimeSpan -Start $StartDate -End $EndDate).Days
if ($dateRange -gt 90) {
    Write-Host "WARNING: Audit log retention is 90 days. Adjusting start date." -ForegroundColor Yellow
    $StartDate = (Get-Date).AddDays(-90)
}

Write-Host "Search range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor Cyan

# Connect
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Connection failed: $_" -ForegroundColor Red
    exit
}

# Search audit log
Write-Host "Searching license change audit records..." -ForegroundColor Cyan
$script:Results = @()

try {
    $searchParams = @{
        StartDate = $StartDate
        EndDate = $EndDate
        Operations = "Add-MsolRoleMember,Remove-MsolRoleMember,Set-MsolUserLicense"
        ResultSize = 5000
    }
    
    if ($UserPrincipalName) {
        $searchParams.Add("ObjectIds", $UserPrincipalName)
    }
    
    if ($AdminUser) {
        $searchParams.Add("UserIds", $AdminUser)
    }
    
    $auditRecords = Search-UnifiedAuditLog @searchParams
    
    Write-Host "Found $($auditRecords.Count) audit record(s).`n" -ForegroundColor Green
    
    foreach ($record in $auditRecords) {
        try {
            $auditData = $record.AuditData | ConvertFrom-Json
            
            $script:Results += [PSCustomObject]@{
                TimeStamp = $record.CreationDate
                AdminUser = $record.UserIds
                Operation = $record.Operations
                TargetUser = $auditData.ObjectId
                Workload = $auditData.Workload
                ClientIP = $auditData.ClientIP
                ResultStatus = $auditData.ResultStatus
                AuditDataJson = $record.AuditData
            }
        }
        catch {
            Write-Warning "Error parsing record: $_"
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
    Write-Host "Audit Summary:" -ForegroundColor Green
    Write-Host "  Total Changes: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Unique Admins: $(($script:Results | Select-Object -Unique AdminUser).Count)" -ForegroundColor White
    Write-Host "  Unique Users Affected: $(($script:Results | Select-Object -Unique TargetUser).Count)" -ForegroundColor White
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table TimeStamp, AdminUser, Operation, TargetUser -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No license changes found in the specified period." -ForegroundColor Yellow
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
