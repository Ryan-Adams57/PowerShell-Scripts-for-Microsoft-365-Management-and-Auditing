<#
====================================================================================
Script Name: 25-Audit-M365AdminActivityReport.ps1
Description: Administrator Activity and Privileged Actions Audit
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
    [string]$AdminUser,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("AzureActiveDirectory","Exchange","SharePoint","Teams","All")]
    [string]$Workload = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Admin_Activity_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Administrator Activity Audit" -ForegroundColor Green
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
    Write-Host "ERROR: Invalid date range." -ForegroundColor Red
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

# Admin operations list
$adminOps = @(
    "Add member to role",
    "Add-RoleGroupMember",
    "Set-Mailbox",
    "Set-User",
    "Set-MsolUser",
    "New-TransportRule",
    "Set-TransportRule",
    "Remove-TransportRule",
    "New-InboundConnector",
    "Set-OrganizationConfig",
    "Set-SharingPolicy",
    "Add-MailboxPermission",
    "Add-RecipientPermission",
    "Set-CASMailbox"
)

# Search
Write-Host "Searching for admin activities..." -ForegroundColor Cyan
$script:Results = @()

try {
    $searchParams = @{
        StartDate = $StartDate
        EndDate = $EndDate
        ResultSize = 5000
    }
    
    if ($AdminUser) {
        $searchParams.Add("UserIds", $AdminUser)
    }
    
    if ($Workload -ne "All") {
        $searchParams.Add("RecordType", $Workload)
    }
    
    $auditRecords = Search-UnifiedAuditLog @searchParams
    
    Write-Host "Found $($auditRecords.Count) audit record(s). Filtering admin activities...`n" -ForegroundColor Green
    
    foreach ($record in $auditRecords) {
        try {
            $auditData = $record.AuditData | ConvertFrom-Json
            
            # Filter for admin operations
            $isAdminOp = $false
            foreach ($op in $adminOps) {
                if ($record.Operations -like "*$op*") {
                    $isAdminOp = $true
                    break
                }
            }
            
            if (-not $isAdminOp -and $auditData.UserType -ne "Admin") {
                continue
            }
            
            $script:Results += [PSCustomObject]@{
                TimeStamp = $record.CreationDate
                AdminUser = $record.UserIds
                Operation = $record.Operations
                Workload = $auditData.Workload
                TargetObject = $auditData.ObjectId
                ClientIP = $auditData.ClientIP
                UserAgent = $auditData.UserAgent
                ResultStatus = $auditData.ResultStatus
                Parameters = if ($auditData.Parameters) { ($auditData.Parameters | ConvertTo-Json -Compress) } else { "N/A" }
                AuditDataJson = $record.AuditData
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
    Write-Host "Admin Activity Summary:" -ForegroundColor Green
    Write-Host "  Total Admin Actions: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Unique Admins: $(($script:Results | Select-Object -Unique AdminUser).Count)" -ForegroundColor White
    Write-Host "  Unique Operations: $(($script:Results | Select-Object -Unique Operation).Count)" -ForegroundColor White
    
    Write-Host "`n  Top Admin Operations:" -ForegroundColor Cyan
    $script:Results | Group-Object Operation | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    Write-Host "`n  Top Administrators:" -ForegroundColor Cyan
    $script:Results | Group-Object AdminUser | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count) actions" -ForegroundColor White
    }
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY NOTE:" -ForegroundColor Red
    Write-Host "Review admin activities regularly for unauthorized changes.`n" -ForegroundColor Yellow
    
    $script:Results | Select-Object -First 10 | Format-Table TimeStamp, AdminUser, Operation, Workload, ResultStatus -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No admin activities found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
