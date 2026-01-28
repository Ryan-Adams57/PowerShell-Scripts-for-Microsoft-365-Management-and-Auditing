<#
====================================================================================
Script Name: 34-Get-AzureADConditionalAccessSignInLogs.ps1
Description: Azure AD Conditional Access sign-in logs and policy evaluations
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
    [ValidateSet("Success","Failure","All")]
    [string]$SignInStatus = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$FailedOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\AzureAD_CA_SignIn_Logs_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Azure AD Conditional Access Sign-In Logs Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Reports"

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

# Connect
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "AuditLog.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve sign-in logs
Write-Host "Retrieving Azure AD sign-in logs..." -ForegroundColor Cyan
Write-Host "Date range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))`n" -ForegroundColor White

$script:Results = @()
$successCount = 0
$failureCount = 0
$caAppliedCount = 0

try {
    $filter = "createdDateTime ge $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and createdDateTime le $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
    
    if ($UserPrincipalName) {
        $filter += " and userPrincipalName eq '$UserPrincipalName'"
    }
    
    $signIns = Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop
    
    Write-Host "Found $($signIns.Count) sign-in record(s). Processing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($signIn in $signIns) {
        $progressCounter++
        Write-Progress -Activity "Processing Sign-Ins" -Status "Record $progressCounter of $($signIns.Count)" -PercentComplete (($progressCounter / $signIns.Count) * 100)
        
        $status = if ($signIn.Status.ErrorCode -eq 0) { "Success" } else { "Failure" }
        
        if ($status -eq "Success") { $successCount++ } else { $failureCount++ }
        
        if ($FailedOnly -and $status -eq "Success") { continue }
        if ($SignInStatus -ne "All" -and $status -ne $SignInStatus) { continue }
        
        # Parse Conditional Access details
        $caPolicies = @()
        $caResult = "Not Applied"
        
        if ($signIn.ConditionalAccessPolicies) {
            $caAppliedCount++
            foreach ($policy in $signIn.ConditionalAccessPolicies) {
                $caPolicies += "$($policy.DisplayName) ($($policy.Result))"
            }
            $caResult = $caPolicies -join "; "
        }
        
        $obj = [PSCustomObject]@{
            CreatedDateTime = $signIn.CreatedDateTime
            UserPrincipalName = $signIn.UserPrincipalName
            UserDisplayName = $signIn.UserDisplayName
            AppDisplayName = $signIn.AppDisplayName
            ClientAppUsed = $signIn.ClientAppUsed
            DeviceDetailBrowser = $signIn.DeviceDetail.Browser
            DeviceDetailOS = $signIn.DeviceDetail.OperatingSystem
            Location = $signIn.Location.City
            Country = $signIn.Location.CountryOrRegion
            IPAddress = $signIn.IpAddress
            SignInStatus = $status
            ErrorCode = $signIn.Status.ErrorCode
            FailureReason = $signIn.Status.FailureReason
            ConditionalAccessPolicies = $caResult
            IsInteractive = $signIn.IsInteractive
            RiskLevel = $signIn.RiskLevelDuringSignIn
            RiskState = $signIn.RiskState
            CorrelationId = $signIn.CorrelationId
        }
        
        $script:Results += $obj
    }
    
    Write-Progress -Activity "Processing Sign-Ins" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Sign-In Log Summary:" -ForegroundColor Green
    Write-Host "  Total Sign-Ins: $($signIns.Count)" -ForegroundColor White
    Write-Host "  Successful: $successCount" -ForegroundColor Green
    Write-Host "  Failed: $failureCount" -ForegroundColor Red
    Write-Host "  With CA Policies Applied: $caAppliedCount" -ForegroundColor Cyan
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table CreatedDateTime, UserPrincipalName, AppDisplayName, SignInStatus, IPAddress -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No sign-in logs found." -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
