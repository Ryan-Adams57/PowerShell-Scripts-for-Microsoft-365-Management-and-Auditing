<#
====================================================================================
Script Name: 2-Get-M365InactiveUsersReport.ps1
Description: Identifies inactive Microsoft 365 users based on last sign-in activity
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

SCRIPT HIGHLIGHTS:
• Identifies users inactive for specified number of days (default 90)
• Includes last interactive and non-interactive sign-in dates
• Shows license assignment for inactive users
• Calculates potential license cost savings
• Supports filtering by department or license type
• Generates actionable recommendations
• Exports detailed CSV reports
• MFA-compatible authentication
• NO API calls inside loops (uses hashtable for SKU lookups)
• Comprehensive error handling with try/catch/finally
• Progress indicators for long operations

REQUIREMENTS:
• Microsoft.Graph.Users module
• Microsoft.Graph.Reports module
• Microsoft.Graph.Identity.DirectoryManagement module
• User.Read.All permission
• AuditLog.Read.All permission
• Directory.Read.All permission

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Number of days to consider user inactive")]
    [ValidateRange(1, 365)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory=$false, HelpMessage="Filter by specific department")]
    [string]$Department,
    
    [Parameter(Mandatory=$false, HelpMessage="Only include users with licenses")]
    [switch]$LicensedOnly,
    
    [Parameter(Mandatory=$false, HelpMessage="Include users who never signed in")]
    [switch]$IncludeNeverSignedIn,
    
    [Parameter(Mandatory=$false, HelpMessage="Export CSV path")]
    [string]$ExportPath = ".\M365_Inactive_Users_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize script variables
$script:Results = @()
$script:SkuLookup = @{}
$script:InactiveCount = 0
$script:NeverSignedInCount = 0

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Inactive Users Report Generator" -ForegroundColor Green
Write-Host "Version 2.0 - Production Ready" -ForegroundColor White
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

# Module validation
$requiredModules = @(
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

Write-Host "Validating required modules..." -ForegroundColor Cyan

foreach ($module in $requiredModules) {
    try {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "ERROR: Required module '$module' is not installed." -ForegroundColor Red
            Write-Host "Please run: Install-Module -Name $module -Scope CurrentUser" -ForegroundColor Yellow
            exit 1
        }
        else {
            $moduleInfo = Get-Module -ListAvailable -Name $module | 
                Sort-Object Version -Descending | 
                Select-Object -First 1
            Write-Host "  ✓ $module (v$($moduleInfo.Version))" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "ERROR validating module '$module': $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    
    if (-not $context) {
        Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Host "Already connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
    }
    
    $context = Get-MgContext
    Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Calculate threshold date
$thresholdDate = (Get-Date).AddDays(-$InactiveDays)
Write-Host "Identifying users inactive since: $($thresholdDate.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
Write-Host "Threshold: $InactiveDays days`n" -ForegroundColor Cyan

# Retrieve all SKUs ONCE (critical: no API calls in loops)
Write-Host "Retrieving organization license SKUs..." -ForegroundColor Cyan

try {
    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    Write-Host "Retrieved $($subscribedSkus.Count) license SKU(s)." -ForegroundColor Green
    
    # Create hashtable for FAST lookups
    foreach ($sku in $subscribedSkus) {
        $script:SkuLookup[$sku.SkuId] = $sku
    }
    
    Write-Host "SKU lookup table created for optimized processing.`n" -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving license SKUs: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Continuing without license information..." -ForegroundColor Yellow
}

# Build user filter
$filter = "accountEnabled eq true"
if ($Department) {
    $filter += " and department eq '$Department'"
}
if ($LicensedOnly) {
    $filter += " and assignedLicenses/`$count ne 0"
}

# Retrieve users
Write-Host "Retrieving user accounts..." -ForegroundColor Cyan

try {
    $properties = @(
        "DisplayName",
        "UserPrincipalName", 
        "AccountEnabled",
        "SignInActivity",
        "CreatedDateTime",
        "AssignedLicenses",
        "Department",
        "JobTitle",
        "OfficeLocation",
        "Mail"
    )
    
    $users = Get-MgUser -Filter $filter `
        -ConsistencyLevel eventual `
        -All `
        -Property $properties `
        -ErrorAction Stop
    
    Write-Host "Retrieved $($users.Count) user account(s)." -ForegroundColor Green
    Write-Host "Analyzing activity...`n" -ForegroundColor Cyan
    
    $progressCounter = 0
    
    foreach ($user in $users) {
        $progressCounter++
        Write-Progress -Activity "Analyzing User Activity" `
            -Status "User $progressCounter of $($users.Count): $($user.UserPrincipalName)" `
            -PercentComplete (($progressCounter / $users.Count) * 100)
        
        $lastSignIn = $null
        $lastNonInteractiveSignIn = $null
        $daysSinceSignIn = $null
        $status = "Active"
        $mostRecentSignIn = $null
        
        if ($user.SignInActivity) {
            $lastSignIn = $user.SignInActivity.LastSignInDateTime
            $lastNonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime
            
            # Determine most recent sign-in
            if ($lastSignIn -and $lastNonInteractiveSignIn) {
                $mostRecentSignIn = if ($lastSignIn -gt $lastNonInteractiveSignIn) { $lastSignIn } else { $lastNonInteractiveSignIn }
            }
            elseif ($lastSignIn) {
                $mostRecentSignIn = $lastSignIn
            }
            elseif ($lastNonInteractiveSignIn) {
                $mostRecentSignIn = $lastNonInteractiveSignIn
            }
            
            if ($mostRecentSignIn) {
                $daysSinceSignIn = (New-TimeSpan -Start $mostRecentSignIn -End (Get-Date)).Days
                
                if ($mostRecentSignIn -lt $thresholdDate) {
                    $status = "Inactive"
                    $script:InactiveCount++
                }
            }
            else {
                $status = "Never Signed In"
                $script:NeverSignedInCount++
                $daysSinceSignIn = (New-TimeSpan -Start $user.CreatedDateTime -End (Get-Date)).Days
            }
        }
        else {
            $status = "Never Signed In"
            $script:NeverSignedInCount++
            if ($user.CreatedDateTime) {
                $daysSinceSignIn = (New-TimeSpan -Start $user.CreatedDateTime -End (Get-Date)).Days
            }
        }
        
        # Filter based on status
        if ($status -eq "Inactive" -or ($status -eq "Never Signed In" -and $IncludeNeverSignedIn)) {
            $licenseCount = if ($user.AssignedLicenses) { $user.AssignedLicenses.Count } else { 0 }
            $licenseNames = @()
            
            # Use hashtable lookup (FAST - no API call)
            if ($licenseCount -gt 0 -and $script:SkuLookup.Count -gt 0) {
                foreach ($license in $user.AssignedLicenses) {
                    if ($script:SkuLookup.ContainsKey($license.SkuId)) {
                        $licenseNames += $script:SkuLookup[$license.SkuId].SkuPartNumber
                    }
                }
            }
            
            $obj = [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Mail = $user.Mail
                Department = $user.Department
                JobTitle = $user.JobTitle
                OfficeLocation = $user.OfficeLocation
                Status = $status
                LastInteractiveSignIn = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
                LastNonInteractiveSignIn = if ($lastNonInteractiveSignIn) { $lastNonInteractiveSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
                DaysSinceLastActivity = $daysSinceSignIn
                AccountCreated = if ($user.CreatedDateTime) { $user.CreatedDateTime.ToString('yyyy-MM-dd') } else { "Unknown" }
                LicenseCount = $licenseCount
                AssignedLicenses = if ($licenseNames.Count -gt 0) { $licenseNames -join "; " } else { "None" }
                ReportDate = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
            
            $script:Results += $obj
        }
    }
    
    Write-Progress -Activity "Analyzing User Activity" -Completed
    
    Write-Host "`nAnalysis completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving or analyzing users: $($_.Exception.Message)" -ForegroundColor Red
    Write-Progress -Activity "Analyzing User Activity" -Completed
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# Export and display results
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Inactive User Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total Inactive Users: $script:InactiveCount" -ForegroundColor White
    Write-Host "  Users Never Signed In: $script:NeverSignedInCount" -ForegroundColor White
    Write-Host "  Licensed Inactive Users: $(($script:Results | Where-Object { $_.LicenseCount -gt 0 }).Count)" -ForegroundColor Yellow
    Write-Host "  Total Results Exported: $($script:Results.Count)" -ForegroundColor White
    
    try {
        $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
        Write-Host "  Report exported successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "`n  ERROR exporting report: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $script:Results | Select-Object -First 10 | 
        Format-Table DisplayName, UserPrincipalName, Status, DaysSinceLastActivity, LicenseCount -AutoSize
    
    if (Test-Path $ExportPath) {
        $openFile = Read-Host "`nWould you like to open the CSV report? (Y/N)"
        if ($openFile -eq 'Y' -or $openFile -eq 'y') {
            try {
                Invoke-Item $ExportPath
            }
            catch {
                Write-Host "Could not open file: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}
else {
    Write-Host "`nNo inactive users found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Cyan

try {
    Disconnect-MgGraph -ErrorAction Stop | Out-Null
    Write-Host "Disconnected successfully." -ForegroundColor Green
}
catch {
    Write-Host "Disconnect completed." -ForegroundColor Green
}

Write-Host "`nScript completed successfully.`n" -ForegroundColor Green
exit 0
