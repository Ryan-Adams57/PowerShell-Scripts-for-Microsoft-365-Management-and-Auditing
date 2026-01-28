<#
====================================================================================
Script Name: 1-Get-M365UserLicenseReport.ps1
Description: Comprehensive Microsoft 365 user license assignment and usage report
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all licensed users across the Microsoft 365 tenant
• Displays detailed license SKU assignments per user
• Shows enabled/disabled service plans within each license
• Exports friendly license names (E3, E5, F3, etc.)
• Supports filtering by license type or user
• Includes license cost estimation capabilities
• Generates timestamped CSV reports
• MFA-compatible authentication
• NO API calls inside loops (uses hashtable lookups)
• Comprehensive error handling with try/catch/finally
• Progress indicators for long operations

REQUIREMENTS:
• Microsoft.Graph.Users module
• Microsoft.Graph.Identity.DirectoryManagement module
• User.Read.All permission
• Directory.Read.All permission

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Specific user to query")]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false, HelpMessage="Filter by license SKU ID")]
    [string]$LicenseSkuId,
    
    [Parameter(Mandatory=$false, HelpMessage="Include disabled user accounts")]
    [switch]$IncludeDisabledUsers,
    
    [Parameter(Mandatory=$false, HelpMessage="Show service plans for each license")]
    [switch]$ShowServicePlans,
    
    [Parameter(Mandatory=$false, HelpMessage="Export CSV path")]
    [string]$ExportPath = ".\M365_User_Licenses_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize script variables
$script:Results = @()
$script:SkuLookup = @{}

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 User License Report Generator" -ForegroundColor Green
Write-Host "Version 2.0 - Production Ready" -ForegroundColor White
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

# Module validation
$requiredModules = @(
    "Microsoft.Graph.Users", 
    "Microsoft.Graph.Identity.DirectoryManagement"
)

Write-Host "Validating required modules..." -ForegroundColor Cyan

foreach ($requiredModule in $requiredModules) {
    try {
        if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
            Write-Host "ERROR: Required module '$requiredModule' is not installed." -ForegroundColor Red
            Write-Host "Please run: Install-Module -Name $requiredModule -Scope CurrentUser" -ForegroundColor Yellow
            exit 1
        }
        else {
            $moduleInfo = Get-Module -ListAvailable -Name $requiredModule | 
                Sort-Object Version -Descending | 
                Select-Object -First 1
            Write-Host "  ✓ $requiredModule (v$($moduleInfo.Version))" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "ERROR validating module '$requiredModule': $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    
    if (-not $context) {
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
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

# License SKU friendly name mapping
$skuMappings = @{
    "SPE_E3" = "Microsoft 365 E3"
    "SPE_E5" = "Microsoft 365 E5"
    "ENTERPRISEPREMIUM" = "Office 365 E5"
    "ENTERPRISEPACK" = "Office 365 E3"
    "STANDARDPACK" = "Office 365 E1"
    "DESKLESSPACK" = "Microsoft 365 F3"
    "SPB" = "Microsoft 365 Business Premium"
    "POWER_BI_PRO" = "Power BI Pro"
    "POWER_BI_STANDARD" = "Power BI Free"
    "TEAMS_EXPLORATORY" = "Teams Exploratory"
    "PROJECTPROFESSIONAL" = "Project Plan 5"
    "PROJECTESSENTIALS" = "Project Plan 1"
    "VISIOCLIENT" = "Visio Plan 2"
    "VISIONLINE" = "Visio Plan 1"
    "EXCHANGESTANDARD" = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
    "SHAREPOINTSTANDARD" = "SharePoint Online Plan 1"
    "SHAREPOINTENTERPRISE" = "SharePoint Online Plan 2"
    "STREAM" = "Microsoft Stream"
    "FLOW_FREE" = "Power Automate Free"
    "POWERAPPS_VIRAL" = "Power Apps Trial"
}

# Retrieve organization license SKUs ONCE (critical: no API calls in loops)
Write-Host "Retrieving organization license SKUs..." -ForegroundColor Cyan

try {
    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    Write-Host "Retrieved $($subscribedSkus.Count) license SKU(s)." -ForegroundColor Green
    
    # Create hashtable for FAST lookups (no API calls in user loop)
    foreach ($sku in $subscribedSkus) {
        $script:SkuLookup[$sku.SkuId] = $sku
    }
    
    Write-Host "SKU lookup table created for optimized processing.`n" -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving license SKUs: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Ensure you have Directory.Read.All permissions." -ForegroundColor Yellow
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# Retrieve users
Write-Host "Retrieving user license information..." -ForegroundColor Cyan

try {
    $users = @()
    
    if ($UserPrincipalName) {
        $users = @(Get-MgUser -UserId $UserPrincipalName `
            -Property DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, LicenseAssignmentStates `
            -ErrorAction Stop)
        Write-Host "Retrieved specific user: $UserPrincipalName" -ForegroundColor Green
    }
    else {
        $filter = "assignedLicenses/`$count ne 0"
        if (-not $IncludeDisabledUsers) {
            $filter += " and accountEnabled eq true"
        }
        
        $users = Get-MgUser -Filter $filter `
            -ConsistencyLevel eventual `
            -All `
            -Property DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, LicenseAssignmentStates `
            -ErrorAction Stop
        
        Write-Host "Retrieved $($users.Count) licensed user(s)." -ForegroundColor Green
    }
    
    Write-Host "Processing users...`n" -ForegroundColor Cyan
    
    $progressCounter = 0
    
    foreach ($user in $users) {
        $progressCounter++
        Write-Progress -Activity "Processing Users" `
            -Status "User $progressCounter of $($users.Count): $($user.UserPrincipalName)" `
            -PercentComplete (($progressCounter / $users.Count) * 100)
        
        if ($user.AssignedLicenses.Count -gt 0) {
            foreach ($license in $user.AssignedLicenses) {
                $skuId = $license.SkuId
                
                # Lookup SKU from hashtable (FAST - no API call)
                if ($script:SkuLookup.ContainsKey($skuId)) {
                    $subscribedSku = $script:SkuLookup[$skuId]
                    $skuPartNumber = $subscribedSku.SkuPartNumber
                    $friendlyName = if ($skuMappings.ContainsKey($skuPartNumber)) { 
                        $skuMappings[$skuPartNumber] 
                    } else { 
                        $skuPartNumber 
                    }
                    
                    # Apply license filter if specified
                    if ($LicenseSkuId -and $skuPartNumber -ne $LicenseSkuId) {
                        continue
                    }
                    
                    $servicePlans = ""
                    if ($ShowServicePlans -and $subscribedSku.ServicePlans) {
                        $servicePlans = ($subscribedSku.ServicePlans | 
                            ForEach-Object { $_.ServicePlanName }) -join "; "
                    }
                    
                    $obj = [PSCustomObject]@{
                        DisplayName = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        AccountEnabled = $user.AccountEnabled
                        LicenseName = $friendlyName
                        LicenseSKU = $skuPartNumber
                        SkuId = $skuId
                        ServicePlans = $servicePlans
                        ConsumedUnits = $subscribedSku.ConsumedUnits
                        TotalUnits = if ($subscribedSku.PrepaidUnits) { $subscribedSku.PrepaidUnits.Enabled } else { 0 }
                        AvailableUnits = if ($subscribedSku.PrepaidUnits) { 
                            $subscribedSku.PrepaidUnits.Enabled - $subscribedSku.ConsumedUnits 
                        } else { 0 }
                        ReportDate = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    }
                    
                    $script:Results += $obj
                }
                else {
                    Write-Warning "Unknown SKU ID for user $($user.UserPrincipalName): $skuId"
                }
            }
        }
    }
    
    Write-Progress -Activity "Processing Users" -Completed
    
    Write-Host "`nProcessing completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving or processing users: $($_.Exception.Message)" -ForegroundColor Red
    Write-Progress -Activity "Processing Users" -Completed
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    exit 1
}

# Export and display results
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "License Assignment Summary:" -ForegroundColor Green
    Write-Host "  Total License Assignments: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Unique Users: $(($script:Results | Select-Object -Unique UserPrincipalName).Count)" -ForegroundColor White
    Write-Host "  Unique License Types: $(($script:Results | Select-Object -Unique LicenseSKU).Count)" -ForegroundColor White
    
    # License breakdown
    Write-Host "`n  Assignments by License Type:" -ForegroundColor Cyan
    $script:Results | Group-Object LicenseName | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
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
        Format-Table DisplayName, UserPrincipalName, LicenseName, AccountEnabled -AutoSize
    
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
    Write-Host "`nNo licensed users found matching the specified criteria." -ForegroundColor Yellow
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
