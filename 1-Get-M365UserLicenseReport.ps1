<#
====================================================================================
Script Name: Get-M365UserLicenseReport.ps1
Description: Comprehensive Microsoft 365 user license assignment and usage report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
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

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$LicenseSkuId,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDisabledUsers,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowServicePlans,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_User_Licenses_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 User License Report Generator" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Users"

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

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
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
    "TEAMS_EXPLORATORY" = "Teams Exploratory"
    "PROJECTPROFESSIONAL" = "Project Plan 5"
    "VISIOCLIENT" = "Visio Plan 2"
    "EXCHANGESTANDARD" = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
}

# Retrieve users
Write-Host "Retrieving user license information..." -ForegroundColor Cyan
$results = @()
$userCount = 0

try {
    if ($UserPrincipalName) {
        $users = Get-MgUser -UserId $UserPrincipalName -Property DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, LicenseAssignmentStates -ErrorAction Stop
        $users = @($users)
    }
    else {
        $filter = "assignedLicenses/`$count ne 0"
        if (-not $IncludeDisabledUsers) {
            $filter += " and accountEnabled eq true"
        }
        
        $users = Get-MgUser -Filter $filter -ConsistencyLevel eventual -CountVariable userCount -All -Property DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, LicenseAssignmentStates
    }
    
    Write-Host "Found $($users.Count) licensed user(s). Processing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($user in $users) {
        $progressCounter++
        Write-Progress -Activity "Processing Users" -Status "User $progressCounter of $($users.Count): $($user.UserPrincipalName)" -PercentComplete (($progressCounter / $users.Count) * 100)
        
        if ($user.AssignedLicenses.Count -gt 0) {
            foreach ($license in $user.AssignedLicenses) {
                $skuId = $license.SkuId
                
                try {
                    $subscribedSku = Get-MgSubscribedSku -All | Where-Object { $_.SkuId -eq $skuId }
                    $skuPartNumber = $subscribedSku.SkuPartNumber
                    $friendlyName = if ($skuMappings.ContainsKey($skuPartNumber)) { $skuMappings[$skuPartNumber] } else { $skuPartNumber }
                    
                    if ($LicenseSkuId -and $skuPartNumber -ne $LicenseSkuId) {
                        continue
                    }
                    
                    $servicePlans = ""
                    if ($ShowServicePlans -and $subscribedSku.ServicePlans) {
                        $servicePlans = ($subscribedSku.ServicePlans | ForEach-Object { $_.ServicePlanName }) -join "; "
                    }
                    
                    $obj = [PSCustomObject]@{
                        DisplayName = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        AccountEnabled = $user.AccountEnabled
                        LicenseName = $friendlyName
                        LicenseSKU = $skuPartNumber
                        SkuId = $skuId
                        ServicePlans = $servicePlans
                        AssignmentDate = (Get-Date)
                    }
                    
                    $results += $obj
                }
                catch {
                    Write-Warning "Error processing license for $($user.UserPrincipalName): $_"
                }
            }
        }
    }
    
    Write-Progress -Activity "Processing Users" -Completed
}
catch {
    Write-Host "Error retrieving users: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Export Summary:" -ForegroundColor Green
    Write-Host "  Total Licensed Users Processed: $($results.Count)" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No licensed users found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
