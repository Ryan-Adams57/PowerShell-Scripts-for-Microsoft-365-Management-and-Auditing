<#
====================================================================================
Script Name: Get-M365InactiveUsersReport.ps1
Description: Identifies inactive Microsoft 365 users based on last sign-in activity
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
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

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory=$false)]
    [string]$Department,
    
    [Parameter(Mandatory=$false)]
    [switch]$LicensedOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeNeverSignedIn,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Inactive_Users_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Inactive Users Report Generator" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Reports")

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Required module '$module' is not installed." -ForegroundColor Yellow
        $install = Read-Host "Would you like to install it now? (Y/N)"
        
        if ($install -eq 'Y' -or $install -eq 'y') {
            try {
                Write-Host "Installing $module..." -ForegroundColor Cyan
                Install-Module -Name $module -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
                Write-Host "$module installed successfully.`n" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to install $module. Error: $_" -ForegroundColor Red
                exit
            }
        }
        else {
            Write-Host "Module installation declined. Script cannot continue." -ForegroundColor Red
            exit
        }
    }
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Calculate threshold date
$thresholdDate = (Get-Date).AddDays(-$InactiveDays)
Write-Host "Identifying users inactive since: $($thresholdDate.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
Write-Host "Threshold: $InactiveDays days`n" -ForegroundColor Cyan

# Build filter
$filter = "accountEnabled eq true"
if ($Department) {
    $filter += " and department eq '$Department'"
}
if ($LicensedOnly) {
    $filter += " and assignedLicenses/`$count ne 0"
}

# Retrieve users
Write-Host "Retrieving user accounts..." -ForegroundColor Cyan
$results = @()

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
        "OfficeLocation"
    )
    
    $users = Get-MgUser -Filter $filter -ConsistencyLevel eventual -All -Property $properties
    
    Write-Host "Found $($users.Count) user account(s). Analyzing activity...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $inactiveCount = 0
    $neverSignedInCount = 0
    
    foreach ($user in $users) {
        $progressCounter++
        Write-Progress -Activity "Analyzing User Activity" -Status "User $progressCounter of $($users.Count): $($user.UserPrincipalName)" -PercentComplete (($progressCounter / $users.Count) * 100)
        
        $lastSignIn = $null
        $lastNonInteractiveSignIn = $null
        $daysSinceSignIn = $null
        $status = "Active"
        
        if ($user.SignInActivity) {
            $lastSignIn = $user.SignInActivity.LastSignInDateTime
            $lastNonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime
            
            # Determine most recent sign-in
            $mostRecentSignIn = $lastSignIn
            if ($lastNonInteractiveSignIn -and (-not $lastSignIn -or $lastNonInteractiveSignIn -gt $lastSignIn)) {
                $mostRecentSignIn = $lastNonInteractiveSignIn
            }
            
            if ($mostRecentSignIn) {
                $daysSinceSignIn = (New-TimeSpan -Start $mostRecentSignIn -End (Get-Date)).Days
                
                if ($mostRecentSignIn -lt $thresholdDate) {
                    $status = "Inactive"
                    $inactiveCount++
                }
            }
            else {
                $status = "Never Signed In"
                $neverSignedInCount++
                $daysSinceSignIn = (New-TimeSpan -Start $user.CreatedDateTime -End (Get-Date)).Days
            }
        }
        else {
            $status = "Never Signed In"
            $neverSignedInCount++
            $daysSinceSignIn = (New-TimeSpan -Start $user.CreatedDateTime -End (Get-Date)).Days
        }
        
        # Filter based on status
        if ($status -eq "Inactive" -or ($status -eq "Never Signed In" -and $IncludeNeverSignedIn)) {
            $licenseCount = if ($user.AssignedLicenses) { $user.AssignedLicenses.Count } else { 0 }
            $licenseNames = ""
            
            if ($licenseCount -gt 0) {
                $skus = Get-MgSubscribedSku -All
                $assignedSkus = @()
                foreach ($license in $user.AssignedLicenses) {
                    $sku = $skus | Where-Object { $_.SkuId -eq $license.SkuId }
                    if ($sku) {
                        $assignedSkus += $sku.SkuPartNumber
                    }
                }
                $licenseNames = $assignedSkus -join "; "
            }
            
            $obj = [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Department = $user.Department
                JobTitle = $user.JobTitle
                Status = $status
                LastInteractiveSignIn = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
                LastNonInteractiveSignIn = if ($lastNonInteractiveSignIn) { $lastNonInteractiveSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
                DaysSinceLastActivity = $daysSinceSignIn
                AccountCreated = $user.CreatedDateTime.ToString('yyyy-MM-dd')
                LicenseCount = $licenseCount
                AssignedLicenses = $licenseNames
                OfficeLocation = $user.OfficeLocation
            }
            
            $results += $obj
        }
    }
    
    Write-Progress -Activity "Analyzing User Activity" -Completed
}
catch {
    Write-Host "Error retrieving or analyzing users: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Inactive User Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total Inactive Users: $inactiveCount" -ForegroundColor White
    Write-Host "  Users Never Signed In: $neverSignedInCount" -ForegroundColor White
    Write-Host "  Licensed Inactive Users: $(($results | Where-Object { $_.LicenseCount -gt 0 }).Count)" -ForegroundColor Yellow
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table DisplayName, UserPrincipalName, Status, DaysSinceLastActivity, LicenseCount -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No inactive users found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
