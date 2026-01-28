<#
====================================================================================
Script Name: 15-Get-M365ConditionalAccessPoliciesReport.ps1
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
    [ValidateSet("enabled","disabled","enabledForReportingButNotEnforced","All")]
    [string]$State = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$PolicyName,
    
    [Parameter(Mandatory=$false)]
    [switch]$EnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Conditional_Access_Policies_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Conditional Access Policies Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Identity.SignIns"

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
    Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve Conditional Access policies
Write-Host "Retrieving Conditional Access policies..." -ForegroundColor Cyan
$results = @()
$enabledCount = 0
$disabledCount = 0
$reportOnlyCount = 0

try {
    $policies = Get-MgIdentityConditionalAccessPolicy -All
    
    if ($PolicyName) {
        $policies = $policies | Where-Object { $_.DisplayName -like "*$PolicyName*" }
    }
    
    Write-Host "Found $($policies.Count) Conditional Access policy/policies. Analyzing configurations...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($policy in $policies) {
        $progressCounter++
        Write-Progress -Activity "Processing Conditional Access Policies" -Status "Policy $progressCounter of $($policies.Count): $($policy.DisplayName)" -PercentComplete (($progressCounter / $policies.Count) * 100)
        
        # Filter by state if specified
        if ($EnabledOnly -and $policy.State -ne "enabled") {
            continue
        }
        
        if ($State -ne "All" -and $policy.State -ne $State) {
            continue
        }
        
        # Count by state
        switch ($policy.State) {
            "enabled" { $enabledCount++ }
            "disabled" { $disabledCount++ }
            "enabledForReportingButNotEnforced" { $reportOnlyCount++ }
        }
        
        # Parse conditions
        $includedUsers = if ($policy.Conditions.Users.IncludeUsers) { $policy.Conditions.Users.IncludeUsers -join "; " } else { "None" }
        $excludedUsers = if ($policy.Conditions.Users.ExcludeUsers) { $policy.Conditions.Users.ExcludeUsers -join "; " } else { "None" }
        $includedGroups = if ($policy.Conditions.Users.IncludeGroups) { $policy.Conditions.Users.IncludeGroups -join "; " } else { "None" }
        $excludedGroups = if ($policy.Conditions.Users.ExcludeGroups) { $policy.Conditions.Users.ExcludeGroups -join "; " } else { "None" }
        $includedRoles = if ($policy.Conditions.Users.IncludeRoles) { $policy.Conditions.Users.IncludeRoles -join "; " } else { "None" }
        $excludedRoles = if ($policy.Conditions.Users.ExcludeRoles) { $policy.Conditions.Users.ExcludeRoles -join "; " } else { "None" }
        
        # Applications
        $includedApps = if ($policy.Conditions.Applications.IncludeApplications) { $policy.Conditions.Applications.IncludeApplications -join "; " } else { "None" }
        $excludedApps = if ($policy.Conditions.Applications.ExcludeApplications) { $policy.Conditions.Applications.ExcludeApplications -join "; " } else { "None" }
        
        # Platforms
        $includedPlatforms = if ($policy.Conditions.Platforms.IncludePlatforms) { $policy.Conditions.Platforms.IncludePlatforms -join "; " } else { "All" }
        $excludedPlatforms = if ($policy.Conditions.Platforms.ExcludePlatforms) { $policy.Conditions.Platforms.ExcludePlatforms -join "; " } else { "None" }
        
        # Locations
        $includedLocations = if ($policy.Conditions.Locations.IncludeLocations) { $policy.Conditions.Locations.IncludeLocations -join "; " } else { "All" }
        $excludedLocations = if ($policy.Conditions.Locations.ExcludeLocations) { $policy.Conditions.Locations.ExcludeLocations -join "; " } else { "None" }
        
        # Client app types
        $clientAppTypes = if ($policy.Conditions.ClientAppTypes) { $policy.Conditions.ClientAppTypes -join "; " } else { "All" }
        
        # Sign-in risk levels
        $signInRiskLevels = if ($policy.Conditions.SignInRiskLevels) { $policy.Conditions.SignInRiskLevels -join "; " } else { "Not Configured" }
        $userRiskLevels = if ($policy.Conditions.UserRiskLevels) { $policy.Conditions.UserRiskLevels -join "; " } else { "Not Configured" }
        
        # Grant controls
        $grantControls = "None"
        if ($policy.GrantControls) {
            $controls = @()
            if ($policy.GrantControls.BuiltInControls) {
                $controls += $policy.GrantControls.BuiltInControls
            }
            if ($policy.GrantControls.CustomAuthenticationFactors) {
                $controls += $policy.GrantControls.CustomAuthenticationFactors
            }
            $grantControls = $controls -join "; "
            $grantOperator = $policy.GrantControls.Operator
        }
        
        # Session controls
        $sessionControls = @()
        if ($policy.SessionControls) {
            if ($policy.SessionControls.ApplicationEnforcedRestrictions) {
                $sessionControls += "App Enforced Restrictions"
            }
            if ($policy.SessionControls.CloudAppSecurity) {
                $sessionControls += "Cloud App Security"
            }
            if ($policy.SessionControls.PersistentBrowser) {
                $sessionControls += "Persistent Browser"
            }
            if ($policy.SessionControls.SignInFrequency) {
                $sessionControls += "Sign-in Frequency"
            }
        }
        $sessionControlsStr = if ($sessionControls.Count -gt 0) { $sessionControls -join "; " } else { "None" }
        
        $obj = [PSCustomObject]@{
            PolicyName = $policy.DisplayName
            State = $policy.State
            CreatedDateTime = $policy.CreatedDateTime
            ModifiedDateTime = $policy.ModifiedDateTime
            IncludedUsers = $includedUsers
            ExcludedUsers = $excludedUsers
            IncludedGroups = $includedGroups
            ExcludedGroups = $excludedGroups
            IncludedRoles = $includedRoles
            ExcludedRoles = $excludedRoles
            IncludedApplications = $includedApps
            ExcludedApplications = $excludedApps
            IncludedPlatforms = $includedPlatforms
            ExcludedPlatforms = $excludedPlatforms
            IncludedLocations = $includedLocations
            ExcludedLocations = $excludedLocations
            ClientAppTypes = $clientAppTypes
            SignInRiskLevels = $signInRiskLevels
            UserRiskLevels = $userRiskLevels
            GrantControls = $grantControls
            GrantOperator = if ($policy.GrantControls) { $policy.GrantControls.Operator } else { "N/A" }
            SessionControls = $sessionControlsStr
            PolicyId = $policy.Id
        }
        
        $results += $obj
    }
    
    Write-Progress -Activity "Processing Conditional Access Policies" -Completed
}
catch {
    Write-Host "Error retrieving Conditional Access policies: $_" -ForegroundColor Red
    Write-Host "Note: This feature requires Azure AD Premium P1 or P2 licensing.`n" -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Conditional Access Policy Summary:" -ForegroundColor Green
    Write-Host "  Total Policies: $($results.Count)" -ForegroundColor White
    Write-Host "  Enabled Policies: $enabledCount" -ForegroundColor Green
    Write-Host "  Report-Only Policies: $reportOnlyCount" -ForegroundColor Yellow
    Write-Host "  Disabled Policies: $disabledCount" -ForegroundColor Red
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "RECOMMENDATION:" -ForegroundColor Cyan
    Write-Host "Review policy assignments to ensure proper coverage and avoid conflicts.`n" -ForegroundColor White
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table PolicyName, State, GrantControls, CreatedDateTime -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No Conditional Access policies found matching the specified criteria." -ForegroundColor Yellow
    Write-Host "Note: Requires Azure AD Premium P1 or P2 licensing.`n" -ForegroundColor Cyan
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
