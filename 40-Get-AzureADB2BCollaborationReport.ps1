<#
====================================================================================
Script Name: 40-Get-AzureADB2BCollaborationReport.ps1
Description: Azure AD B2B external collaboration settings and guest governance
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
    [switch]$IncludeGuestUsers,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowCrossTenantAccess,
    
    [Parameter(Mandatory=$false)]
    [switch]$DetailedOutput,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\AzureAD_B2B_Collaboration_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Azure AD B2B Collaboration Settings Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph.Identity.SignIns", "Microsoft.Graph.Users")

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
    Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All", "User.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Initialize results
$script:Results = @()
$guestUserCount = 0

# Retrieve B2B authorization policy
Write-Host "Retrieving Azure AD B2B collaboration policies..." -ForegroundColor Cyan

try {
    $authPolicy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
    
    if ($authPolicy) {
        Write-Host "Retrieved B2B authorization policy.`n" -ForegroundColor Green
        
        # Parse guest user role permissions
        $guestUserRole = $authPolicy.GuestUserRoleId
        $guestUserRoleName = switch ($guestUserRole) {
            "10dae51f-b6af-4016-8d66-8c2a99b929b3" { "Guest User Access (Most Restrictive)" }
            "2af84b1e-32c8-42b7-82bc-daa82404023b" { "Restricted Guest User Access" }
            "a0b1b346-4d3e-4e8b-98f8-753987be4970" { "Guest User Access (Same as Member Users)" }
            default { "Custom Role" }
        }
        
        # Parse external collaboration settings
        $allowInvites = $authPolicy.AllowInvitesFrom
        $allowedToInviteGuests = $authPolicy.AllowedToSignUpEmailBasedSubscriptions
        
        # Create policy summary object
        $obj = [PSCustomObject]@{
            Category = "B2B Authorization Policy"
            Setting = "Guest User Role"
            Value = $guestUserRoleName
            Description = "Defines what guest users can see and do"
            Enabled = "N/A"
            Details = $guestUserRole
        }
        $script:Results += $obj
        
        $obj = [PSCustomObject]@{
            Category = "B2B Authorization Policy"
            Setting = "Who Can Invite Guests"
            Value = $allowInvites
            Description = "Controls who can invite external users"
            Enabled = "N/A"
            Details = "AllowInvitesFrom property"
        }
        $script:Results += $obj
        
        $obj = [PSCustomObject]@{
            Category = "B2B Authorization Policy"
            Setting = "Email Subscriptions"
            Value = $allowedToSignUpEmailBasedSubscriptions
            Description = "Allow email-based subscriptions"
            Enabled = "N/A"
            Details = "Email subscription setting"
        }
        $script:Results += $obj
    }
}
catch {
    Write-Host "Error retrieving B2B authorization policy: $_" -ForegroundColor Red
}

# Retrieve external identities policy
try {
    Write-Host "Retrieving external identities policy..." -ForegroundColor Cyan
    
    $externalPolicy = Get-MgPolicyExternalIdentitiesPolicy -ErrorAction SilentlyContinue
    
    if ($externalPolicy) {
        $obj = [PSCustomObject]@{
            Category = "External Identities Policy"
            Setting = "Allow External Identities"
            Value = $externalPolicy.AllowExternalIdentitiesToLeave
            Description = "Whether external identities can leave the organization"
            Enabled = $externalPolicy.AllowExternalIdentitiesToLeave
            Details = "External identities configuration"
        }
        $script:Results += $obj
    }
}
catch {
    Write-Warning "Could not retrieve external identities policy (may not be configured)"
}

# Retrieve cross-tenant access settings
if ($ShowCrossTenantAccess) {
    Write-Host "Retrieving cross-tenant access settings..." -ForegroundColor Cyan
    
    try {
        $crossTenantAccess = Get-MgPolicyCrossTenantAccessPolicyDefault -ErrorAction SilentlyContinue
        
        if ($crossTenantAccess) {
            # Inbound settings
            $inboundAllowed = $crossTenantAccess.InboundTrust.IsCompliantDeviceAccepted
            $obj = [PSCustomObject]@{
                Category = "Cross-Tenant Access"
                Setting = "Inbound - Accept Compliant Devices"
                Value = $inboundAllowed
                Description = "Accept compliant devices from other tenants"
                Enabled = $inboundAllowed
                Details = "Inbound trust configuration"
            }
            $script:Results += $obj
            
            # Outbound settings
            $outboundAllowed = $crossTenantAccess.B2BCollaborationOutbound.Applications.AccessType
            $obj = [PSCustomObject]@{
                Category = "Cross-Tenant Access"
                Setting = "Outbound Collaboration"
                Value = $outboundAllowed
                Description = "Outbound B2B collaboration access type"
                Enabled = "N/A"
                Details = "Outbound B2B settings"
            }
            $script:Results += $obj
        }
    }
    catch {
        Write-Warning "Could not retrieve cross-tenant access settings"
    }
}

# Get guest users if requested
if ($IncludeGuestUsers) {
    Write-Host "Retrieving guest user accounts..." -ForegroundColor Cyan
    
    try {
        $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -ErrorAction Stop
        $guestUserCount = $guestUsers.Count
        
        Write-Host "Found $guestUserCount guest user(s).`n" -ForegroundColor Green
        
        if ($DetailedOutput) {
            $progressCounter = 0
            
            foreach ($guest in $guestUsers) {
                $progressCounter++
                Write-Progress -Activity "Processing Guest Users" -Status "Guest $progressCounter of $guestUserCount" -PercentComplete (($progressCounter / $guestUserCount) * 100)
                
                # Get last sign-in
                $lastSignIn = "Never"
                if ($guest.SignInActivity) {
                    $lastSignIn = $guest.SignInActivity.LastSignInDateTime
                }
                
                $obj = [PSCustomObject]@{
                    Category = "Guest User"
                    Setting = $guest.DisplayName
                    Value = $guest.UserPrincipalName
                    Description = "External guest user account"
                    Enabled = $guest.AccountEnabled
                    Details = "Last Sign-In: $lastSignIn"
                    CreatedDateTime = $guest.CreatedDateTime
                    ExternalUserState = $guest.ExternalUserState
                }
                $script:Results += $obj
            }
            
            Write-Progress -Activity "Processing Guest Users" -Completed
        }
        else {
            $obj = [PSCustomObject]@{
                Category = "Guest Users Summary"
                Setting = "Total Guest Users"
                Value = $guestUserCount
                Description = "Total number of guest users in directory"
                Enabled = "N/A"
                Details = "Use -DetailedOutput to see individual guests"
            }
            $script:Results += $obj
        }
    }
    catch {
        Write-Host "Error retrieving guest users: $_" -ForegroundColor Red
    }
}

# Export and display results
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "B2B Collaboration Summary:" -ForegroundColor Green
    Write-Host "  Total Settings Retrieved: $($script:Results.Count)" -ForegroundColor White
    
    if ($IncludeGuestUsers) {
        Write-Host "  Total Guest Users: $guestUserCount" -ForegroundColor Yellow
        Write-Host "  Active Guest Users: $(($guestUsers | Where-Object { $_.AccountEnabled -eq $true }).Count)" -ForegroundColor Green
    }
    
    # Export to CSV
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY NOTE:" -ForegroundColor Red
    Write-Host "Review B2B collaboration settings regularly for security compliance.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $script:Results | Select-Object -First 10 | Format-Table Category, Setting, Value, Enabled -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No B2B collaboration settings found." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
