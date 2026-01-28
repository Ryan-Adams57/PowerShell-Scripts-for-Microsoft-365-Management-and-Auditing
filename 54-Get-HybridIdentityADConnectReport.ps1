<#
====================================================================================
Script Name: 54-Get-HybridIdentityADConnectReport.ps1
Description: Azure AD Connect sync status and hybrid identity configuration report
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
    [switch]$ErrorsOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSyncDetails,
    
    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 7,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Hybrid_Identity_ADConnect_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Hybrid Identity & Azure AD Connect Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Identity.DirectoryManagement"

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

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "Directory.Read.All", "Organization.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

$script:Results = @()
$syncErrorCount = 0
$syncSuccessCount = 0

Write-Host "Retrieving Azure AD Connect sync status..." -ForegroundColor Cyan

try {
    # Get organization details for directory sync info
    $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    
    if ($org) {
        $directorySyncEnabled = $org.OnPremisesSyncEnabled
        $lastDirSyncTime = $org.OnPremisesLastSyncDateTime
        
        Write-Host "Directory Sync Enabled: $directorySyncEnabled" -ForegroundColor $(if ($directorySyncEnabled) { "Green" } else { "Yellow" })
        
        if ($directorySyncEnabled) {
            Write-Host "Last Sync Time: $lastDirSyncTime`n" -ForegroundColor Green
            
            # Get on-premises synced objects
            Write-Host "Retrieving synced objects..." -ForegroundColor Cyan
            
            $syncedUsers = Get-MgUser -Filter "onPremisesSyncEnabled eq true" -All -ErrorAction SilentlyContinue
            $syncedGroups = Get-MgGroup -Filter "onPremisesSyncEnabled eq true" -All -ErrorAction SilentlyContinue
            
            Write-Host "Found $($syncedUsers.Count) synced users" -ForegroundColor White
            Write-Host "Found $($syncedGroups.Count) synced groups`n" -ForegroundColor White
            
            # Process synced users
            $progressCounter = 0
            $startDate = (Get-Date).AddDays(-$DaysBack)
            
            foreach ($user in $syncedUsers) {
                $progressCounter++
                
                if ($progressCounter % 100 -eq 0) {
                    Write-Progress -Activity "Processing Synced Users" -Status "User $progressCounter of $($syncedUsers.Count)" -PercentComplete (($progressCounter / $syncedUsers.Count) * 100)
                }
                
                $hasSyncError = $false
                $syncErrors = ""
                
                # Check for sync errors
                if ($user.OnPremisesProvisioningErrors -and $user.OnPremisesProvisioningErrors.Count -gt 0) {
                    $hasSyncError = $true
                    $syncErrorCount++
                    $syncErrors = ($user.OnPremisesProvisioningErrors | ForEach-Object { $_.ErrorDetail }) -join "; "
                }
                else {
                    $syncSuccessCount++
                }
                
                if ($ErrorsOnly -and -not $hasSyncError) {
                    continue
                }
                
                $lastSyncTime = $user.OnPremisesLastSyncDateTime
                
                if ($lastSyncTime -and $lastSyncTime -lt $startDate) {
                    continue
                }
                
                $obj = [PSCustomObject]@{
                    ObjectType = "User"
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                    LastSyncDateTime = $lastSyncTime
                    HasSyncError = $hasSyncError
                    SyncErrors = $syncErrors
                    OnPremisesDomainName = $user.OnPremisesDomainName
                    OnPremisesSamAccountName = $user.OnPremisesSamAccountName
                    OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
                    ImmutableId = $user.OnPremisesImmutableId
                }
                
                $script:Results += $obj
            }
            
            Write-Progress -Activity "Processing Synced Users" -Completed
            
            # Process synced groups if requested
            if ($IncludeSyncDetails) {
                Write-Host "Processing synced groups..." -ForegroundColor Cyan
                
                foreach ($group in $syncedGroups) {
                    $hasSyncError = $false
                    $syncErrors = ""
                    
                    if ($group.OnPremisesProvisioningErrors -and $group.OnPremisesProvisioningErrors.Count -gt 0) {
                        $hasSyncError = $true
                        $syncErrorCount++
                        $syncErrors = ($group.OnPremisesProvisioningErrors | ForEach-Object { $_.ErrorDetail }) -join "; "
                    }
                    
                    if ($ErrorsOnly -and -not $hasSyncError) {
                        continue
                    }
                    
                    $obj = [PSCustomObject]@{
                        ObjectType = "Group"
                        DisplayName = $group.DisplayName
                        UserPrincipalName = "N/A"
                        OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                        LastSyncDateTime = $group.OnPremisesLastSyncDateTime
                        HasSyncError = $hasSyncError
                        SyncErrors = $syncErrors
                        OnPremisesDomainName = $group.OnPremisesDomainName
                        OnPremisesSamAccountName = $group.OnPremisesSamAccountName
                        OnPremisesDistinguishedName = $group.OnPremisesDistinguishedName
                        ImmutableId = $group.OnPremisesSecurityIdentifier
                    }
                    
                    $script:Results += $obj
                }
            }
        }
        else {
            Write-Host "Directory synchronization is not enabled for this tenant.`n" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host "Error retrieving AD Connect data: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Hybrid Identity Summary:" -ForegroundColor Green
    Write-Host "  Directory Sync Enabled: $directorySyncEnabled" -ForegroundColor White
    Write-Host "  Last Sync Time: $lastDirSyncTime" -ForegroundColor White
    Write-Host "  Total Synced Objects: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Objects with Sync Errors: $syncErrorCount" -ForegroundColor Red
    Write-Host "  Successfully Synced: $syncSuccessCount" -ForegroundColor Green
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($syncErrorCount -gt 0) {
        Write-Host "WARNING: Sync errors detected. Review immediately.`n" -ForegroundColor Red
    }
    
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $script:Results | Select-Object -First 10 | Format-Table ObjectType, DisplayName, HasSyncError, LastSyncDateTime -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No synced objects found." -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
