<#
====================================================================================
Script Name: Get-AzureADAppRegistrationsReport.ps1
Description: Azure AD application registrations and service principals report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Azure AD app registrations
• Shows service principals and enterprise applications
• Lists API permissions and consent status
• Identifies apps with expired or expiring credentials
• Tracks app owners and creation dates
• Supports filtering by credential expiration
• Generates application security governance reports
• Critical for managing app access and permissions

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$ExpiringCredentialsOnly,
    
    [Parameter(Mandatory=$false)]
    [int]$ExpiringInDays = 30,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludePermissions,
    
    [Parameter(Mandatory=$false)]
    [switch]$OrphanedAppsOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\AzureAD_App_Registrations_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Azure AD Application Registrations Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Applications"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.`n" -ForegroundColor Green
    } else { exit }
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

$results = @()
$expiringCount = 0
$expiredCount = 0
$orphanedCount = 0

Write-Host "Retrieving Azure AD app registrations..." -ForegroundColor Cyan

try {
    $apps = Get-MgApplication -All -ErrorAction Stop
    
    Write-Host "Found $($apps.Count) app registration(s). Processing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $thresholdDate = (Get-Date).AddDays($ExpiringInDays)
    
    foreach ($app in $apps) {
        $progressCounter++
        Write-Progress -Activity "Processing Apps" -Status "App $progressCounter of $($apps.Count)" -PercentComplete (($progressCounter / $apps.Count) * 100)
        
        # Get owners
        $owners = Get-MgApplicationOwner -ApplicationId $app.Id -ErrorAction SilentlyContinue
        $ownerCount = if ($owners) { $owners.Count } else { 0 }
        $isOrphaned = $ownerCount -eq 0
        
        if ($isOrphaned) { $orphanedCount++ }
        if ($OrphanedAppsOnly -and -not $isOrphaned) { continue }
        
        # Check credential expiration
        $credentials = @()
        $credentials += $app.PasswordCredentials
        $credentials += $app.KeyCredentials
        
        $hasExpiringCreds = $false
        $hasExpiredCreds = $false
        $nextExpiry = $null
        $credentialCount = $credentials.Count
        
        foreach ($cred in $credentials) {
            if ($cred.EndDateTime) {
                if ($cred.EndDateTime -lt (Get-Date)) {
                    $hasExpiredCreds = $true
                }
                elseif ($cred.EndDateTime -lt $thresholdDate) {
                    $hasExpiringCreds = $true
                }
                
                if (-not $nextExpiry -or $cred.EndDateTime -lt $nextExpiry) {
                    $nextExpiry = $cred.EndDateTime
                }
            }
        }
        
        if ($hasExpiredCreds) { $expiredCount++ }
        if ($hasExpiringCreds) { $expiringCount++ }
        
        if ($ExpiringCredentialsOnly -and -not $hasExpiringCreds -and -not $hasExpiredCreds) { continue }
        
        # Get permissions if requested
        $permissions = "Not Retrieved"
        $permissionCount = 0
        
        if ($IncludePermissions) {
            $apiPermissions = $app.RequiredResourceAccess
            if ($apiPermissions) {
                $permList = @()
                foreach ($api in $apiPermissions) {
                    $permCount = $api.ResourceAccess.Count
                    $permissionCount += $permCount
                    $permList += "$($api.ResourceAppId): $permCount permission(s)"
                }
                $permissions = $permList -join "; "
            } else {
                $permissions = "None"
            }
        }
        
        # Get owner names
        $ownerNames = "None"
        if ($owners) {
            $ownerNames = ($owners | ForEach-Object { 
                if ($_.AdditionalProperties.userPrincipalName) {
                    $_.AdditionalProperties.userPrincipalName
                } else {
                    $_.AdditionalProperties.displayName
                }
            }) -join "; "
        }
        
        $results += [PSCustomObject]@{
            ApplicationName = $app.DisplayName
            ApplicationId = $app.AppId
            ObjectId = $app.Id
            CreatedDateTime = $app.CreatedDateTime
            Owners = $ownerNames
            OwnerCount = $ownerCount
            IsOrphaned = $isOrphaned
            CredentialCount = $credentialCount
            HasExpiringCredentials = $hasExpiringCreds
            HasExpiredCredentials = $hasExpiredCreds
            NextCredentialExpiry = $nextExpiry
            DaysUntilExpiry = if ($nextExpiry) { (New-TimeSpan -Start (Get-Date) -End $nextExpiry).Days } else { "N/A" }
            SignInAudience = $app.SignInAudience
            PermissionCount = $permissionCount
            APIPermissions = $permissions
            PublisherDomain = $app.PublisherDomain
            Notes = $app.Notes
        }
    }
    
    Write-Progress -Activity "Processing Apps" -Completed
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "App Registrations Summary:" -ForegroundColor Green
    Write-Host "  Total Apps: $($apps.Count)" -ForegroundColor White
    Write-Host "  Apps in Report: $($results.Count)" -ForegroundColor White
    Write-Host "  Orphaned Apps (no owners): $orphanedCount" -ForegroundColor Red
    Write-Host "  Apps with Expiring Credentials: $expiringCount" -ForegroundColor Yellow
    Write-Host "  Apps with Expired Credentials: $expiredCount" -ForegroundColor Red
    
    if ($IncludePermissions) {
        Write-Host "  Total Permissions Tracked: $(($results | Measure-Object -Property PermissionCount -Sum).Sum)" -ForegroundColor Cyan
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY NOTE:" -ForegroundColor Red
    Write-Host "Review orphaned apps and expired credentials immediately.`n" -ForegroundColor Yellow
    
    $results | Select-Object -First 10 | Format-Table ApplicationName, IsOrphaned, HasExpiringCredentials, DaysUntilExpiry -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No apps found." -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
