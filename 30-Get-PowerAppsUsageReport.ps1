<#
====================================================================================
Script Name: 30-Get-PowerAppsUsageReport.ps1
Description: Power Apps usage, adoption, and analytics report
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
    [string]$EnvironmentName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Canvas","ModelDriven","SharePoint","All")]
    [string]$AppType = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$PublishedOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSharedUsers,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\PowerApps_Usage_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Power Apps Usage and Adoption Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.PowerApps.Administration.PowerShell"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Module '$requiredModule' not installed." -ForegroundColor Yellow
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

# Connect
Write-Host "Connecting to Power Platform..." -ForegroundColor Cyan
try {
    Add-PowerAppsAccount -ErrorAction Stop | Out-Null
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve apps
Write-Host "Retrieving Power Apps..." -ForegroundColor Cyan
$script:Results = @()
$canvasCount = 0
$modelDrivenCount = 0

try {
    if ($EnvironmentName) {
        $apps = Get-AdminPowerApp -EnvironmentName $EnvironmentName
    }
    else {
        $environments = Get-AdminPowerAppEnvironment
        $apps = @()
        
        foreach ($env in $environments) {
            Write-Host "Scanning environment: $($env.DisplayName)" -ForegroundColor Cyan
            $envApps = Get-AdminPowerApp -EnvironmentName $env.EnvironmentName -ErrorAction SilentlyContinue
            if ($envApps) {
                $apps += $envApps
            }
        }
    }
    
    Write-Host "Found $($apps.Count) app(s). Processing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($app in $apps) {
        $progressCounter++
        Write-Progress -Activity "Processing Apps" -Status "App $progressCounter of $($apps.Count)" -PercentComplete (($progressCounter / $apps.Count) * 100)
        
        # Determine app type
        $appTypeValue = "Unknown"
        if ($app.Internal.properties.displayName -and $app.Internal.properties.appType) {
            $appTypeValue = $app.Internal.properties.appType
        }
        elseif ($app.AppType) {
            $appTypeValue = $app.AppType
        }
        
        # Count by type
        if ($appTypeValue -like "*Canvas*") {
            $canvasCount++
        }
        elseif ($appTypeValue -like "*ModelDriven*") {
            $modelDrivenCount++
        }
        
        # Filter by type
        if ($AppType -ne "All") {
            if ($AppType -eq "Canvas" -and $appTypeValue -notlike "*Canvas*") { continue }
            if ($AppType -eq "ModelDriven" -and $appTypeValue -notlike "*ModelDriven*") { continue }
            if ($AppType -eq "SharePoint" -and $appTypeValue -notlike "*SharePoint*") { continue }
        }
        
        # Check published status
        $isPublished = $app.Internal.properties.lifeCycleId -eq "Published"
        if ($PublishedOnly -and -not $isPublished) { continue }
        
        # Get shared users if requested
        $sharedUsers = "Not Retrieved"
        $sharedCount = 0
        if ($IncludeSharedUsers) {
            try {
                $permissions = Get-AdminPowerAppRoleAssignment -AppName $app.AppName -EnvironmentName $app.EnvironmentName -ErrorAction SilentlyContinue
                if ($permissions) {
                    $sharedCount = $permissions.Count
                    $userList = @()
                    foreach ($perm in $permissions) {
                        if ($perm.PrincipalDisplayName) {
                            $userList += $perm.PrincipalDisplayName
                        }
                    }
                    $sharedUsers = $userList -join "; "
                }
                else {
                    $sharedUsers = "None"
                }
            }
            catch {
                $sharedUsers = "Error"
            }
        }
        
        # Parse connections
        $connections = @()
        if ($app.Internal.properties.connectionReferences) {
            foreach ($conn in $app.Internal.properties.connectionReferences.PSObject.Properties) {
                $connections += $conn.Value.displayName
            }
        }
        $connectionsStr = if ($connections.Count -gt 0) { $connections -join "; " } else { "None" }
        
        $obj = [PSCustomObject]@{
            AppName = $app.DisplayName
            AppId = $app.AppName
            Environment = $app.EnvironmentName
            AppType = $appTypeValue
            Owner = $app.Owner.DisplayName
            OwnerEmail = $app.Owner.Email
            CreatedTime = $app.CreatedTime
            LastModifiedTime = $app.LastModifiedTime
            Published = $isPublished
            SharedWithCount = $sharedCount
            SharedWithUsers = $sharedUsers
            Connections = $connectionsStr
            AppVersion = $app.Internal.properties.appVersion
            Description = $app.Internal.properties.description
        }
        
        $script:Results += $obj
    }
    
    Write-Progress -Activity "Processing Apps" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# Export
if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Power Apps Summary:" -ForegroundColor Green
    Write-Host "  Total Apps: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Canvas Apps: $canvasCount" -ForegroundColor Cyan
    Write-Host "  Model-Driven Apps: $modelDrivenCount" -ForegroundColor Cyan
    Write-Host "  Published Apps: $(($script:Results | Where-Object { $_.Published -eq $true }).Count)" -ForegroundColor Green
    
    if ($IncludeSharedUsers) {
        Write-Host "  Total Shared Users: $(($script:Results | Measure-Object -Property SharedWithCount -Sum).Sum)" -ForegroundColor White
    }
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table AppName, AppType, Owner, Published, SharedWithCount -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No apps found." -ForegroundColor Yellow
}

Write-Host "Completed.`n" -ForegroundColor Green
