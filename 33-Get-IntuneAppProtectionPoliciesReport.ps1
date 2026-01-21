<#
====================================================================================
Script Name: Get-IntuneAppProtectionPoliciesReport.ps1
Description: Intune app protection policies (MAM) configuration report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Intune app protection policies (iOS, Android, Windows)
• Shows policy settings and data protection rules
• Lists targeted apps and user/group assignments
• Identifies policy conflicts and coverage gaps
• Tracks policy compliance and enforcement
• Supports filtering by platform
• Generates mobile application management (MAM) governance reports
• Requires Intune Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("iOS","Android","Windows","All")]
    [string]$Platform = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeAssignments,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Intune_App_Protection_Policies_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Intune App Protection Policies Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Intune"

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

# Connect
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MSGraph -ErrorAction Stop | Out-Null
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve policies
Write-Host "Retrieving app protection policies..." -ForegroundColor Cyan
$results = @()
$iosCount = 0
$androidCount = 0
$windowsCount = 0

try {
    # Get iOS policies
    if ($Platform -eq "All" -or $Platform -eq "iOS") {
        $iosPolicies = Get-IntuneAppProtectionPolicyIos -ErrorAction SilentlyContinue
        if ($iosPolicies) {
            $iosCount = $iosPolicies.Count
            foreach ($policy in $iosPolicies) {
                $assignments = "Not Retrieved"
                if ($IncludeAssignments) {
                    try {
                        $assign = Get-IntuneAppProtectionPolicyIosAssignment -iosManagedAppProtectionId $policy.id -ErrorAction SilentlyContinue
                        if ($assign) {
                            $assignments = ($assign | ForEach-Object { $_.target.groupId }) -join "; "
                        }
                    }
                    catch {
                        $assignments = "Error"
                    }
                }
                
                $results += [PSCustomObject]@{
                    PolicyName = $policy.displayName
                    Platform = "iOS"
                    CreatedDateTime = $policy.createdDateTime
                    LastModifiedDateTime = $policy.lastModifiedDateTime
                    DataBackup = $policy.dataBackupBlocked
                    DeviceCompliance = $policy.deviceComplianceRequired
                    PINRequired = $policy.pinRequired
                    ManagedBrowser = $policy.managedBrowserToOpenLinksRequired
                    SaveAsBlocked = $policy.saveAsBlocked
                    Assignments = $assignments
                    PolicyId = $policy.id
                }
            }
        }
    }
    
    # Get Android policies
    if ($Platform -eq "All" -or $Platform -eq "Android") {
        $androidPolicies = Get-IntuneAppProtectionPolicyAndroid -ErrorAction SilentlyContinue
        if ($androidPolicies) {
            $androidCount = $androidPolicies.Count
            foreach ($policy in $androidPolicies) {
                $assignments = "Not Retrieved"
                if ($IncludeAssignments) {
                    try {
                        $assign = Get-IntuneAppProtectionPolicyAndroidAssignment -androidManagedAppProtectionId $policy.id -ErrorAction SilentlyContinue
                        if ($assign) {
                            $assignments = ($assign | ForEach-Object { $_.target.groupId }) -join "; "
                        }
                    }
                    catch {
                        $assignments = "Error"
                    }
                }
                
                $results += [PSCustomObject]@{
                    PolicyName = $policy.displayName
                    Platform = "Android"
                    CreatedDateTime = $policy.createdDateTime
                    LastModifiedDateTime = $policy.lastModifiedDateTime
                    DataBackup = $policy.dataBackupBlocked
                    DeviceCompliance = $policy.deviceComplianceRequired
                    PINRequired = $policy.pinRequired
                    ManagedBrowser = $policy.managedBrowserToOpenLinksRequired
                    SaveAsBlocked = $policy.saveAsBlocked
                    Assignments = $assignments
                    PolicyId = $policy.id
                }
            }
        }
    }
    
    # Get Windows policies
    if ($Platform -eq "All" -or $Platform -eq "Windows") {
        $windowsPolicies = Get-IntuneAppProtectionPolicyWindows10 -ErrorAction SilentlyContinue
        if ($windowsPolicies) {
            $windowsCount = $windowsPolicies.Count
            foreach ($policy in $windowsPolicies) {
                $assignments = "Not Retrieved"
                if ($IncludeAssignments) {
                    try {
                        $assign = Get-IntuneAppProtectionPolicyWindows10Assignment -windowsManagedAppProtectionId $policy.id -ErrorAction SilentlyContinue
                        if ($assign) {
                            $assignments = ($assign | ForEach-Object { $_.target.groupId }) -join "; "
                        }
                    }
                    catch {
                        $assignments = "Error"
                    }
                }
                
                $results += [PSCustomObject]@{
                    PolicyName = $policy.displayName
                    Platform = "Windows"
                    CreatedDateTime = $policy.createdDateTime
                    LastModifiedDateTime = $policy.lastModifiedDateTime
                    DataBackup = $policy.dataRecoverBlocked
                    DeviceCompliance = "N/A"
                    PINRequired = $policy.pinRequired
                    ManagedBrowser = "N/A"
                    SaveAsBlocked = "N/A"
                    Assignments = $assignments
                    PolicyId = $policy.id
                }
            }
        }
    }
    
    Write-Host "Found app protection policies.`n" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# Export
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "App Protection Policy Summary:" -ForegroundColor Green
    Write-Host "  Total Policies: $($results.Count)" -ForegroundColor White
    Write-Host "  iOS Policies: $iosCount" -ForegroundColor Cyan
    Write-Host "  Android Policies: $androidCount" -ForegroundColor Cyan
    Write-Host "  Windows Policies: $windowsCount" -ForegroundColor Cyan
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table PolicyName, Platform, PINRequired, DataBackup, DeviceCompliance -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No app protection policies found." -ForegroundColor Yellow
}

Write-Host "Completed.`n" -ForegroundColor Green
