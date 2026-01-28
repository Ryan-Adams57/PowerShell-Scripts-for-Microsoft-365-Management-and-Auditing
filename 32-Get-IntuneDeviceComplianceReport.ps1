<#
====================================================================================
Script Name: Get-IntuneDeviceComplianceReport.ps1
Description: Intune device compliance status and policy adherence report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Intune-managed devices and compliance status
• Shows compliant, non-compliant, and grace period devices
• Lists compliance policy assignments and violations
• Identifies devices requiring attention or action
• Tracks last check-in times and device health
• Supports filtering by compliance state and OS
• Generates security and compliance audit reports
• Requires Intune Administrator or Global Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Compliant","NonCompliant","InGracePeriod","All")]
    [string]$ComplianceState = "All",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Windows","iOS","Android","macOS","All")]
    [string]$DeviceOS = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$NonCompliantOnly,
    
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 30,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Intune_Device_Compliance_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Intune Device Compliance Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Intune"

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
Write-Host "Connecting to Microsoft Graph (Intune)..." -ForegroundColor Cyan
try {
    Connect-MSGraph -ErrorAction Stop | Out-Null
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve devices
Write-Host "Retrieving Intune managed devices..." -ForegroundColor Cyan
$results = @()
$compliantCount = 0
$nonCompliantCount = 0
$gracePeriodCount = 0

try {
    $devices = Get-IntuneManagedDevice
    
    Write-Host "Found $($devices.Count) device(s). Checking compliance...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $thresholdDate = (Get-Date).AddDays(-$InactiveDays)
    
    foreach ($device in $devices) {
        $progressCounter++
        Write-Progress -Activity "Processing Devices" -Status "Device $progressCounter of $($devices.Count)" -PercentComplete (($progressCounter / $devices.Count) * 100)
        
        $deviceCompliance = $device.complianceState
        $deviceOS = $device.operatingSystem
        
        # Count by compliance state
        switch ($deviceCompliance) {
            "compliant" { $compliantCount++ }
            "noncompliant" { $nonCompliantCount++ }
            "inGracePeriod" { $gracePeriodCount++ }
        }
        
        # Filter by compliance state
        if ($ComplianceState -ne "All") {
            if ($ComplianceState -eq "Compliant" -and $deviceCompliance -ne "compliant") { continue }
            if ($ComplianceState -eq "NonCompliant" -and $deviceCompliance -ne "noncompliant") { continue }
            if ($ComplianceState -eq "InGracePeriod" -and $deviceCompliance -ne "inGracePeriod") { continue }
        }
        
        if ($NonCompliantOnly -and $deviceCompliance -eq "compliant") { continue }
        
        # Filter by OS
        if ($DeviceOS -ne "All" -and $deviceOS -ne $DeviceOS) { continue }
        
        # Calculate days since last sync
        $daysSinceSync = if ($device.lastSyncDateTime) {
            (New-TimeSpan -Start $device.lastSyncDateTime -End (Get-Date)).Days
        } else {
            999
        }
        
        $isInactive = ($daysSinceSync -gt $InactiveDays)
        
        $obj = [PSCustomObject]@{
            DeviceName = $device.deviceName
            UserPrincipalName = $device.userPrincipalName
            ComplianceState = $deviceCompliance
            OperatingSystem = $deviceOS
            OSVersion = $device.osVersion
            Model = $device.model
            Manufacturer = $device.manufacturer
            SerialNumber = $device.serialNumber
            LastSyncDateTime = $device.lastSyncDateTime
            DaysSinceSync = $daysSinceSync
            IsInactive = $isInactive
            EnrollmentDate = $device.enrolledDateTime
            ManagementAgent = $device.managementAgent
            IsEncrypted = $device.isEncrypted
            IsSupervised = $device.isSupervised
            JailBroken = $device.jailBroken
            DeviceId = $device.id
        }
        
        $results += $obj
    }
    
    Write-Progress -Activity "Processing Devices" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# Export
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Device Compliance Summary:" -ForegroundColor Green
    Write-Host "  Total Devices: $($devices.Count)" -ForegroundColor White
    Write-Host "  Compliant: $compliantCount" -ForegroundColor Green
    Write-Host "  Non-Compliant: $nonCompliantCount" -ForegroundColor Red
    Write-Host "  In Grace Period: $gracePeriodCount" -ForegroundColor Yellow
    Write-Host "  Inactive Devices (>$InactiveDays days): $(($results | Where-Object { $_.IsInactive -eq $true }).Count)" -ForegroundColor Yellow
    
    # OS breakdown
    Write-Host "`n  Devices by OS:" -ForegroundColor Cyan
    $results | Group-Object OperatingSystem | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($nonCompliantCount -gt 0) {
        Write-Host "ACTION REQUIRED: $nonCompliantCount non-compliant device(s)!" -ForegroundColor Red
    }
    
    $results | Select-Object -First 10 | Format-Table DeviceName, UserPrincipalName, ComplianceState, OperatingSystem, DaysSinceSync -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No devices found." -ForegroundColor Yellow
}

Write-Host "Completed.`n" -ForegroundColor Green
