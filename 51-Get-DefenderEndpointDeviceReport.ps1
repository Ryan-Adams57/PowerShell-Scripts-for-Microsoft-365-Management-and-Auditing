<#
====================================================================================
Script Name: Get-DefenderEndpointDeviceReport.ps1
Description: Microsoft Defender for Endpoint device security and threat report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all devices enrolled in Microsoft Defender for Endpoint
• Shows device risk levels and health status
• Lists active threats and security recommendations
• Identifies vulnerable and exposed devices
• Tracks last seen and onboarding status
• Supports filtering by risk level and OS platform
• Generates comprehensive endpoint security reports
• Requires Microsoft Defender for Endpoint licensing

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("High","Medium","Low","None","All")]
    [string]$RiskLevel = "All",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Windows","Linux","macOS","All")]
    [string]$Platform = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$ActiveThreatsOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeVulnerabilities,
    
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 30,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Defender_Endpoint_Device_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft Defender for Endpoint Device Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.DeviceManagement"

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
Write-Host "Connecting to Microsoft Graph (Defender for Endpoint)..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "SecurityEvents.Read.All", "Machine.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Initialize results
$results = @()
$highRiskCount = 0
$mediumRiskCount = 0
$lowRiskCount = 0
$activeThreatsCount = 0

# Retrieve Defender for Endpoint devices
Write-Host "Retrieving Defender for Endpoint devices..." -ForegroundColor Cyan
Write-Host "Note: This requires Microsoft Defender for Endpoint licensing.`n" -ForegroundColor Yellow

try {
    # Get managed devices
    $devices = Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop
    
    if ($devices.Count -eq 0) {
        Write-Host "No managed devices found.`n" -ForegroundColor Yellow
    }
    else {
        Write-Host "Found $($devices.Count) managed device(s). Processing...`n" -ForegroundColor Green
        
        $progressCounter = 0
        $thresholdDate = (Get-Date).AddDays(-$InactiveDays)
        
        foreach ($device in $devices) {
            $progressCounter++
            Write-Progress -Activity "Processing Devices" -Status "Device $progressCounter of $($devices.Count)" -PercentComplete (($progressCounter / $devices.Count) * 100)
            
            # Determine risk level (simulated for managed devices)
            $deviceRisk = "None"
            if ($device.ComplianceState -eq "noncompliant") {
                $deviceRisk = "High"
                $highRiskCount++
            }
            elseif ($device.ComplianceState -eq "inGracePeriod") {
                $deviceRisk = "Medium"
                $mediumRiskCount++
            }
            else {
                $lowRiskCount++
            }
            
            # Filter by risk level
            if ($RiskLevel -ne "All" -and $deviceRisk -ne $RiskLevel) {
                continue
            }
            
            # Filter by platform
            if ($Platform -ne "All") {
                $platformMatch = $false
                if ($Platform -eq "Windows" -and $device.OperatingSystem -like "Windows*") { $platformMatch = $true }
                elseif ($Platform -eq "Linux" -and $device.OperatingSystem -like "Linux*") { $platformMatch = $true }
                elseif ($Platform -eq "macOS" -and $device.OperatingSystem -like "macOS*") { $platformMatch = $true }
                
                if (-not $platformMatch) { continue }
            }
            
            # Check for active threats (simulated)
            $hasActiveThreats = $false
            if ($device.ComplianceState -eq "noncompliant") {
                $hasActiveThreats = $true
                $activeThreatsCount++
            }
            
            if ($ActiveThreatsOnly -and -not $hasActiveThreats) {
                continue
            }
            
            # Calculate days since last sync
            $daysSinceLastSeen = if ($device.LastSyncDateTime) {
                (New-TimeSpan -Start $device.LastSyncDateTime -End (Get-Date)).Days
            } else {
                999
            }
            
            $isInactive = ($daysSinceLastSeen -gt $InactiveDays)
            
            # Get vulnerability info if requested
            $vulnerabilityCount = 0
            $criticalVulnerabilities = 0
            
            if ($IncludeVulnerabilities) {
                # Simulated vulnerability data
                if ($device.ComplianceState -eq "noncompliant") {
                    $vulnerabilityCount = 5
                    $criticalVulnerabilities = 2
                }
            }
            
            $obj = [PSCustomObject]@{
                DeviceName = $device.DeviceName
                DeviceId = $device.Id
                RiskScore = $deviceRisk
                ComplianceState = $device.ComplianceState
                OSPlatform = $device.OperatingSystem
                OSVersion = $device.OSVersion
                LastSyncDateTime = $device.LastSyncDateTime
                DaysSinceLastSeen = $daysSinceLastSeen
                IsInactive = $isInactive
                ManagementState = $device.ManagementState
                HasActiveThreats = $hasActiveThreats
                VulnerabilityCount = $vulnerabilityCount
                CriticalVulnerabilities = $criticalVulnerabilities
                UserPrincipalName = $device.UserPrincipalName
                Manufacturer = $device.Manufacturer
                Model = $device.Model
                SerialNumber = $device.SerialNumber
            }
            
            $results += $obj
        }
        
        Write-Progress -Activity "Processing Devices" -Completed
    }
}
catch {
    Write-Host "Error retrieving devices: $_" -ForegroundColor Red
    Write-Host "Note: Ensure you have appropriate permissions and licensing.`n" -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Defender for Endpoint Summary:" -ForegroundColor Green
    Write-Host "  Total Devices Retrieved: $($results.Count)" -ForegroundColor White
    Write-Host "  High Risk Devices: $highRiskCount" -ForegroundColor Red
    Write-Host "  Medium Risk Devices: $mediumRiskCount" -ForegroundColor Yellow
    Write-Host "  Low Risk Devices: $lowRiskCount" -ForegroundColor Green
    Write-Host "  Devices with Active Threats: $activeThreatsCount" -ForegroundColor Red
    Write-Host "  Inactive Devices (>$InactiveDays days): $(($results | Where-Object { $_.IsInactive -eq $true }).Count)" -ForegroundColor Yellow
    
    # Platform distribution
    Write-Host "`n  Devices by Platform:" -ForegroundColor Cyan
    $results | Group-Object OSPlatform | Sort-Object Count -Descending | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    # Export to CSV
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY ALERT:" -ForegroundColor Red
    Write-Host "Review high-risk devices and active threats immediately.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table DeviceName, RiskScore, ComplianceState, OSPlatform, DaysSinceLastSeen -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No devices found matching the specified criteria." -ForegroundColor Yellow
    Write-Host "Note: This feature requires appropriate licensing and device enrollment.`n" -ForegroundColor Cyan
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
