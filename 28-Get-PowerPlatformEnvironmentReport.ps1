<#
====================================================================================
Script Name: Get-PowerPlatformEnvironmentReport.ps1
Description: Power Platform environments and resource inventory report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Power Platform environments in the tenant
• Shows environment type, region, and security settings
• Lists Power Apps, Power Automate flows, and connectors per environment
• Identifies environments with DLP policies applied
• Tracks environment creation and admin assignments
• Supports filtering by environment type
• Generates governance and compliance reports
• Requires Power Platform Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Production","Sandbox","Trial","Default","All")]
    [string]$EnvironmentType = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeResourceCounts,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowDLPPolicies,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\PowerPlatform_Environment_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Power Platform Environment Report Generator" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.PowerApps.Administration.PowerShell")

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

# Connect to Power Platform
Write-Host "Connecting to Power Platform..." -ForegroundColor Cyan

try {
    Add-PowerAppsAccount -ErrorAction Stop | Out-Null
    Write-Host "Successfully connected to Power Platform.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Power Platform. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve environments
Write-Host "Retrieving Power Platform environments..." -ForegroundColor Cyan
$results = @()
$productionCount = 0
$sandboxCount = 0
$trialCount = 0

try {
    $environments = Get-AdminPowerAppEnvironment
    
    Write-Host "Found $($environments.Count) environment(s). Analyzing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($env in $environments) {
        $progressCounter++
        Write-Progress -Activity "Processing Environments" -Status "Environment $progressCounter of $($environments.Count): $($env.DisplayName)" -PercentComplete (($progressCounter / $environments.Count) * 100)
        
        $envType = $env.EnvironmentType
        
        # Count by type
        switch ($envType) {
            "Production" { $productionCount++ }
            "Sandbox" { $sandboxCount++ }
            "Trial" { $trialCount++ }
        }
        
        # Filter by type
        if ($EnvironmentType -ne "All" -and $envType -ne $EnvironmentType) {
            continue
        }
        
        # Get resource counts if requested
        $appCount = 0
        $flowCount = 0
        $connectorCount = 0
        
        if ($IncludeResourceCounts) {
            try {
                $apps = Get-AdminPowerApp -EnvironmentName $env.EnvironmentName -ErrorAction SilentlyContinue
                $flows = Get-AdminFlow -EnvironmentName $env.EnvironmentName -ErrorAction SilentlyContinue
                $connectors = Get-AdminPowerAppConnector -EnvironmentName $env.EnvironmentName -ErrorAction SilentlyContinue
                
                $appCount = if ($apps) { $apps.Count } else { 0 }
                $flowCount = if ($flows) { $flows.Count } else { 0 }
                $connectorCount = if ($connectors) { $connectors.Count } else { 0 }
            }
            catch {
                Write-Warning "Could not retrieve resource counts for $($env.DisplayName)"
            }
        }
        
        # Get DLP policies if requested
        $dlpPolicies = "Not Retrieved"
        if ($ShowDLPPolicies) {
            try {
                $policies = Get-AdminDlpPolicy -EnvironmentName $env.EnvironmentName -ErrorAction SilentlyContinue
                $dlpPolicies = if ($policies) { ($policies | ForEach-Object { $_.DisplayName }) -join "; " } else { "None" }
            }
            catch {
                $dlpPolicies = "Error Retrieving"
            }
        }
        
        # Parse security settings
        $securityGroup = if ($env.SecurityGroupId) { $env.SecurityGroupId } else { "None" }
        
        $obj = [PSCustomObject]@{
            EnvironmentName = $env.DisplayName
            EnvironmentId = $env.EnvironmentName
            Type = $envType
            Region = $env.Location.Name
            CreatedTime = $env.CreatedTime
            CreatedBy = $env.CreatedBy.DisplayName
            IsDefault = $env.IsDefault
            SecurityGroupId = $securityGroup
            PowerAppsCount = $appCount
            FlowsCount = $flowCount
            ConnectorsCount = $connectorCount
            DLPPolicies = $dlpPolicies
            InternalState = $env.States.Management.Id
            RuntimeState = $env.States.Runtime.Id
            LinkedEnvironmentUrl = $env.LinkedEnvironmentMetadata.InstanceUrl
        }
        
        $results += $obj
    }
    
    Write-Progress -Activity "Processing Environments" -Completed
}
catch {
    Write-Host "Error retrieving environments: $_" -ForegroundColor Red
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Power Platform Environment Summary:" -ForegroundColor Green
    Write-Host "  Total Environments: $($results.Count)" -ForegroundColor White
    Write-Host "  Production Environments: $productionCount" -ForegroundColor Green
    Write-Host "  Sandbox Environments: $sandboxCount" -ForegroundColor Yellow
    Write-Host "  Trial Environments: $trialCount" -ForegroundColor Cyan
    
    if ($IncludeResourceCounts) {
        Write-Host "`n  Total Resources:" -ForegroundColor Cyan
        Write-Host "    Power Apps: $(($results | Measure-Object -Property PowerAppsCount -Sum).Sum)" -ForegroundColor White
        Write-Host "    Flows: $(($results | Measure-Object -Property FlowsCount -Sum).Sum)" -ForegroundColor White
        Write-Host "    Connectors: $(($results | Measure-Object -Property ConnectorsCount -Sum).Sum)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "GOVERNANCE NOTE:" -ForegroundColor Cyan
    Write-Host "Review environments regularly for compliance and security best practices.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table EnvironmentName, Type, Region, PowerAppsCount, FlowsCount -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No Power Platform environments found matching the specified criteria." -ForegroundColor Yellow
}

Write-Host "Script completed successfully.`n" -ForegroundColor Green
