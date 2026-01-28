<#
====================================================================================
Script Name: Get-PowerAutomateFlowsInventory.ps1
Description: Power Automate cloud flows inventory and usage report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Power Automate flows across all environments
• Shows flow state (enabled, disabled, suspended)
• Identifies flow owners and last modified dates
• Tracks connector usage and dependencies
• Lists trigger types and run frequency
• Supports filtering by environment and flow state
• Generates governance and usage analytics
• Requires Power Platform Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Started","Stopped","Suspended","All")]
    [string]$FlowState = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$EnabledOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeConnections,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\PowerAutomate_Flows_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Power Automate Flows Inventory Report" -ForegroundColor Green
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

# Retrieve flows
Write-Host "Retrieving Power Automate flows..." -ForegroundColor Cyan
$results = @()
$enabledCount = 0
$disabledCount = 0

try {
    if ($EnvironmentName) {
        $flows = Get-AdminFlow -EnvironmentName $EnvironmentName
    }
    else {
        $environments = Get-AdminPowerAppEnvironment
        $flows = @()
        
        foreach ($env in $environments) {
            Write-Host "Scanning environment: $($env.DisplayName)" -ForegroundColor Cyan
            $envFlows = Get-AdminFlow -EnvironmentName $env.EnvironmentName -ErrorAction SilentlyContinue
            if ($envFlows) {
                $flows += $envFlows
            }
        }
    }
    
    Write-Host "Found $($flows.Count) flow(s). Processing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($flow in $flows) {
        $progressCounter++
        Write-Progress -Activity "Processing Flows" -Status "Flow $progressCounter of $($flows.Count)" -PercentComplete (($progressCounter / $flows.Count) * 100)
        
        $flowEnabled = $flow.Enabled
        $flowState = $flow.FlowState
        
        if ($flowEnabled) { $enabledCount++ } else { $disabledCount++ }
        
        if ($EnabledOnly -and -not $flowEnabled) { continue }
        if ($FlowState -ne "All" -and $flowState -ne $FlowState) { continue }
        
        # Get connections if requested
        $connections = "Not Retrieved"
        if ($IncludeConnections) {
            try {
                $connProps = $flow.Internal.properties.connectionReferences
                if ($connProps) {
                    $connNames = @()
                    foreach ($conn in $connProps.PSObject.Properties) {
                        $connNames += $conn.Value.displayName
                    }
                    $connections = $connNames -join "; "
                }
                else {
                    $connections = "None"
                }
            }
            catch {
                $connections = "Error"
            }
        }
        
        $obj = [PSCustomObject]@{
            FlowName = $flow.DisplayName
            FlowId = $flow.FlowName
            Environment = $flow.EnvironmentName
            Enabled = $flowEnabled
            State = $flowState
            CreatedTime = $flow.CreatedTime
            LastModifiedTime = $flow.LastModifiedTime
            CreatedBy = $flow.CreatedBy.DisplayName
            Owner = $flow.Internal.properties.creator.userPrincipalName
            TriggerType = $flow.Internal.properties.definitionSummary.triggers[0].type
            Connections = $connections
            FlowType = if ($flow.Internal.properties.isManaged) { "Managed" } else { "Unmanaged" }
        }
        
        $results += $obj
    }
    
    Write-Progress -Activity "Processing Flows" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit
}

# Export
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Flow Inventory Summary:" -ForegroundColor Green
    Write-Host "  Total Flows: $($results.Count)" -ForegroundColor White
    Write-Host "  Enabled: $enabledCount" -ForegroundColor Green
    Write-Host "  Disabled: $disabledCount" -ForegroundColor Yellow
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table FlowName, Enabled, State, Owner, TriggerType -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No flows found." -ForegroundColor Yellow
}

Write-Host "Completed.`n" -ForegroundColor Green
