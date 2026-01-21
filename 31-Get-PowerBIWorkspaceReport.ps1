<#
====================================================================================
Script Name: Get-PowerBIWorkspaceReport.ps1
Description: Power BI workspace, dataset, and report inventory
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all Power BI workspaces in the tenant
• Lists datasets, reports, and dashboards per workspace
• Shows workspace ownership and member access
• Identifies orphaned workspaces without admins
• Tracks premium capacity assignments
• Supports filtering by workspace type
• Generates usage and governance analytics
• Requires Power BI Administrator role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Workspace","Group","PersonalGroup","All")]
    [string]$WorkspaceType = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDatasets,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeReports,
    
    [Parameter(Mandatory=$false)]
    [switch]$OrphanedOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\PowerBI_Workspace_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Power BI Workspace and Dataset Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "MicrosoftPowerBIMgmt"

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
Write-Host "Connecting to Power BI..." -ForegroundColor Cyan
try {
    Connect-PowerBIServiceAccount -ErrorAction Stop | Out-Null
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve workspaces
Write-Host "Retrieving Power BI workspaces..." -ForegroundColor Cyan
$results = @()
$orphanedCount = 0
$premiumCount = 0

try {
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All
    
    Write-Host "Found $($workspaces.Count) workspace(s). Processing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($workspace in $workspaces) {
        $progressCounter++
        Write-Progress -Activity "Processing Workspaces" -Status "Workspace $progressCounter of $($workspaces.Count)" -PercentComplete (($progressCounter / $workspaces.Count) * 100)
        
        $wsType = $workspace.Type
        
        # Filter by type
        if ($WorkspaceType -ne "All" -and $wsType -ne $WorkspaceType) { continue }
        
        # Check for admins
        $admins = $workspace.Users | Where-Object { $_.AccessRight -eq "Admin" }
        $isOrphaned = ($admins.Count -eq 0)
        
        if ($isOrphaned) { $orphanedCount++ }
        if ($OrphanedOnly -and -not $isOrphaned) { continue }
        
        # Check premium capacity
        $isPremium = $workspace.IsOnDedicatedCapacity
        if ($isPremium) { $premiumCount++ }
        
        # Get datasets if requested
        $datasetCount = 0
        $datasetNames = "Not Retrieved"
        if ($IncludeDatasets) {
            try {
                $datasets = Get-PowerBIDataset -Scope Organization -WorkspaceId $workspace.Id -ErrorAction SilentlyContinue
                if ($datasets) {
                    $datasetCount = $datasets.Count
                    $datasetNames = ($datasets | ForEach-Object { $_.Name }) -join "; "
                }
                else {
                    $datasetNames = "None"
                }
            }
            catch {
                $datasetNames = "Error"
            }
        }
        
        # Get reports if requested
        $reportCount = 0
        $reportNames = "Not Retrieved"
        if ($IncludeReports) {
            try {
                $reports = Get-PowerBIReport -Scope Organization -WorkspaceId $workspace.Id -ErrorAction SilentlyContinue
                if ($reports) {
                    $reportCount = $reports.Count
                    $reportNames = ($reports | ForEach-Object { $_.Name }) -join "; "
                }
                else {
                    $reportNames = "None"
                }
            }
            catch {
                $reportNames = "Error"
            }
        }
        
        # Parse members
        $adminList = ($admins | ForEach-Object { $_.UserPrincipalName }) -join "; "
        $memberCount = $workspace.Users.Count
        
        $obj = [PSCustomObject]@{
            WorkspaceName = $workspace.Name
            WorkspaceId = $workspace.Id
            Type = $wsType
            State = $workspace.State
            IsOrphaned = $isOrphaned
            AdminCount = $admins.Count
            Admins = if ($adminList) { $adminList } else { "None" }
            TotalMembers = $memberCount
            DatasetCount = $datasetCount
            Datasets = $datasetNames
            ReportCount = $reportCount
            Reports = $reportNames
            IsPremium = $isPremium
            CapacityId = $workspace.CapacityId
            Description = $workspace.Description
        }
        
        $results += $obj
    }
    
    Write-Progress -Activity "Processing Workspaces" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-PowerBIServiceAccount | Out-Null
    exit
}

# Export
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Power BI Workspace Summary:" -ForegroundColor Green
    Write-Host "  Total Workspaces: $($results.Count)" -ForegroundColor White
    Write-Host "  Orphaned Workspaces: $orphanedCount" -ForegroundColor Red
    Write-Host "  Premium Workspaces: $premiumCount" -ForegroundColor Cyan
    
    if ($IncludeDatasets) {
        Write-Host "  Total Datasets: $(($results | Measure-Object -Property DatasetCount -Sum).Sum)" -ForegroundColor White
    }
    if ($IncludeReports) {
        Write-Host "  Total Reports: $(($results | Measure-Object -Property ReportCount -Sum).Sum)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($orphanedCount -gt 0) {
        Write-Host "WARNING: $orphanedCount orphaned workspace(s) without admins!" -ForegroundColor Red
    }
    
    $results | Select-Object -First 10 | Format-Table WorkspaceName, Type, IsOrphaned, AdminCount, DatasetCount, ReportCount -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No workspaces found." -ForegroundColor Yellow
}

Disconnect-PowerBIServiceAccount | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
