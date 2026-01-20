<#
====================================================================================
Script Name: Get-M365UnusedLicensesReport.ps1
Description: Unused and Available License Inventory Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Shows total vs assigned licenses per SKU
• Calculates unused license count and percentage
• Estimates potential monthly cost savings
• Lists all unassigned license types in tenant
• Highlights over-provisioned SKUs
• Generates license optimization recommendations
• Exports cost analysis and utilization CSV
• Critical for budget planning and optimization

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$ShowCostEstimates,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumUnused = 1,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Unused_Licenses_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Unused Licenses Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Identity.DirectoryManagement"

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
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Organization.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# License cost estimates (approximate monthly USD)
$licenseCosts = @{
    "SPE_E3" = 36
    "SPE_E5" = 57
    "ENTERPRISEPREMIUM" = 57
    "ENTERPRISEPACK" = 23
    "STANDARDPACK" = 12.50
    "SPB" = 22
    "POWER_BI_PRO" = 10
    "DESKLESSPACK" = 10
}

Write-Host "Retrieving license subscriptions..." -ForegroundColor Cyan
$results = @()
$totalUnused = 0
$totalWaste = 0

try {
    $subscriptions = Get-MgSubscribedSku -All
    
    Write-Host "Found $($subscriptions.Count) license SKU(s).`n" -ForegroundColor Green
    
    foreach ($sku in $subscriptions) {
        $enabled = $sku.PrepaidUnits.Enabled
        $consumed = $sku.ConsumedUnits
        $unused = $enabled - $consumed
        
        if ($unused -lt $MinimumUnused) {
            continue
        }
        
        $percentUnused = if ($enabled -gt 0) { [math]::Round(($unused / $enabled) * 100, 2) } else { 0 }
        
        # Estimate cost
        $monthlyCost = 0
        $estimatedWaste = 0
        
        if ($ShowCostEstimates -and $licenseCosts.ContainsKey($sku.SkuPartNumber)) {
            $monthlyCost = $licenseCosts[$sku.SkuPartNumber]
            $estimatedWaste = $unused * $monthlyCost
            $totalWaste += $estimatedWaste
        }
        
        $totalUnused += $unused
        
        $results += [PSCustomObject]@{
            LicenseName = $sku.SkuPartNumber
            TotalLicenses = $enabled
            AssignedLicenses = $consumed
            UnusedLicenses = $unused
            PercentUnused = $percentUnused
            EstimatedMonthlyCost = if ($monthlyCost -gt 0) { $monthlyCost } else { "N/A" }
            EstimatedWaste = if ($estimatedWaste -gt 0) { [math]::Round($estimatedWaste, 2) } else { "N/A" }
            SkuId = $sku.SkuId
            Status = if ($percentUnused -gt 50) { "High Waste" } elseif ($percentUnused -gt 25) { "Moderate Waste" } else { "Low Waste" }
        }
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "License Utilization Summary:" -ForegroundColor Green
    Write-Host "  Total License SKUs: $($results.Count)" -ForegroundColor White
    Write-Host "  Total Unused Licenses: $totalUnused" -ForegroundColor Yellow
    
    if ($ShowCostEstimates -and $totalWaste -gt 0) {
        Write-Host "  Estimated Monthly Waste: `$$([math]::Round($totalWaste, 2))" -ForegroundColor Red
        Write-Host "  Estimated Annual Waste: `$$([math]::Round($totalWaste * 12, 2))" -ForegroundColor Red
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Sort-Object UnusedLicenses -Descending | Select-Object -First 10 | Format-Table LicenseName, TotalLicenses, AssignedLicenses, UnusedLicenses, PercentUnused -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No unused licenses found." -ForegroundColor Green
}

Disconnect-MgGraph | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
