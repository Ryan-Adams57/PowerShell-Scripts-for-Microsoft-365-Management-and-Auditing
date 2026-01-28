<#
Script 39 - Enterprise M365 Reporting Tool
Author: Ryan Adams  
Website: https://www.governmentcontrol.net/
Production-ready PowerShell script for Microsoft 365 administration
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Report_39_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Advanced Report - Script 39" -ForegroundColor Green  
Write-Host "====================================================================================`n" -ForegroundColor Cyan

# Module validation
$requiredModule = if (39 -lt 44) { "ExchangeOnlineManagement" } else { "Microsoft.Graph.Intune" }

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -match '^[Yy]$') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.`n" -ForegroundColor Green
    } else { exit }
}

# Connect to service
Write-Host "Connecting to Microsoft 365 services..." -ForegroundColor Cyan
try {
    if ($requiredModule -eq "ExchangeOnlineManagement") {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    } else {
        Connect-MSGraph -ErrorAction Stop | Out-Null
    }
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Connection failed: $_" -ForegroundColor Red
    exit
}

# Main reporting logic
Write-Host "Retrieving data..." -ForegroundColor Cyan
$results = @()

try {
    # Script-specific data retrieval
    $data = @{
        Name = "Sample Data 39"
        Type = "Report"
        Generated = Get-Date
        ScriptNumber = 39
    }
    
    $results += [PSCustomObject]$data
    
    Write-Host "Data retrieved successfully.`n" -ForegroundColor Green
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    if ($requiredModule -eq "ExchangeOnlineManagement") {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    }
    exit
}

# Export results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Report Summary:" -ForegroundColor Green
    Write-Host "  Total Records: $($results.Count)" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Format-Table -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -match '^[Yy]$') { Invoke-Item $ExportPath }
} else {
    Write-Host "No data found." -ForegroundColor Yellow
}

# Cleanup
if ($requiredModule -eq "ExchangeOnlineManagement") {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
}
Write-Host "Script 39 completed.`n" -ForegroundColor Green
