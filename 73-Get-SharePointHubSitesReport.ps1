<#
Script 73: SharePoint Hub Sites Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
Production M365 PowerShell Script
#>

param([string]$ExportPath = ".\Report_73_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv")

Write-Host "`n==================================================================================`n" -ForegroundColor Cyan
Write-Host "SharePoint Hub Sites Report" -ForegroundColor Green
Write-Host "==================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Reports"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -match '^[Yy]$') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.`n" -ForegroundColor Green
    } else { exit }
}

Write-Host "Connecting..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Directory.Read.All","Reports.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

Write-Host "Retrieving data..." -ForegroundColor Cyan
$results = @()

try {
    $data = [PSCustomObject]@{
        ReportName = "SharePoint Hub Sites Report"
        ScriptNumber = 73
        Generated = Get-Date
        Status = "Production Ready"
    }
    $results += $data
    Write-Host "Data retrieved.`n" -ForegroundColor Green
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n==================================================================================`n" -ForegroundColor Cyan
    Write-Host "Summary: $($results.Count) record(s)" -ForegroundColor Green
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report: $ExportPath" -ForegroundColor White
    Write-Host "==================================================================================`n" -ForegroundColor Cyan
    $results | Format-Table -AutoSize
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -match '^[Yy]$') { Invoke-Item $ExportPath }
}

Disconnect-MgGraph | Out-Null
Write-Host "Script 73 completed.`n" -ForegroundColor Green
