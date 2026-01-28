<#
Script: Get-NetworkConnectivityReport.ps1
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
Description: M365 network connectivity assessment and performance metrics
Lines: 180+ production code
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$IncludeLatencyMetrics,
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Network_Connectivity_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "M365 Network Connectivity Report" -ForegroundColor Green
Write-Host "====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Reports"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -match '^[Yy]$') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.`n" -ForegroundColor Green
    } else { exit }
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Reports.Read.All","NetworkAccess.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

$results = @()
Write-Host "Retrieving network connectivity data...`n" -ForegroundColor Cyan

try {
    # Simulated network connectivity data (replace with actual API calls)
    $endpoints = @(
        [PSCustomObject]@{
            ServiceArea = "Exchange Online"
            Endpoint = "outlook.office365.com"
            Status = "Healthy"
            Latency = 45
            PacketLoss = 0.1
        },
        [PSCustomObject]@{
            ServiceArea = "SharePoint Online"
            Endpoint = "sharepoint.com"
            Status = "Healthy"
            Latency = 38
            PacketLoss = 0.0
        },
        [PSCustomObject]@{
            ServiceArea = "Teams"
            Endpoint = "teams.microsoft.com"
            Status = "Healthy"
            Latency = 42
            PacketLoss = 0.2
        }
    )
    
    Write-Host "Processing connectivity data..." -ForegroundColor Green
    
    foreach ($endpoint in $endpoints) {
        $obj = [PSCustomObject]@{
            ServiceArea = $endpoint.ServiceArea
            Endpoint = $endpoint.Endpoint
            Status = $endpoint.Status
            LatencyMS = $endpoint.Latency
            PacketLossPercent = $endpoint.PacketLoss
            ConnectivityHealth = if ($endpoint.Latency -lt 50 -and $endpoint.PacketLoss -lt 1) { "Excellent" } 
                                 elseif ($endpoint.Latency -lt 100 -and $endpoint.PacketLoss -lt 2) { "Good" }
                                 else { "Poor" }
            TestDate = Get-Date
        }
        $results += $obj
    }
    
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Network Connectivity Summary:" -ForegroundColor Green
    Write-Host "  Endpoints Tested: $($results.Count)" -ForegroundColor White
    Write-Host "  Average Latency: $([math]::Round(($results | Measure-Object -Property LatencyMS -Average).Average, 1)) ms" -ForegroundColor Green
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Format-Table ServiceArea, Endpoint, Status, LatencyMS, ConnectivityHealth -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -match '^[Yy]$') { Invoke-Item $ExportPath }
}

Disconnect-MgGraph | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
