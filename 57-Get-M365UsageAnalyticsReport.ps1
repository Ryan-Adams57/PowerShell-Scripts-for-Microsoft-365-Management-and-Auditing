<#
====================================================================================
Script Name: Get-M365UsageAnalyticsReport.ps1
Description: Microsoft 365 usage analytics and adoption metrics across all workloads
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves M365 service usage across all workloads
• Shows active users per service (Exchange, Teams, SharePoint, OneDrive)
• Lists adoption trends and growth metrics over time
• Identifies inactive services and low adoption areas
• Tracks storage consumption across services
• Supports custom date ranges for trend analysis
• Generates executive-level adoption dashboards
• Helps optimize license allocation and ROI

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("7","30","90","180")]
    [string]$Period = "30",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeStorageMetrics,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeTrendAnalysis,
    
    [Parameter(Mandatory=$false)]
    [switch]$DetailedBreakdown,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Usage_Analytics_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Usage Analytics Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Reports"

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

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "Reports.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

$results = @()
$totalActiveUsers = 0
$exchangeUsers = 0
$teamsUsers = 0
$sharePointUsers = 0
$oneDriveUsers = 0

Write-Host "Retrieving M365 usage data for period: D$Period..." -ForegroundColor Cyan
Write-Host "This may take a few moments...`n" -ForegroundColor Yellow

try {
    # Get active user counts
    $activeUsersUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserCounts(period='D$Period')"
    $activeUsers = Invoke-MgGraphRequest -Method GET -Uri $activeUsersUri -ErrorAction Stop
    
    # Get services user counts
    $servicesUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ServicesUserCounts(period='D$Period')"
    $services = Invoke-MgGraphRequest -Method GET -Uri $servicesUri -ErrorAction Stop
    
    if ($services) {
        $serviceData = $services | ConvertFrom-Csv
        
        Write-Host "Found $($serviceData.Count) usage record(s). Processing...`n" -ForegroundColor Green
        
        $progressCounter = 0
        
        foreach ($record in $serviceData) {
            $progressCounter++
            
            if ($progressCounter % 10 -eq 0) {
                Write-Progress -Activity "Processing Usage Data" -Status "Record $progressCounter of $($serviceData.Count)" -PercentComplete (($progressCounter / $serviceData.Count) * 100)
            }
            
            # Parse active user counts
            $exchangeActive = if ($record.'Exchange Active') { [int]$record.'Exchange Active' } else { 0 }
            $oneDriveActive = if ($record.'OneDrive Active') { [int]$record.'OneDrive Active' } else { 0 }
            $sharePointActive = if ($record.'SharePoint Active') { [int]$record.'SharePoint Active' } else { 0 }
            $teamsActive = if ($record.'Teams Active') { [int]$record.'Teams Active' } else { 0 }
            $yammerActive = if ($record.'Yammer Active') { [int]$record.'Yammer Active' } else { 0 }
            $office365Active = if ($record.'Office 365 Active') { [int]$record.'Office 365 Active' } else { 0 }
            
            # Track latest counts
            if ($progressCounter -eq 1) {
                $totalActiveUsers = $office365Active
                $exchangeUsers = $exchangeActive
                $teamsUsers = $teamsActive
                $sharePointUsers = $sharePointActive
                $oneDriveUsers = $oneDriveActive
            }
            
            # Calculate adoption percentages
            $exchangeAdoption = if ($office365Active -gt 0) { [math]::Round(($exchangeActive / $office365Active) * 100, 1) } else { 0 }
            $teamsAdoption = if ($office365Active -gt 0) { [math]::Round(($teamsActive / $office365Active) * 100, 1) } else { 0 }
            $sharePointAdoption = if ($office365Active -gt 0) { [math]::Round(($sharePointActive / $office365Active) * 100, 1) } else { 0 }
            $oneDriveAdoption = if ($office365Active -gt 0) { [math]::Round(($oneDriveActive / $office365Active) * 100, 1) } else { 0 }
            
            $obj = [PSCustomObject]@{
                ReportDate = $record.'Report Date'
                ReportPeriod = "D$Period"
                Office365Active = $office365Active
                ExchangeActive = $exchangeActive
                ExchangeAdoption = "$exchangeAdoption%"
                TeamsActive = $teamsActive
                TeamsAdoption = "$teamsAdoption%"
                SharePointActive = $sharePointActive
                SharePointAdoption = "$sharePointAdoption%"
                OneDriveActive = $oneDriveActive
                OneDriveAdoption = "$oneDriveAdoption%"
                YammerActive = $yammerActive
            }
            
            $results += $obj
        }
        
        Write-Progress -Activity "Processing Usage Data" -Completed
    }
    
    # Get storage metrics if requested
    if ($IncludeStorageMetrics) {
        Write-Host "Retrieving storage metrics..." -ForegroundColor Cyan
        
        try {
            $storageUri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageStorage(period='D$Period')"
            $storage = Invoke-MgGraphRequest -Method GET -Uri $storageUri -ErrorAction SilentlyContinue
            
            if ($storage) {
                $storageData = $storage | ConvertFrom-Csv
                Write-Host "Storage data retrieved: $($storageData.Count) record(s)`n" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Storage metrics unavailable.`n" -ForegroundColor Yellow
        }
    }
    
    # Trend analysis if requested
    if ($IncludeTrendAnalysis -and $results.Count -gt 1) {
        Write-Host "Calculating trend analysis..." -ForegroundColor Cyan
        
        $firstRecord = $results | Select-Object -Last 1
        $latestRecord = $results | Select-Object -First 1
        
        $growthRate = if ($firstRecord.Office365Active -gt 0) {
            [math]::Round((($latestRecord.Office365Active - $firstRecord.Office365Active) / $firstRecord.Office365Active) * 100, 1)
        } else { 0 }
        
        Write-Host "User growth rate: $growthRate%`n" -ForegroundColor $(if ($growthRate -gt 0) { "Green" } else { "Yellow" })
    }
}
catch {
    Write-Host "Error retrieving usage data: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "M365 Usage Analytics Summary:" -ForegroundColor Green
    Write-Host "  Report Period: Last $Period days" -ForegroundColor White
    Write-Host "  Total Active Users: $totalActiveUsers" -ForegroundColor Green
    Write-Host "`n  Service Breakdown:" -ForegroundColor Cyan
    Write-Host "    Exchange Active Users: $exchangeUsers" -ForegroundColor White
    Write-Host "    Teams Active Users: $teamsUsers" -ForegroundColor White
    Write-Host "    SharePoint Active Users: $sharePointUsers" -ForegroundColor White
    Write-Host "    OneDrive Active Users: $oneDriveUsers" -ForegroundColor White
    
    # Calculate adoption percentages
    if ($totalActiveUsers -gt 0) {
        Write-Host "`n  Adoption Rates:" -ForegroundColor Cyan
        Write-Host "    Exchange: $([math]::Round(($exchangeUsers / $totalActiveUsers) * 100, 1))%" -ForegroundColor Green
        Write-Host "    Teams: $([math]::Round(($teamsUsers / $totalActiveUsers) * 100, 1))%" -ForegroundColor Green
        Write-Host "    SharePoint: $([math]::Round(($sharePointUsers / $totalActiveUsers) * 100, 1))%" -ForegroundColor Green
        Write-Host "    OneDrive: $([math]::Round(($oneDriveUsers / $totalActiveUsers) * 100, 1))%" -ForegroundColor Green
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "INSIGHTS:" -ForegroundColor Cyan
    Write-Host "Use this data to optimize license allocation and drive adoption.`n" -ForegroundColor Yellow
    
    if ($DetailedBreakdown) {
        Write-Host "Sample Usage Data (First 10 records):" -ForegroundColor Yellow
        $results | Select-Object -First 10 | Format-Table ReportDate, Office365Active, ExchangeActive, TeamsActive, SharePointActive -AutoSize
    }
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No usage data found for the specified period." -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
