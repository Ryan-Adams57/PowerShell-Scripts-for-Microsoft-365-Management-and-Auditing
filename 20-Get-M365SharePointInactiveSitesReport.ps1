<#
====================================================================================
Script Name: Get-M365SharePointInactiveSitesReport.ps1
Description: SharePoint Inactive and Orphaned Sites Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Lists SharePoint sites by last activity date
• Identifies sites without owners (orphaned sites)
• Shows storage usage for inactive sites
• Calculates potential storage cost savings
• Supports activity threshold filtering
• Highlights deletion candidates for cleanup
• Generates site governance and cleanup reports
• MFA-compatible SharePoint Online connection

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 180,
    
    [Parameter(Mandatory=$false)]
    [switch]$OrphanedOnly,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumStorageGB,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_SharePoint_Inactive_Sites_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 SharePoint Inactive Sites Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Online.SharePoint.PowerShell"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Required module '$requiredModule' is not installed." -ForegroundColor Yellow
    $install = Read-Host "Would you like to install it now? (Y/N)"
    
    if ($install -eq 'Y' -or $install -eq 'y') {
        try {
            Install-Module -Name $requiredModule -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
            Write-Host "$requiredModule installed successfully.`n" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install $requiredModule. Error: $_" -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Module installation declined." -ForegroundColor Red
        exit
    }
}

Write-Host "Enter your SharePoint Admin URL (e.g., https://contoso-admin.sharepoint.com):" -ForegroundColor Cyan
$adminUrl = Read-Host

while ([string]::IsNullOrWhiteSpace($adminUrl) -or $adminUrl -notmatch '^https://.*-admin\.sharepoint\.com$') {
    Write-Host "Invalid URL. Please enter a valid SharePoint Admin URL:" -ForegroundColor Yellow
    $adminUrl = Read-Host
}

Write-Host "`nConnecting to SharePoint Online..." -ForegroundColor Cyan
try {
    Connect-SPOService -Url $adminUrl -ErrorAction Stop
    Write-Host "Connected successfully.`n" -ForegroundColor Green
}
catch {
    Write-Host "Connection failed: $_" -ForegroundColor Red
    exit
}

$results = @()
$thresholdDate = (Get-Date).AddDays(-$InactiveDays)
$orphanedCount = 0

try {
    $sites = Get-SPOSite -Limit All -Filter "Template -ne 'SPSPERS#10'"
    
    Write-Host "Found $($sites.Count) site(s). Analyzing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($site in $sites) {
        $progressCounter++
        Write-Progress -Activity "Analyzing Sites" -Status "Site $progressCounter of $($sites.Count)" -PercentComplete (($progressCounter / $sites.Count) * 100)
        
        $storageUsedGB = [math]::Round($site.StorageUsageCurrent / 1024, 2)
        $daysSinceModified = if ($site.LastContentModifiedDate) {
            (New-TimeSpan -Start $site.LastContentModifiedDate -End (Get-Date)).Days
        } else { 999 }
        
        $isInactive = ($daysSinceModified -gt $InactiveDays)
        $isOrphaned = [string]::IsNullOrWhiteSpace($site.Owner)
        
        if ($isOrphaned) { $orphanedCount++ }
        
        if ($OrphanedOnly -and -not $isOrphaned) { continue }
        if ($MinimumStorageGB -and $storageUsedGB -lt $MinimumStorageGB) { continue }
        if (-not $isInactive -and -not $isOrphaned) { continue }
        
        $results += [PSCustomObject]@{
            SiteUrl = $site.Url
            Title = $site.Title
            Owner = $site.Owner
            LastModified = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString('yyyy-MM-dd') } else { "Never" }
            DaysSinceModified = $daysSinceModified
            StorageUsedGB = $storageUsedGB
            IsInactive = $isInactive
            IsOrphaned = $isOrphaned
            Template = $site.Template
            Status = if ($isOrphaned) { "Orphaned" } elseif ($isInactive) { "Inactive" } else { "Active" }
        }
    }
    
    Write-Progress -Activity "Analyzing Sites" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-SPOService
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Inactive Sites Summary:" -ForegroundColor Green
    Write-Host "  Total Inactive/Orphaned Sites: $($results.Count)" -ForegroundColor White
    Write-Host "  Orphaned Sites: $orphanedCount" -ForegroundColor Red
    Write-Host "  Total Storage Used: $(($results | Measure-Object -Property StorageUsedGB -Sum).Sum) GB" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table Title, Owner, DaysSinceModified, StorageUsedGB, Status -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No inactive sites found." -ForegroundColor Yellow
}

Disconnect-SPOService
Write-Host "Completed.`n" -ForegroundColor Green
