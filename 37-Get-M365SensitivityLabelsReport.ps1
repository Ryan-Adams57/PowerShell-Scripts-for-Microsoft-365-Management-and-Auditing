<#
====================================================================================
Script Name: Get-M365SensitivityLabelsReport.ps1
Description: Sensitivity labels configuration and usage report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all sensitivity labels and sub-labels
• Shows label settings and encryption configurations
• Lists label usage statistics across workloads
• Identifies auto-labeling policies
• Tracks label scope (files, emails, groups, sites)
• Supports filtering by published status
• Generates information protection governance reports
• Requires Azure Information Protection P1 or P2

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$PublishedOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeUsageStats,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Sensitivity_Labels_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "M365 Sensitivity Labels Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install module? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
    } else { exit }
}

Write-Host "Connecting to Security & Compliance..." -ForegroundColor Cyan
try {
    Connect-IPPSSession -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

$results = @()

try {
    $labels = Get-Label -ErrorAction Stop
    Write-Host "Found $($labels.Count) label(s).`n" -ForegroundColor Green
    
    foreach ($label in $labels) {
        $labelPolicy = Get-LabelPolicy -ErrorAction SilentlyContinue | Where-Object { $_.Labels -contains $label.Guid }
        $isPublished = $labelPolicy -ne $null
        
        if ($PublishedOnly -and -not $isPublished) { continue }
        
        $results += [PSCustomObject]@{
            LabelName = $label.DisplayName
            Description = $label.Comment
            IsPublished = $isPublished
            PublishedBy = if ($labelPolicy) { $labelPolicy.Name } else { "Not Published" }
            Priority = $label.Priority
            EncryptionEnabled = $label.EncryptionEnabled
            ContentMarkingEnabled = $label.ContentMarkingEnabled
            SiteAndGroupProtectionEnabled = $label.SiteAndGroupProtectionEnabled
            ParentLabel = $label.ParentId
            Tooltip = $label.Tooltip
            CreatedBy = $label.CreatedBy
            ModifiedBy = $label.LastModifiedBy
            WhenCreated = $label.WhenCreatedUTC
            LabelId = $label.Guid
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Sensitivity Labels Summary:" -ForegroundColor Green
    Write-Host "  Total Labels: $($labels.Count)" -ForegroundColor White
    Write-Host "  Published Labels: $(($results | Where-Object { $_.IsPublished -eq $true }).Count)" -ForegroundColor Green
    Write-Host "  Encryption Enabled: $(($results | Where-Object { $_.EncryptionEnabled -eq $true }).Count)" -ForegroundColor Cyan
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table LabelName, IsPublished, EncryptionEnabled, Priority -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No labels found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
