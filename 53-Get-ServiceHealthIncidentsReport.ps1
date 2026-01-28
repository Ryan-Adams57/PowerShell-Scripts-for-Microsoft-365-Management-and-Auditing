<#
====================================================================================
Script Name: Get-ServiceHealthIncidentsReport.ps1
Description: Microsoft 365 service health incidents and advisories report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves current and historical service health incidents
• Shows service outages, degradations, and advisories
• Lists affected workloads and user impact
• Tracks incident status and resolution times
• Supports filtering by service and time range
• Generates service availability reports
• Helps with SLA tracking and incident management
• Requires Service Health Administrator or Global Reader role

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Incident","Advisory","All")]
    [string]$IssueType = "All",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Active","Resolved","All")]
    [string]$Status = "All",
    
    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 30,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Service_Health_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Service Health Incidents Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Reports"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install $requiredModule? (Y/N)"
    if ($install -eq 'Y' -or $install -eq 'y') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Installed.`n" -ForegroundColor Green
    } else { exit }
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "ServiceHealth.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

$results = @()
$incidentCount = 0
$advisoryCount = 0

Write-Host "Retrieving service health data..." -ForegroundColor Cyan

try {
    $uri = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues"
    $healthIssues = Invoke-MgGraphRequest -Method GET -Uri $uri
    
    if ($healthIssues.value) {
        $issues = $healthIssues.value
        Write-Host "Found $($issues.Count) service health issue(s). Processing...`n" -ForegroundColor Green
        
        $startDate = (Get-Date).AddDays(-$DaysBack)
        
        foreach ($issue in $issues) {
            if ($issue.startDateTime) {
                $issueDate = [DateTime]$issue.startDateTime
                if ($issueDate -lt $startDate) { continue }
            }
            
            $issueClassification = $issue.classification
            if ($IssueType -ne "All" -and $issueClassification -ne $IssueType) { continue }
            
            $issueStatus = $issue.status
            if ($Status -ne "All") {
                if ($Status -eq "Active" -and $issueStatus -ne "serviceOperational") { continue }
                if ($Status -eq "Resolved" -and $issueStatus -eq "serviceOperational") { continue }
            }
            
            if ($issueClassification -eq "Incident") { $incidentCount++ }
            if ($issueClassification -eq "Advisory") { $advisoryCount++ }
            
            $results += [PSCustomObject]@{
                IssueId = $issue.id
                Title = $issue.title
                Classification = $issueClassification
                Status = $issueStatus
                Service = $issue.service
                Feature = $issue.feature
                StartDateTime = $issue.startDateTime
                EndDateTime = $issue.endDateTime
                LastModifiedDateTime = $issue.lastModifiedDateTime
                ImpactDescription = $issue.impactDescription
                Origin = $issue.origin
                IsResolved = $issueStatus -eq "serviceRestored"
            }
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Service Health Summary:" -ForegroundColor Green
    Write-Host "  Total Issues: $($results.Count)" -ForegroundColor White
    Write-Host "  Incidents: $incidentCount" -ForegroundColor Red
    Write-Host "  Advisories: $advisoryCount" -ForegroundColor Yellow
    Write-Host "  Active Issues: $(($results | Where-Object { -not $_.IsResolved }).Count)" -ForegroundColor Yellow
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table Title, Classification, Status, Service, StartDateTime -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No service health issues found." -ForegroundColor Green
}

Disconnect-MgGraph | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
