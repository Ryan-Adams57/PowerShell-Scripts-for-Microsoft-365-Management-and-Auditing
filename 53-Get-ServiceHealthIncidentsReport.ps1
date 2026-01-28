# ====================================================================================
# M365 POWERSHELL REPORTING SCRIPT - PRODUCTION READY
# Enterprise-grade reporting with comprehensive error handling
# ====================================================================================
#
<#
====================================================================================
Script Name: 53-Get-ServiceHealthIncidentsReport.ps1
Description: Microsoft 365 service health incidents and advisories report
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

REQUIREMENTS:
• PowerShell 5.1 or higher
• Appropriate M365 administrator permissions
• Required modules (validated at runtime)

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]
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

Set-StrictMode -Version Latest

# Comprehensive logging function
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info","Warning","Error","Success")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Info" { "Cyan" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] $Level: $Message" -ForegroundColor $color
    if ($Level -eq "Error") {
        $script:ErrorCount++
    }
}

# Parameter validation function
function Test-ScriptParameters {
    Write-Log "Validating script parameters..." -Level Info
    if ($ExportPath -and -not (Test-Path (Split-Path $ExportPath -Parent) -ErrorAction SilentlyContinue)) {
        Write-Log "Export path directory does not exist. Creating..." -Level Warning
        try {
            New-Item -Path (Split-Path $ExportPath -Parent) -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Log "Directory created successfully" -Level Success
        } catch {
            Write-Log "Failed to create directory: $($_.Exception.Message)" -Level Error
            return $false
        }
    }
    Write-Log "Parameter validation complete" -Level Success
    return $true
}

$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:Results = @()
$script:ErrorCount = 0

Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "M365 Reporting Script - Production Ready" -ForegroundColor Green
Write-Host "Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
Write-Host "\n====================================================================================\n" -ForegroundColor Cyan

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

$script:Results = @()
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
            
            $script:Results += [PSCustomObject]@{
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

if ($script:Results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Service Health Summary:" -ForegroundColor Green
    Write-Host "  Total Issues: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Incidents: $incidentCount" -ForegroundColor Red
    Write-Host "  Advisories: $advisoryCount" -ForegroundColor Yellow
    Write-Host "  Active Issues: $(($script:Results | Where-Object { -not $_.IsResolved }).Count)" -ForegroundColor Yellow
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $script:Results | Select-Object -First 10 | Format-Table Title, Classification, Status, Service, StartDateTime -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
} else {
    Write-Host "No service health issues found." -ForegroundColor Green
}

Disconnect-MgGraph | Out-Null

# Comprehensive cleanup and summary
Write-Log "Performing cleanup operations..." -Level Info

$script:EndTime = Get-Date
$script:Duration = $script:EndTime - $script:StartTime

try {
    # Disconnect from services
    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    if (Get-Command Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Log "Disconnected from services" -Level Success
} catch {
    Write-Log "Disconnect completed with warnings" -Level Warning
}

Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "Execution Summary:" -ForegroundColor Green
Write-Host "  Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor White
Write-Host "  Duration: $($script:Duration.TotalSeconds) seconds" -ForegroundColor White
Write-Host "  Results Collected: $($script:Results.Count)" -ForegroundColor White
Write-Host "  Errors Encountered: $script:ErrorCount" -ForegroundColor White
if ($ExportPath -and (Test-Path $ExportPath)) {
    Write-Host "  Report Location: $ExportPath" -ForegroundColor Cyan
}
Write-Host "\n====================================================================================\n" -ForegroundColor Cyan

if ($script:ErrorCount -eq 0) {
    Write-Log "Script completed successfully!" -Level Success
    exit 0
} else {
    Write-Log "Script completed with $script:ErrorCount error(s)" -Level Warning
    exit 1
}
Write-Host "Completed.`n" -ForegroundColor Green
