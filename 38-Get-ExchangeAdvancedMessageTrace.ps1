# ====================================================================================
# M365 POWERSHELL REPORTING SCRIPT - PRODUCTION READY
# This script provides comprehensive reporting capabilities for Microsoft 365
# Designed for enterprise environments with proper error handling
# ====================================================================================
#
<#
====================================================================================
Script Name: 38-Get-ExchangeAdvancedMessageTrace.ps1
Description: Exchange Online advanced message trace and delivery report
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
    [datetime]$StartDate = (Get-Date).AddDays(-7),
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date),
    
    [Parameter(Mandatory=$false)]
    [string]$SenderAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$RecipientAddress,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Delivered","Failed","Pending","All")]
    [string]$Status = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\Exchange_Advanced_Message_Trace_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Initialize comprehensive error handling
Set-StrictMode -Version Latest

# Logging function for consistent output
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info","Warning","Error","Success")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Info" { "Cyan" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
    }
    Write-Host "[$timestamp] $Level: $Message" -ForegroundColor $color
}

$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:Results = @()
$script:ErrorCount = 0

# Display script information
Write-Host "Script: $($MyInvocation.MyCommand.Name)" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
Write-Host "" -ForegroundColor Cyan

Write-Host "`n" -NoNewline; Write-Host "="*80 -ForegroundColor Cyan
Write-Host "Exchange Online Advanced Message Trace Report" -ForegroundColor Green
Write-Host "="*80 -ForegroundColor Cyan; Write-Host ""

$requiredModule = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    $install = Read-Host "Install module? (Y/N)"
    if ($install -match '^[Yy]$') {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
    } else { exit }
}

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red; exit
}

$dateRange = (New-TimeSpan -Start $StartDate -End $EndDate).Days
Write-Host "Date range: $dateRange days" -ForegroundColor Cyan
Write-Host "Start: $($StartDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor White
Write-Host "End: $($EndDate.ToString('yyyy-MM-dd HH:mm'))`n" -ForegroundColor White

$script:Results = @()

try {
    $traceParams = @{
        StartDate = $StartDate
        EndDate = $EndDate
        PageSize = 5000
    }
    
    if ($SenderAddress) { $traceParams.Add("SenderAddress", $SenderAddress) }
    if ($RecipientAddress) { $traceParams.Add("RecipientAddress", $RecipientAddress) }
    if ($Status -ne "All") { $traceParams.Add("Status", $Status) }
    
    if ($dateRange -le 10) {
        Write-Host "Using Get-MessageTrace (last 10 days)..." -ForegroundColor Cyan
        $messages = Get-MessageTrace @traceParams
    } else {
        Write-Host "Using Get-HistoricalSearch (beyond 10 days)..." -ForegroundColor Cyan
        $jobName = "AdvancedTrace_$(Get-Date -Format 'yyyyMMddHHmmss')"
        Start-HistoricalSearch @traceParams -ReportTitle $jobName -ReportType MessageTrace
        Write-Host "Historical search initiated. Job: $jobName" -ForegroundColor Yellow
        Write-Host "Note: Results will be available via email when complete.`n" -ForegroundColor Yellow
        $messages = @()
    }
    
    if ($messages) {
        Write-Host "Found $($messages.Count) message(s). Processing...`n" -ForegroundColor Green
        
        foreach ($msg in $messages) {
            $script:Results += [PSCustomObject]@{
                Received = $msg.Received
                SenderAddress = $msg.SenderAddress
                RecipientAddress = $msg.RecipientAddress
                Subject = $msg.Subject
                Status = $msg.Status
                FromIP = $msg.FromIP
                ToIP = $msg.ToIP
                Size = $msg.Size
                MessageId = $msg.MessageId
                MessageTraceId = $msg.MessageTraceId
            }
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

if ($script:Results.Count -gt 0) {
    Write-Host "`n" -NoNewline; Write-Host "="*80 -ForegroundColor Cyan
    Write-Host "Message Trace Summary:" -ForegroundColor Green
    Write-Host "  Total Messages: $($script:Results.Count)" -ForegroundColor White
    Write-Host "  Delivered: $(($script:Results | Where-Object { $_.Status -eq 'Delivered' }).Count)" -ForegroundColor Green
    Write-Host "  Failed: $(($script:Results | Where-Object { $_.Status -eq 'Failed' }).Count)" -ForegroundColor Red
    
    $script:Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "="*80 -ForegroundColor Cyan; Write-Host ""
    
    $script:Results | Select-Object -First 10 | Format-Table Received, SenderAddress, RecipientAddress, Status, Size -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -match '^[Yy]$') { Invoke-Item $ExportPath }
} else {
    Write-Host "No messages found or historical search queued." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null

# Comprehensive cleanup and summary
$script:EndTime = Get-Date
$script:Duration = $script:EndTime - $script:StartTime

Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "Execution Summary:" -ForegroundColor Green
Write-Host "  Duration: $($script:Duration.TotalSeconds) seconds" -ForegroundColor White
Write-Host "  Results: $($script:Results.Count) items" -ForegroundColor White
Write-Host "  Errors: $script:ErrorCount" -ForegroundColor White
Write-Host "\n====================================================================================\n" -ForegroundColor Cyan
Write-Host "Completed.`n" -ForegroundColor Green
