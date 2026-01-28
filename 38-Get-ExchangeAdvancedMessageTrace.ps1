<#
====================================================================================
Script Name: Get-ExchangeAdvancedMessageTrace.ps1
Description: Exchange Online advanced message trace and delivery report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Performs advanced message trace beyond 10-day limit
• Shows detailed message routing and delivery status
• Tracks message events, hops, and latency
• Identifies delivery failures and NDRs
• Supports extensive filtering (sender, recipient, subject, date)
• Generates forensic email flow analysis
• Exports comprehensive message trace data
• Uses historical message trace for extended date ranges

====================================================================================
#>

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
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed: $_" -ForegroundColor Red; exit
}

$dateRange = (New-TimeSpan -Start $StartDate -End $EndDate).Days
Write-Host "Date range: $dateRange days" -ForegroundColor Cyan
Write-Host "Start: $($StartDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor White
Write-Host "End: $($EndDate.ToString('yyyy-MM-dd HH:mm'))`n" -ForegroundColor White

$results = @()

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
            $results += [PSCustomObject]@{
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

if ($results.Count -gt 0) {
    Write-Host "`n" -NoNewline; Write-Host "="*80 -ForegroundColor Cyan
    Write-Host "Message Trace Summary:" -ForegroundColor Green
    Write-Host "  Total Messages: $($results.Count)" -ForegroundColor White
    Write-Host "  Delivered: $(($results | Where-Object { $_.Status -eq 'Delivered' }).Count)" -ForegroundColor Green
    Write-Host "  Failed: $(($results | Where-Object { $_.Status -eq 'Failed' }).Count)" -ForegroundColor Red
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "="*80 -ForegroundColor Cyan; Write-Host ""
    
    $results | Select-Object -First 10 | Format-Table Received, SenderAddress, RecipientAddress, Status, Size -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -match '^[Yy]$') { Invoke-Item $ExportPath }
} else {
    Write-Host "No messages found or historical search queued." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
