<#
====================================================================================
Script Name: Get-M365DefenderThreatProtectionReport.ps1
Description: Microsoft Defender for Office 365 threat protection analytics
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves Defender for Office 365 threat detections
• Shows Safe Links and Safe Attachments activity
• Identifies ATP policy violations and detections
• Tracks malware family classifications
• Supports date range filtering for threat analysis
• Generates comprehensive threat intelligence reports
• Exports detailed security metrics to CSV
• Requires Microsoft Defender for Office 365 Plan 1 or 2

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [datetime]$StartDate = (Get-Date).AddDays(-30),
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date),
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("SafeLinks","SafeAttachments","AntiPhish","All")]
    [string]$ProtectionType = "All",
    
    [Parameter(Mandatory=$false)]
    [string]$RecipientAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Defender_Threat_Protection_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft Defender for Office 365 Threat Protection Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"

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

# Validate date range
if ($StartDate -gt $EndDate) {
    Write-Host "ERROR: Start date cannot be after end date." -ForegroundColor Red
    exit
}

$dateRange = (New-TimeSpan -Start $StartDate -End $EndDate).Days
if ($dateRange -gt 90) {
    Write-Host "WARNING: Date range exceeds 90 days. Adjusting start date." -ForegroundColor Yellow
    $StartDate = (Get-Date).AddDays(-90)
}

Write-Host "Search Parameters:" -ForegroundColor Cyan
Write-Host "  Start Date: $($StartDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "  End Date: $($EndDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "  Protection Type: $ProtectionType`n" -ForegroundColor White

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve threat protection data
Write-Host "Retrieving Microsoft Defender threat protection data..." -ForegroundColor Cyan
Write-Host "Note: This requires Defender for Office 365 licensing.`n" -ForegroundColor Yellow

$results = @()
$safeLinksCount = 0
$safeAttachmentsCount = 0
$antiPhishCount = 0

try {
    # Get ATP policy information
    Write-Host "Checking ATP policies..." -ForegroundColor Cyan
    
    $safeLinksPolicy = Get-SafeLinksPolicy -ErrorAction SilentlyContinue
    $safeAttachmentPolicy = Get-SafeAttachmentPolicy -ErrorAction SilentlyContinue
    $antiPhishPolicy = Get-AntiPhishPolicy -ErrorAction SilentlyContinue
    
    Write-Host "Found $($safeLinksPolicy.Count) Safe Links policy/policies" -ForegroundColor Green
    Write-Host "Found $($safeAttachmentPolicy.Count) Safe Attachments policy/policies" -ForegroundColor Green
    Write-Host "Found $($antiPhishPolicy.Count) Anti-Phish policy/policies`n" -ForegroundColor Green
    
    # Get threat detections from audit log
    Write-Host "Searching for threat detections in audit log..." -ForegroundColor Cyan
    
    $searchParams = @{
        StartDate = $StartDate
        EndDate = $EndDate
        Operations = "AdvancedThreatProtection"
        ResultSize = 5000
    }
    
    if ($RecipientAddress) {
        $searchParams.Add("UserIds", $RecipientAddress)
    }
    
    $auditRecords = Search-UnifiedAuditLog @searchParams -ErrorAction SilentlyContinue
    
    if ($auditRecords) {
        Write-Host "Found $($auditRecords.Count) ATP detection record(s). Processing...`n" -ForegroundColor Green
        
        $progressCounter = 0
        
        foreach ($record in $auditRecords) {
            $progressCounter++
            Write-Progress -Activity "Processing ATP Detections" -Status "Record $progressCounter of $($auditRecords.Count)" -PercentComplete (($progressCounter / $auditRecords.Count) * 100)
            
            try {
                $auditData = $record.AuditData | ConvertFrom-Json
                
                $detectionType = "Unknown"
                if ($auditData.Workload -like "*SafeLinks*") {
                    $detectionType = "SafeLinks"
                    $safeLinksCount++
                }
                elseif ($auditData.Workload -like "*SafeAttachments*") {
                    $detectionType = "SafeAttachments"
                    $safeAttachmentsCount++
                }
                elseif ($auditData.Workload -like "*AntiPhish*") {
                    $detectionType = "AntiPhish"
                    $antiPhishCount++
                }
                
                # Filter by protection type
                if ($ProtectionType -ne "All" -and $detectionType -ne $ProtectionType) {
                    continue
                }
                
                $obj = [PSCustomObject]@{
                    DetectionTime = $record.CreationDate
                    ProtectionType = $detectionType
                    RecipientAddress = $auditData.RecipientAddress
                    SenderAddress = $auditData.SenderAddress
                    Subject = $auditData.Subject
                    ThreatType = $auditData.ThreatType
                    Action = $auditData.Action
                    FileName = $auditData.FileName
                    URL = $auditData.Url
                    DetectionMethod = $auditData.DetectionMethod
                    Verdict = $auditData.Verdict
                    ClientIP = $auditData.ClientIP
                }
                
                $results += $obj
            }
            catch {
                Write-Warning "Error parsing ATP record: $_"
            }
        }
        
        Write-Progress -Activity "Processing ATP Detections" -Completed
    }
    else {
        Write-Host "No ATP detections found in audit log for the specified date range." -ForegroundColor Yellow
        Write-Host "Generating policy configuration report instead...`n" -ForegroundColor Cyan
        
        # Create policy configuration report
        foreach ($policy in $safeLinksPolicy) {
            $obj = [PSCustomObject]@{
                PolicyType = "SafeLinks"
                PolicyName = $policy.Name
                Enabled = $policy.IsEnabled
                TrackClicks = $policy.TrackClicks
                AllowClickThrough = $policy.AllowClickThrough
                ScanUrls = $policy.ScanUrls
                EnableForInternalSenders = $policy.EnableForInternalSenders
                Priority = $policy.Priority
                AppliedTo = ($policy.AppliedTo -join "; ")
                Configuration = "Policy Active"
            }
            $results += $obj
        }
        
        foreach ($policy in $safeAttachmentPolicy) {
            $obj = [PSCustomObject]@{
                PolicyType = "SafeAttachments"
                PolicyName = $policy.Name
                Enabled = $policy.Enable
                Action = $policy.Action
                Redirect = $policy.Redirect
                RedirectAddress = $policy.RedirectAddress
                Priority = $policy.Priority
                AppliedTo = ($policy.AppliedTo -join "; ")
                Configuration = "Policy Active"
            }
            $results += $obj
        }
        
        foreach ($policy in $antiPhishPolicy) {
            $obj = [PSCustomObject]@{
                PolicyType = "AntiPhish"
                PolicyName = $policy.Name
                Enabled = $policy.Enabled
                EnableMailboxIntelligence = $policy.EnableMailboxIntelligence
                EnableSpoofIntelligence = $policy.EnableSpoofIntelligence
                PhishThresholdLevel = $policy.PhishThresholdLevel
                TargetedUserProtectionAction = $policy.TargetedUserProtectionAction
                Priority = $policy.Priority
                Configuration = "Policy Active"
            }
            $results += $obj
        }
    }
}
catch {
    Write-Host "Error retrieving threat protection data: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Defender for Office 365 Summary:" -ForegroundColor Green
    Write-Host "  Total Records: $($results.Count)" -ForegroundColor White
    
    if ($safeLinksCount -gt 0 -or $safeAttachmentsCount -gt 0 -or $antiPhishCount -gt 0) {
        Write-Host "  Safe Links Detections: $safeLinksCount" -ForegroundColor Yellow
        Write-Host "  Safe Attachments Detections: $safeAttachmentsCount" -ForegroundColor Yellow
        Write-Host "  Anti-Phish Detections: $antiPhishCount" -ForegroundColor Yellow
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY NOTE:" -ForegroundColor Red
    Write-Host "Review ATP detections regularly and adjust policies as needed.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No threat protection data found for the specified criteria." -ForegroundColor Yellow
    Write-Host "Note: Requires Microsoft Defender for Office 365 Plan 1 or 2." -ForegroundColor Cyan
}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
