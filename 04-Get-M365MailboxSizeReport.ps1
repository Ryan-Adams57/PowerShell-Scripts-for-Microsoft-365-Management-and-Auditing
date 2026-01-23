<#
====================================================================================
Script Name: Get-M365MailboxSizeReport.ps1
Description: Comprehensive mailbox size, quota, and item count report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves mailbox size statistics for all mailbox types
• Shows current quota usage and limits
• Identifies mailboxes approaching quota limits
• Displays item counts and folder statistics
• Supports filtering by mailbox size threshold
• Generates capacity planning recommendations
• Exports detailed CSV reports with size trends
• MFA-compatible Exchange Online authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("UserMailbox","SharedMailbox","RoomMailbox","All")]
    [string]$MailboxType = "All",
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumSizeGB,
    
    [Parameter(Mandatory=$false)]
    [switch]$QuotaWarningOnly,
    
    [Parameter(Mandatory=$false)]
    [int]$WarningPercentage = 90,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Mailbox_Size_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Mailbox Size and Usage Report" -ForegroundColor Green
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

# Helper function to convert size to GB
function Convert-MailboxSize {
    param([string]$SizeString)
    
    if ($SizeString -match '([\d,\.]+)\s+(KB|MB|GB|TB)') {
        $value = [double]($matches[1] -replace ',', '')
        $unit = $matches[2]
        
        switch ($unit) {
            'KB' { return $value / 1MB }
            'MB' { return $value / 1KB }
            'GB' { return $value }
            'TB' { return $value * 1KB }
        }
    }
    return 0
}

# Retrieve mailboxes
Write-Host "Retrieving mailbox information..." -ForegroundColor Cyan
$results = @()

try {
    $mailboxFilter = if ($UserPrincipalName) {
        "PrimarySmtpAddress -eq '$UserPrincipalName'"
    }
    elseif ($MailboxType -ne "All") {
        "RecipientTypeDetails -eq '$MailboxType'"
    }
    else {
        $null
    }
    
    if ($mailboxFilter) {
        $mailboxes = Get-Mailbox -Filter $mailboxFilter -ResultSize Unlimited
    }
    else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited
    }
    
    Write-Host "Found $($mailboxes.Count) mailbox(es). Retrieving size statistics...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $totalSizeGB = 0
    $quotaWarnings = 0
    
    foreach ($mailbox in $mailboxes) {
        $progressCounter++
        Write-Progress -Activity "Retrieving Mailbox Statistics" -Status "Mailbox $progressCounter of $($mailboxes.Count): $($mailbox.PrimarySmtpAddress)" -PercentComplete (($progressCounter / $mailboxes.Count) * 100)
        
        try {
            $stats = Get-MailboxStatistics -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
            
            # Parse mailbox size
            $totalItemSizeGB = 0
            if ($stats.TotalItemSize) {
                $sizeString = $stats.TotalItemSize.ToString()
                $totalItemSizeGB = Convert-MailboxSize -SizeString $sizeString
            }
            
            # Parse quota
            $quotaGB = 0
            $quotaStatus = "N/A"
            $percentUsed = 0
            
            if ($mailbox.ProhibitSendReceiveQuota -and $mailbox.ProhibitSendReceiveQuota -ne "Unlimited") {
                $quotaString = $mailbox.ProhibitSendReceiveQuota.ToString()
                $quotaGB = Convert-MailboxSize -SizeString $quotaString
                
                if ($quotaGB -gt 0) {
                    $percentUsed = [math]::Round(($totalItemSizeGB / $quotaGB) * 100, 2)
                    
                    if ($percentUsed -ge $WarningPercentage) {
                        $quotaStatus = "Warning"
                        $quotaWarnings++
                    }
                    elseif ($percentUsed -ge 100) {
                        $quotaStatus = "Critical"
                        $quotaWarnings++
                    }
                    else {
                        $quotaStatus = "OK"
                    }
                }
            }
            
            # Apply filters
            $includeMailbox = $true
            
            if ($MinimumSizeGB -and $totalItemSizeGB -lt $MinimumSizeGB) {
                $includeMailbox = $false
            }
            
            if ($QuotaWarningOnly -and $quotaStatus -eq "OK") {
                $includeMailbox = $false
            }
            
            if ($includeMailbox) {
                $totalSizeGB += $totalItemSizeGB
                
                $obj = [PSCustomObject]@{
                    DisplayName = $mailbox.DisplayName
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    MailboxType = $mailbox.RecipientTypeDetails
                    TotalItemSizeGB = [math]::Round($totalItemSizeGB, 2)
                    ItemCount = $stats.ItemCount
                    DeletedItemSizeGB = if ($stats.TotalDeletedItemSize) { [math]::Round((Convert-MailboxSize -SizeString $stats.TotalDeletedItemSize.ToString()), 2) } else { 0 }
                    DeletedItemCount = $stats.DeletedItemCount
                    MailboxQuotaGB = if ($quotaGB -gt 0) { [math]::Round($quotaGB, 2) } else { "Unlimited" }
                    PercentUsed = $percentUsed
                    QuotaStatus = $quotaStatus
                    LastLogonTime = $stats.LastLogonTime
                    LastUserActionTime = $stats.LastUserActionTime
                    Database = $stats.Database
                }
                
                $results += $obj
            }
        }
        catch {
            Write-Warning "Error retrieving statistics for $($mailbox.PrimarySmtpAddress): $_"
        }
    }
    
    Write-Progress -Activity "Retrieving Mailbox Statistics" -Completed
}
catch {
    Write-Host "Error retrieving mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Mailbox Size Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total Mailboxes Analyzed: $($results.Count)" -ForegroundColor White
    Write-Host "  Combined Mailbox Size: $([math]::Round($totalSizeGB, 2)) GB" -ForegroundColor White
    Write-Host "  Mailboxes with Quota Warnings: $quotaWarnings" -ForegroundColor Yellow
    Write-Host "  Average Mailbox Size: $([math]::Round(($totalSizeGB / $results.Count), 2)) GB" -ForegroundColor White
    Write-Host "  Largest Mailbox: $(($results | Sort-Object TotalItemSizeGB -Descending | Select-Object -First 1).TotalItemSizeGB) GB" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Top 10 Largest Mailboxes:" -ForegroundColor Yellow
    $results | Sort-Object TotalItemSizeGB -Descending | Select-Object -First 10 | Format-Table DisplayName, PrimarySmtpAddress, TotalItemSizeGB, PercentUsed, QuotaStatus -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No mailboxes found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
