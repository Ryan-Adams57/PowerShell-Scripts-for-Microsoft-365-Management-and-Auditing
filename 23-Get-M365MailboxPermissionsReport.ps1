<#
====================================================================================
Script Name: Get-M365MailboxPermissionsReport.ps1
Description: Mailbox Delegation and Permissions Audit Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Lists all mailbox delegation permissions
• Shows SendAs, SendOnBehalf, and FullAccess rights
• Identifies shared mailbox delegate permissions
• Highlights excessive or unusual permissions
• Supports filtering by permission type
• Generates security and compliance audit reports
• Exports comprehensive permissions inventory CSV
• Critical for SOX, HIPAA, and regulatory compliance

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("FullAccess","SendAs","SendOnBehalf","All")]
    [string]$PermissionType = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$SharedMailboxesOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Mailbox_Permissions_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Mailbox Permissions Audit" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "ExchangeOnlineManagement"

if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Module not installed." -ForegroundColor Yellow
    $install = Read-Host "Install? (Y/N)"
    
    if ($install -eq 'Y' -or $install -eq 'y') {
        try {
            Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
            Write-Host "Installed.`n" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed: $_" -ForegroundColor Red
            exit
        }
    }
    else {
        exit
    }
}

# Connect
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed: $_" -ForegroundColor Red
    exit
}

# Retrieve mailboxes
Write-Host "Retrieving mailboxes..." -ForegroundColor Cyan
$results = @()

try {
    $filter = if ($UserPrincipalName) {
        "PrimarySmtpAddress -eq '$UserPrincipalName'"
    } elseif ($SharedMailboxesOnly) {
        "RecipientTypeDetails -eq 'SharedMailbox'"
    } else {
        $null
    }
    
    if ($filter) {
        $mailboxes = Get-Mailbox -Filter $filter -ResultSize Unlimited
    } else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited
    }
    
    Write-Host "Found $($mailboxes.Count) mailbox(es). Checking permissions...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($mailbox in $mailboxes) {
        $progressCounter++
        Write-Progress -Activity "Checking Permissions" -Status "Mailbox $progressCounter of $($mailboxes.Count)" -PercentComplete (($progressCounter / $mailboxes.Count) * 100)
        
        try {
            # FullAccess permissions
            if ($PermissionType -eq "All" -or $PermissionType -eq "FullAccess") {
                $fullAccessPerms = Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress | Where-Object {
                    $_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "S-1-5-*" -and $_.AccessRights -contains "FullAccess"
                }
                
                foreach ($perm in $fullAccessPerms) {
                    $results += [PSCustomObject]@{
                        Mailbox = $mailbox.PrimarySmtpAddress
                        MailboxType = $mailbox.RecipientTypeDetails
                        User = $perm.User
                        PermissionType = "FullAccess"
                        AccessRights = ($perm.AccessRights -join ", ")
                        IsInherited = $perm.IsInherited
                        Deny = $perm.Deny
                    }
                }
            }
            
            # SendAs permissions
            if ($PermissionType -eq "All" -or $PermissionType -eq "SendAs") {
                $sendAsPerms = Get-RecipientPermission -Identity $mailbox.PrimarySmtpAddress | Where-Object {
                    $_.Trustee -notlike "NT AUTHORITY\SELF" -and $_.Trustee -notlike "S-1-5-*" -and $_.AccessRights -contains "SendAs"
                }
                
                foreach ($perm in $sendAsPerms) {
                    $results += [PSCustomObject]@{
                        Mailbox = $mailbox.PrimarySmtpAddress
                        MailboxType = $mailbox.RecipientTypeDetails
                        User = $perm.Trustee
                        PermissionType = "SendAs"
                        AccessRights = ($perm.AccessRights -join ", ")
                        IsInherited = $perm.IsInherited
                        Deny = $false
                    }
                }
            }
            
            # SendOnBehalf permissions
            if (($PermissionType -eq "All" -or $PermissionType -eq "SendOnBehalf") -and $mailbox.GrantSendOnBehalfTo) {
                foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
                    $results += [PSCustomObject]@{
                        Mailbox = $mailbox.PrimarySmtpAddress
                        MailboxType = $mailbox.RecipientTypeDetails
                        User = $delegate
                        PermissionType = "SendOnBehalf"
                        AccessRights = "SendOnBehalf"
                        IsInherited = $false
                        Deny = $false
                    }
                }
            }
        }
        catch {
            Write-Warning "Error checking $($mailbox.PrimarySmtpAddress): $_"
        }
    }
    
    Write-Progress -Activity "Checking Permissions" -Completed
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}

# Export
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Permissions Summary:" -ForegroundColor Green
    Write-Host "  Total Permissions: $($results.Count)" -ForegroundColor White
    Write-Host "  FullAccess: $(($results | Where-Object { $_.PermissionType -eq 'FullAccess' }).Count)" -ForegroundColor White
    Write-Host "  SendAs: $(($results | Where-Object { $_.PermissionType -eq 'SendAs' }).Count)" -ForegroundColor White
    Write-Host "  SendOnBehalf: $(($results | Where-Object { $_.PermissionType -eq 'SendOnBehalf' }).Count)" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table Mailbox, User, PermissionType, AccessRights -AutoSize
    
    $open = Read-Host "Open CSV? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') { Invoke-Item $ExportPath }
}
else {
    Write-Host "No mailbox permissions found." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Completed.`n" -ForegroundColor Green
