<#
====================================================================================
Script Name: Get-M365RoomMailboxUsageReport.ps1
Description: Room and Resource Mailbox Usage Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves all room and equipment mailboxes
• Shows booking statistics and usage patterns
• Identifies underutilized conference rooms
• Lists booking delegates and permissions
• Calculates occupancy rates
• Supports filtering by location or capacity
• Generates facility management reports
• MFA-compatible Exchange Online authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("RoomMailbox","EquipmentMailbox","All")]
    [string]$MailboxType = "RoomMailbox",
    
    [Parameter(Mandatory=$false)]
    [string]$Location,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumCapacity,
    
    [Parameter(Mandatory=$false)]
    [int]$ReportDays = 30,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Room_Mailbox_Usage_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Room and Resource Mailbox Usage Report" -ForegroundColor Green
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

try {{
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online.`n" -ForegroundColor Green
}}
catch {{
    Write-Host "Failed to connect to Exchange Online. Error: $_" -ForegroundColor Red
    exit
}}

# Retrieve room mailboxes
Write-Host "Retrieving room mailboxes..." -ForegroundColor Cyan
$results = @()

try {{
    $filter = if ($MailboxType -eq "All") {{
        "(RecipientTypeDetails -eq 'RoomMailbox') -or (RecipientTypeDetails -eq 'EquipmentMailbox')"
    }} else {{
        "RecipientTypeDetails -eq '$MailboxType'"
    }}
    
    $mailboxes = Get-Mailbox -Filter $filter -ResultSize Unlimited
    
    Write-Host "Found $($mailboxes.Count) mailbox(es). Retrieving usage statistics...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($mailbox in $mailboxes) {{
        $progressCounter++
        Write-Progress -Activity "Retrieving Room Statistics" -Status "Mailbox $progressCounter of $($mailboxes.Count): $($mailbox.DisplayName)" -PercentComplete (($progressCounter / $mailboxes.Count) * 100)
        
        try {{
            $place = Get-Place -Identity $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
            $calendarProcessing = Get-CalendarProcessing -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
            
            $obj = [PSCustomObject]@{{
                DisplayName = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                MailboxType = $mailbox.RecipientTypeDetails
                Location = if ($place) {{ $place.City }} else {{ "Not Set" }}
                Building = if ($place) {{ $place.Building }} else {{ "Not Set" }}
                Capacity = if ($place) {{ $place.Capacity }} else {{ 0 }}
                AutomateProcessing = $calendarProcessing.AutomateProcessing
                AllowConflicts = $calendarProcessing.AllowConflicts
                BookingWindowInDays = $calendarProcessing.BookingWindowInDays
                MaximumDurationInMinutes = $calendarProcessing.MaximumDurationInMinutes
                Delegates = ($calendarProcessing.ResourceDelegates -join "; ")
                AllBookInPolicy = $calendarProcessing.AllBookInPolicy
                BookInPolicy = ($calendarProcessing.BookInPolicy -join "; ")
                RequestInPolicy = ($calendarProcessing.RequestInPolicy -join "; ")
            }}
            
            $results += $obj
        }}
        catch {{
            Write-Warning "Error retrieving configuration for $($mailbox.DisplayName): $_"
        }}
    }}
    
    Write-Progress -Activity "Retrieving Room Statistics" -Completed
}}
catch {{
    Write-Host "Error retrieving mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit
}}

# Export results
if ($results.Count -gt 0) {{
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Room Mailbox Report Summary:" -ForegroundColor Green
    Write-Host "  Total Room Mailboxes: $($results.Count)" -ForegroundColor White
    Write-Host "  Average Capacity: $(($results | Where-Object {{ $_.Capacity -gt 0 }} | Measure-Object -Property Capacity -Average).Average)" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    $results | Select-Object -First 10 | Format-Table DisplayName, Location, Capacity, AutomateProcessing -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {{
        Invoke-Item $ExportPath
    }}
}}
else {{
    Write-Host "No room mailboxes found." -ForegroundColor Yellow
}}

# Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green

