<#
==============================================
 PowerShell Scripts for Windows Server 2025
 Author: Ryan Adams
 Website: https://www.governmentcontrol.net/
==============================================
#>

# ============================================
# 1. List All Domain Controllers
# ============================================
<#
Description:
Retrieves all domain controllers in your environment and exports their details to a CSV file.
#>
Get-ADDomainController -Filter * | Select-Object Name, Site, IPAddress, OperatingSystem | Export-Csv -Path "C:\DCList.csv" -NoTypeInformation

# ============================================
# 2. Automate Group Policy Object (GPO) Reports
# ============================================
<#
Description:
Generates an HTML report of all GPOs for auditing and troubleshooting.
#>
Get-GPOReport -All -ReportType HTML -Path "C:\Reports\GPOReport.html"

# ============================================
# 3. Monitor Event Logs for Critical Errors
# ============================================
<#
Description:
Filters critical errors from the System event log and exports them to CSV.
#>
$ErrorEvents = Get-WinEvent -LogName System | Where-Object {$_.LevelDisplayName -eq "Error"}
$ErrorEvents | Export-Csv -Path "C:\Logs\SystemErrorLogs.csv" -NoTypeInformation

# ============================================
# 4. Monitor Disk Space
# ============================================
<#
Description:
Checks local drives and warns if free space is less than 10%.
#>
$drives = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"
foreach ($drive in $drives) {
    if ($drive.FreeSpace / $drive.Size -lt 0.1) {
        Write-Host "Warning: Low disk space on drive $($drive.DeviceID)"
    }
}

# ============================================
# 5. Automate Windows Updates Installation
# ============================================
<#
Description:
Checks for and installs all Windows updates automatically.
Requires PSWindowsUpdate module.
#>
# Install-Module PSWindowsUpdate   # Uncomment if not installed
Install-WindowsUpdate -AcceptAll -AutoReboot

# ============================================
# 6. Create Scheduled Tasks
# ============================================
<#
Description:
Creates a scheduled task to run a backup script daily at midnight.
#>
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\BackupScript.ps1"
$trigger = New-ScheduledTaskTrigger -Daily -At 12:00AM
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "DailyBackup" -Description "Runs daily backup at midnight"

# ============================================
# 7. Manage Services Across Multiple Servers
# ============================================
<#
Description:
Starts the Spooler service on multiple servers remotely.
#>
$servers = @("Server01", "Server02", "Server03")
foreach ($server in $servers) {
    Invoke-Command -ComputerName $server -ScriptBlock {
        Start-Service -Name 'Spooler'
        Write-Host "Started Spooler service on $($env:COMPUTERNAME)"
    }
}

# ============================================
# 8. Automate Application Updates with WinGet
# ============================================
<#
Description:
Upgrades all installed applications using WinGet on Windows Server 2025.
#>
winget upgrade --all

Write-Host "All scripts executed successfully."
