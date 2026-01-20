<#
====================================================================================
Script Name: Get-M365OneDriveUsageReport.ps1
Description: Comprehensive OneDrive for Business storage and usage analysis
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves OneDrive site storage usage for all users
• Shows total storage allocated and consumed
• Identifies inactive OneDrive sites
• Calculates storage utilization percentages
• Highlights users approaching storage limits
• Supports filtering by storage size or activity
• Generates capacity planning insights
• MFA-compatible SharePoint Online authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [int]$MinimumStorageGB,
    
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory=$false)]
    [switch]$InactiveSitesOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$OverQuotaOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_OneDrive_Usage_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 OneDrive Usage Report Generator" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.Online.SharePoint.PowerShell", "Microsoft.Graph.Users")

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Required module '$module' is not installed." -ForegroundColor Yellow
        $install = Read-Host "Would you like to install it now? (Y/N)"
        
        if ($install -eq 'Y' -or $install -eq 'y') {
            try {
                Write-Host "Installing $module..." -ForegroundColor Cyan
                Install-Module -Name $module -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
                Write-Host "$module installed successfully.`n" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to install $module. Error: $_" -ForegroundColor Red
                exit
            }
        }
        else {
            Write-Host "Module installation declined. Script cannot continue." -ForegroundColor Red
            exit
        }
    }
}

# Get tenant admin URL
Write-Host "Please enter your SharePoint Admin URL (e.g., https://contoso-admin.sharepoint.com):" -ForegroundColor Cyan
$adminUrl = Read-Host

while ([string]::IsNullOrWhiteSpace($adminUrl) -or $adminUrl -notmatch '^https://.*-admin\.sharepoint\.com$') {
    Write-Host "Invalid URL format. Please enter a valid SharePoint Admin URL:" -ForegroundColor Yellow
    $adminUrl = Read-Host
}

# Connect to SharePoint Online
Write-Host "`nConnecting to SharePoint Online..." -ForegroundColor Cyan

try {
    Connect-SPOService -Url $adminUrl -ErrorAction Stop
    Write-Host "Successfully connected to SharePoint Online.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to SharePoint Online. Error: $_" -ForegroundColor Red
    exit
}

# Also connect to Microsoft Graph for user details
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    Disconnect-SPOService
    exit
}

# Retrieve OneDrive sites
Write-Host "Retrieving OneDrive for Business sites..." -ForegroundColor Cyan
$results = @()
$thresholdDate = (Get-Date).AddDays(-$InactiveDays)

try {
    if ($UserPrincipalName) {
        $sites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Owner -eq '$UserPrincipalName' -and Template -eq 'SPSPERS#10'"
    }
    else {
        $sites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Template -eq 'SPSPERS#10'"
    }
    
    Write-Host "Found $($sites.Count) OneDrive site(s). Analyzing storage usage...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $totalStorageGB = 0
    $totalUsedGB = 0
    $inactiveCount = 0
    $overQuotaCount = 0
    
    foreach ($site in $sites) {
        $progressCounter++
        Write-Progress -Activity "Analyzing OneDrive Sites" -Status "Site $progressCounter of $($sites.Count): $($site.Owner)" -PercentComplete (($progressCounter / $sites.Count) * 100)
        
        try {
            # Calculate storage metrics
            $storageUsedGB = [math]::Round($site.StorageUsageCurrent / 1024, 2)
            $storageQuotaGB = [math]::Round($site.StorageQuota / 1024, 2)
            $percentUsed = if ($storageQuotaGB -gt 0) { [math]::Round(($storageUsedGB / $storageQuotaGB) * 100, 2) } else { 0 }
            
            # Determine activity status
            $daysSinceModified = if ($site.LastContentModifiedDate) {
                (New-TimeSpan -Start $site.LastContentModifiedDate -End (Get-Date)).Days
            } else {
                999
            }
            
            $isInactive = ($daysSinceModified -gt $InactiveDays)
            $isOverQuota = ($percentUsed -gt 100)
            
            if ($isInactive) { $inactiveCount++ }
            if ($isOverQuota) { $overQuotaCount++ }
            
            # Apply filters
            $includeSite = $true
            
            if ($InactiveSitesOnly -and -not $isInactive) {
                $includeSite = $false
            }
            
            if ($OverQuotaOnly -and -not $isOverQuota) {
                $includeSite = $false
            }
            
            if ($MinimumStorageGB -and $storageUsedGB -lt $MinimumStorageGB) {
                $includeSite = $false
            }
            
            if ($includeSite) {
                $totalStorageGB += $storageQuotaGB
                $totalUsedGB += $storageUsedGB
                
                # Get user details from Graph
                $userDisplayName = "Unknown"
                $userEnabled = "Unknown"
                
                try {
                    $owner = $site.Owner
                    if ($owner) {
                        $mgUser = Get-MgUser -UserId $owner -Property DisplayName, AccountEnabled -ErrorAction SilentlyContinue
                        if ($mgUser) {
                            $userDisplayName = $mgUser.DisplayName
                            $userEnabled = $mgUser.AccountEnabled
                        }
                    }
                }
                catch {
                    # Continue with Unknown values
                }
                
                $obj = [PSCustomObject]@{
                    DisplayName = $userDisplayName
                    Owner = $site.Owner
                    AccountEnabled = $userEnabled
                    SiteUrl = $site.Url
                    StorageUsedGB = $storageUsedGB
                    StorageQuotaGB = $storageQuotaGB
                    PercentUsed = $percentUsed
                    LastModifiedDate = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
                    DaysSinceModified = $daysSinceModified
                    IsInactive = $isInactive
                    IsOverQuota = $isOverQuota
                    Status = if ($isOverQuota) { "Over Quota" } elseif ($isInactive) { "Inactive" } else { "Active" }
                    FileCount = $site.StorageUsageCurrent
                    SharingCapability = $site.SharingCapability
                    CreatedDate = $site.Template
                }
                
                $results += $obj
            }
        }
        catch {
            Write-Warning "Error processing site $($site.Url): $_"
        }
    }
    
    Write-Progress -Activity "Analyzing OneDrive Sites" -Completed
}
catch {
    Write-Host "Error retrieving OneDrive sites: $_" -ForegroundColor Red
    Disconnect-SPOService
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "OneDrive Usage Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total OneDrive Sites: $($results.Count)" -ForegroundColor White
    Write-Host "  Total Storage Allocated: $([math]::Round($totalStorageGB, 2)) GB" -ForegroundColor White
    Write-Host "  Total Storage Used: $([math]::Round($totalUsedGB, 2)) GB" -ForegroundColor White
    Write-Host "  Average Storage Per Site: $([math]::Round(($totalUsedGB / $results.Count), 2)) GB" -ForegroundColor White
    Write-Host "  Inactive Sites (>$InactiveDays days): $inactiveCount" -ForegroundColor Yellow
    Write-Host "  Sites Over Quota: $overQuotaCount" -ForegroundColor Red
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display sample results
    Write-Host "Top 10 Largest OneDrive Sites:" -ForegroundColor Yellow
    $results | Sort-Object StorageUsedGB -Descending | Select-Object -First 10 | Format-Table DisplayName, Owner, StorageUsedGB, PercentUsed, Status -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No OneDrive sites found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from services..." -ForegroundColor Cyan
Disconnect-SPOService
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
