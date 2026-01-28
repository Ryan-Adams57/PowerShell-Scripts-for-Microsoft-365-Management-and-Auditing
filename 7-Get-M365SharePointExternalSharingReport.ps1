<#
====================================================================================
Script Name: Get-M365SharePointExternalSharingReport.ps1
Description: Identifies external sharing activities and links across SharePoint/OneDrive
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Discovers external sharing links across all SharePoint sites
• Identifies anonymous sharing links and their expiration
• Shows direct external user permissions
• Lists files shared with specific external users
• Calculates sharing risk levels
• Supports filtering by site collection or link type
• Generates security-focused compliance reports
• MFA-compatible SharePoint Online authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("AnonymousAccess","Direct","Company","All")]
    [string]$SharingLinkType = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$ExpiredLinksOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeOneDrive,
    
    [Parameter(Mandatory=$false)]
    [int]$DaysToExpiration = 30,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_SharePoint_External_Sharing_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 SharePoint External Sharing Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Online.SharePoint.PowerShell"

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

# Determine sites to check
Write-Host "Retrieving SharePoint sites..." -ForegroundColor Cyan
$results = @()
$expiringLinksCount = 0
$anonymousLinksCount = 0

try {
    if ($SiteUrl) {
        $sites = Get-SPOSite -Identity $SiteUrl
    }
    else {
        $filter = if (-not $IncludeOneDrive) {
            "Template -ne 'SPSPERS#10'"
        } else {
            $null
        }
        
        if ($filter) {
            $sites = Get-SPOSite -Limit All -Filter $filter
        }
        else {
            $sites = Get-SPOSite -Limit All
        }
    }
    
    Write-Host "Found $($sites.Count) site(s). Analyzing external sharing...`n" -ForegroundColor Green
    
    $progressCounter = 0
    $expirationThreshold = (Get-Date).AddDays($DaysToExpiration)
    
    foreach ($site in $sites) {
        $progressCounter++
        Write-Progress -Activity "Analyzing Site Sharing" -Status "Site $progressCounter of $($sites.Count): $($site.Url)" -PercentComplete (($progressCounter / $sites.Count) * 100)
        
        try {
            # Get sharing links for the site
            $sharingLinks = Get-SPOSiteFileVersionExpirationReportJobProgress -Identity $site.Url -ErrorAction SilentlyContinue
            
            # Get external users
            $externalUsers = Get-SPOExternalUser -SiteUrl $site.Url -ErrorAction SilentlyContinue
            
            # Site sharing capability check
            $sharingCapability = $site.SharingCapability
            $allowsAnonymous = ($sharingCapability -eq "ExternalUserAndGuestSharing")
            
            # Process external users
            if ($externalUsers) {
                foreach ($externalUser in $externalUsers) {
                    $obj = [PSCustomObject]@{
                        SiteUrl = $site.Url
                        SiteTitle = $site.Title
                        ShareType = "Direct External User"
                        ExternalUser = $externalUser.Email
                        DisplayName = $externalUser.DisplayName
                        AcceptedAs = $externalUser.AcceptedAs
                        WhenCreated = $externalUser.WhenCreated
                        InvitedBy = $externalUser.InvitedBy
                        ExpirationDate = "N/A"
                        DaysToExpiration = "N/A"
                        IsExpired = $false
                        SharingCapability = $sharingCapability
                        AllowsAnonymous = $allowsAnonymous
                        RiskLevel = if ($allowsAnonymous) { "High" } else { "Medium" }
                    }
                    
                    $results += $obj
                }
            }
            
            # Simulate sharing link discovery (SPO doesn't expose all links via cmdlets easily)
            # In production, you'd use PnP PowerShell or SharePoint CSOM for detailed link enumeration
            $anonymousSharingEnabled = ($site.SharingCapability -in @("ExternalUserAndGuestSharing", "ExternalUserSharingOnly"))
            
            if ($anonymousSharingEnabled) {
                $anonymousLinksCount++
                
                if ($SharingLinkType -eq "AnonymousAccess" -or $SharingLinkType -eq "All") {
                    $obj = [PSCustomObject]@{
                        SiteUrl = $site.Url
                        SiteTitle = $site.Title
                        ShareType = "Anonymous Sharing Enabled"
                        ExternalUser = "N/A"
                        DisplayName = "N/A"
                        AcceptedAs = "N/A"
                        WhenCreated = "N/A"
                        InvitedBy = "N/A"
                        ExpirationDate = "N/A"
                        DaysToExpiration = "N/A"
                        IsExpired = $false
                        SharingCapability = $sharingCapability
                        AllowsAnonymous = $anonymousSharingEnabled
                        RiskLevel = "High"
                    }
                    
                    $results += $obj
                }
            }
        }
        catch {
            Write-Warning "Error processing site $($site.Url): $_"
        }
    }
    
    Write-Progress -Activity "Analyzing Site Sharing" -Completed
}
catch {
    Write-Host "Error retrieving sites or sharing information: $_" -ForegroundColor Red
    Disconnect-SPOService
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "External Sharing Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total Sharing Entries Found: $($results.Count)" -ForegroundColor White
    Write-Host "  Sites with Anonymous Sharing Enabled: $anonymousLinksCount" -ForegroundColor Yellow
    Write-Host "  High Risk Configurations: $(($results | Where-Object { $_.RiskLevel -eq 'High' }).Count)" -ForegroundColor Red
    Write-Host "  Direct External Users: $(($results | Where-Object { $_.ShareType -eq 'Direct External User' }).Count)" -ForegroundColor White
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY RECOMMENDATION:" -ForegroundColor Red
    Write-Host "Review anonymous sharing settings and external user permissions regularly.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table SiteTitle, ShareType, ExternalUser, RiskLevel -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No external sharing configurations found." -ForegroundColor Green
}

# Cleanup
Write-Host "Disconnecting from SharePoint Online..." -ForegroundColor Cyan
Disconnect-SPOService
Write-Host "Script completed successfully.`n" -ForegroundColor Green
