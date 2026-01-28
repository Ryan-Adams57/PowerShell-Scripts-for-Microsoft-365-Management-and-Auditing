<#
====================================================================================
Script Name: Get-AzureADIdentityProtectionReport.ps1
Description: Azure AD Identity Protection policies and risk detections report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Retrieves Azure AD Identity Protection policies
• Shows user risk and sign-in risk policies
• Lists risky users and risky sign-ins
• Identifies risk detections and event types
• Tracks policy enforcement and remediation
• Supports filtering by risk level
• Generates security monitoring and incident reports
• Requires Azure AD Premium P2 licensing

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Low","Medium","High","All")]
    [string]$RiskLevel = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeRiskyUsers,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeRiskDetections,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowPolicies,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\AzureAD_Identity_Protection_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Azure AD Identity Protection Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModule = "Microsoft.Graph.Identity.SignIns"

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

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "IdentityRiskEvent.Read.All", "IdentityRiskyUser.Read.All", "Policy.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Initialize results
$results = @()
$riskyUserCount = 0
$riskDetectionCount = 0

# Retrieve Identity Protection policies
if ($ShowPolicies) {
    Write-Host "Retrieving Identity Protection policies..." -ForegroundColor Cyan
    
    try {
        # Get user risk policy
        $userRiskPolicy = Get-MgIdentityConditionalAccessPolicy -Filter "displayName eq 'Identity Protection - User risk policy'" -ErrorAction SilentlyContinue
        
        if ($userRiskPolicy) {
            $obj = [PSCustomObject]@{
                Type = "User Risk Policy"
                Name = $userRiskPolicy.DisplayName
                State = $userRiskPolicy.State
                RiskLevel = "N/A"
                UserPrincipalName = "N/A"
                DetectionType = "N/A"
                RiskState = "N/A"
                Details = "Policy controls user risk-based access"
                CreatedDateTime = $userRiskPolicy.CreatedDateTime
                ModifiedDateTime = $userRiskPolicy.ModifiedDateTime
            }
            $results += $obj
        }
        
        # Get sign-in risk policy
        $signInRiskPolicy = Get-MgIdentityConditionalAccessPolicy -Filter "displayName eq 'Identity Protection - Sign-in risk policy'" -ErrorAction SilentlyContinue
        
        if ($signInRiskPolicy) {
            $obj = [PSCustomObject]@{
                Type = "Sign-in Risk Policy"
                Name = $signInRiskPolicy.DisplayName
                State = $signInRiskPolicy.State
                RiskLevel = "N/A"
                UserPrincipalName = "N/A"
                DetectionType = "N/A"
                RiskState = "N/A"
                Details = "Policy controls sign-in risk-based access"
                CreatedDateTime = $signInRiskPolicy.CreatedDateTime
                ModifiedDateTime = $signInRiskPolicy.ModifiedDateTime
            }
            $results += $obj
        }
        
        Write-Host "Retrieved Identity Protection policies.`n" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not retrieve Identity Protection policies"
    }
}

# Retrieve risky users
if ($IncludeRiskyUsers) {
    Write-Host "Retrieving risky users..." -ForegroundColor Cyan
    
    try {
        $riskyUsers = Get-MgRiskyUser -All -ErrorAction Stop
        $riskyUserCount = $riskyUsers.Count
        
        Write-Host "Found $riskyUserCount risky user(s).`n" -ForegroundColor Yellow
        
        foreach ($user in $riskyUsers) {
            # Filter by risk level
            if ($RiskLevel -ne "All" -and $user.RiskLevel -ne $RiskLevel.ToLower()) {
                continue
            }
            
            $obj = [PSCustomObject]@{
                Type = "Risky User"
                Name = $user.UserDisplayName
                State = $user.RiskState
                RiskLevel = $user.RiskLevel
                UserPrincipalName = $user.UserPrincipalName
                DetectionType = "N/A"
                RiskState = $user.RiskState
                Details = "User flagged with risk: $($user.RiskDetail)"
                RiskLastUpdatedDateTime = $user.RiskLastUpdatedDateTime
                RiskDetail = $user.RiskDetail
            }
            $results += $obj
        }
    }
    catch {
        Write-Host "Error retrieving risky users: $_" -ForegroundColor Red
        Write-Host "Note: This feature requires Azure AD Premium P2 licensing.`n" -ForegroundColor Yellow
    }
}

# Retrieve risk detections
if ($IncludeRiskDetections) {
    Write-Host "Retrieving risk detections..." -ForegroundColor Cyan
    
    try {
        $riskDetections = Get-MgRiskDetection -All -ErrorAction Stop
        $riskDetectionCount = $riskDetections.Count
        
        Write-Host "Found $riskDetectionCount risk detection(s).`n" -ForegroundColor Yellow
        
        $progressCounter = 0
        
        foreach ($detection in $riskDetections) {
            $progressCounter++
            Write-Progress -Activity "Processing Risk Detections" -Status "Detection $progressCounter of $riskDetectionCount" -PercentComplete (($progressCounter / $riskDetectionCount) * 100)
            
            # Filter by risk level
            if ($RiskLevel -ne "All" -and $detection.RiskLevel -ne $RiskLevel.ToLower()) {
                continue
            }
            
            $obj = [PSCustomObject]@{
                Type = "Risk Detection"
                Name = $detection.RiskEventType
                State = $detection.RiskState
                RiskLevel = $detection.RiskLevel
                UserPrincipalName = $detection.UserPrincipalName
                DetectionType = $detection.DetectionTimingType
                RiskState = $detection.RiskState
                Details = $detection.AdditionalInfo
                DetectedDateTime = $detection.DetectedDateTime
                Location = $detection.Location.City
                IPAddress = $detection.IpAddress
                Source = $detection.Source
            }
            $results += $obj
        }
        
        Write-Progress -Activity "Processing Risk Detections" -Completed
    }
    catch {
        Write-Host "Error retrieving risk detections: $_" -ForegroundColor Red
        Write-Host "Note: This feature requires Azure AD Premium P2 licensing.`n" -ForegroundColor Yellow
    }
}

# If no specific options selected, get summary
if (-not $ShowPolicies -and -not $IncludeRiskyUsers -and -not $IncludeRiskDetections) {
    Write-Host "Retrieving Identity Protection summary..." -ForegroundColor Cyan
    
    try {
        # Try to get risky users count
        $riskyUsers = Get-MgRiskyUser -All -ErrorAction SilentlyContinue
        if ($riskyUsers) {
            $riskyUserCount = $riskyUsers.Count
        }
        
        # Try to get risk detections count
        $riskDetections = Get-MgRiskDetection -All -ErrorAction SilentlyContinue
        if ($riskDetections) {
            $riskDetectionCount = $riskDetections.Count
        }
        
        $obj = [PSCustomObject]@{
            Type = "Summary"
            Name = "Identity Protection Overview"
            State = "Active"
            RiskLevel = "N/A"
            UserPrincipalName = "N/A"
            DetectionType = "N/A"
            RiskState = "N/A"
            Details = "Risky Users: $riskyUserCount | Risk Detections: $riskDetectionCount"
            TotalRiskyUsers = $riskyUserCount
            TotalRiskDetections = $riskDetectionCount
        }
        $results += $obj
    }
    catch {
        Write-Warning "Could not retrieve Identity Protection summary"
    }
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Identity Protection Summary:" -ForegroundColor Green
    Write-Host "  Total Records: $($results.Count)" -ForegroundColor White
    
    if ($IncludeRiskyUsers) {
        Write-Host "  Total Risky Users: $riskyUserCount" -ForegroundColor Red
        Write-Host "    High Risk: $(($results | Where-Object { $_.Type -eq 'Risky User' -and $_.RiskLevel -eq 'high' }).Count)" -ForegroundColor Red
        Write-Host "    Medium Risk: $(($results | Where-Object { $_.Type -eq 'Risky User' -and $_.RiskLevel -eq 'medium' }).Count)" -ForegroundColor Yellow
        Write-Host "    Low Risk: $(($results | Where-Object { $_.Type -eq 'Risky User' -and $_.RiskLevel -eq 'low' }).Count)" -ForegroundColor Green
    }
    
    if ($IncludeRiskDetections) {
        Write-Host "  Total Risk Detections: $riskDetectionCount" -ForegroundColor Yellow
    }
    
    # Export to CSV
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY ALERT:" -ForegroundColor Red
    Write-Host "Review risky users and detections immediately. Investigate suspicious activity.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table Type, Name, RiskLevel, State, UserPrincipalName -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No Identity Protection data found." -ForegroundColor Yellow
    Write-Host "Note: This feature requires Azure AD Premium P2 licensing.`n" -ForegroundColor Cyan
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
