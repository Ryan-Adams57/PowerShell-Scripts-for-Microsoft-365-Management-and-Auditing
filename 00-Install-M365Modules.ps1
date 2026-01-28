<#
====================================================================================
Script Name: 00-Install-M365Modules.ps1
Description: Master module installer for all M365 PowerShell scripts
Version: 2.0 - Production Ready
Last Updated: 2026-01-28
====================================================================================

SCRIPT HIGHLIGHTS:
• Installs ALL required modules for the complete M365 script collection
• Checks for existing modules and versions
• Updates outdated modules automatically
• Supports selective installation by category
• Validates successful installation
• Provides installation summary and troubleshooting tips
• One-time setup for entire script collection
• Comprehensive error handling with try/catch/finally
• Progress indicators for all operations
• No API calls in loops (optimized module checking)

REQUIREMENTS:
• PowerShell 5.1 or higher
• Internet connectivity
• Administrator rights (recommended) or CurrentUser scope

====================================================================================
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("All","Core","Security","Compliance","Teams","Exchange","SharePoint","PowerPlatform","Intune","Analytics")]
    [string]$Category = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$UpdateExisting,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipPrompts,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize script variables
$script:SuccessCount = 0
$script:FailureCount = 0
$script:SkippedCount = 0
$script:UpdatedCount = 0
$script:InstallResults = @()

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "M365 PowerShell Modules - Master Installer" -ForegroundColor Green
Write-Host "Installing dependencies for M365 script collection" -ForegroundColor White
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

# Define all required modules organized by category
$moduleCategories = @{
    Core = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Reports",
        "Microsoft.Graph.Identity.DirectoryManagement"
    )
    
    Security = @(
        "Microsoft.Graph.Identity.SignIns",
        "Microsoft.Graph.Security",
        "Microsoft.Graph.Identity.Governance",
        "Microsoft.Graph.DeviceManagement"
    )
    
    Compliance = @(
        "Microsoft.Graph.Compliance",
        "ExchangeOnlineManagement"
    )
    
    Teams = @(
        "MicrosoftTeams"
    )
    
    Exchange = @(
        "ExchangeOnlineManagement"
    )
    
    SharePoint = @(
        "Microsoft.Online.SharePoint.PowerShell",
        "PnP.PowerShell"
    )
    
    PowerPlatform = @(
        "Microsoft.PowerApps.Administration.PowerShell",
        "Microsoft.PowerApps.PowerShell",
        "MicrosoftPowerBIMgmt"
    )
    
    Intune = @(
        "Microsoft.Graph.DeviceManagement",
        "Microsoft.Graph.DeviceManagement.Enrolment"
    )
    
    Analytics = @(
        "Microsoft.Graph.Reports"
    )
}

# Function to install or update a module
function Install-M365Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory=$false)]
        [bool]$UpdateIfExists = $false
    )
    
    try {
        $existingModule = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue | 
            Sort-Object Version -Descending | 
            Select-Object -First 1
        
        if ($existingModule) {
            Write-Host "  Found existing version: $($existingModule.Version)" -ForegroundColor Yellow
            
            if ($UpdateIfExists) {
                Write-Host "  Updating module..." -ForegroundColor Cyan
                try {
                    Update-Module -Name $ModuleName -Force -ErrorAction Stop
                    $newVersion = (Get-Module -ListAvailable -Name $ModuleName -ErrorAction Stop | 
                        Sort-Object Version -Descending | 
                        Select-Object -First 1).Version
                    Write-Host "  ✓ Updated successfully to version $newVersion" -ForegroundColor Green
                    $script:UpdatedCount++
                    
                    return [PSCustomObject]@{
                        Module = $ModuleName
                        Status = "Updated"
                        Version = $newVersion
                        Result = "Success"
                    }
                }
                catch {
                    Write-Host "  ⚠ Update failed, keeping existing version" -ForegroundColor Yellow
                    $script:SkippedCount++
                    
                    return [PSCustomObject]@{
                        Module = $ModuleName
                        Status = "Skipped"
                        Version = $existingModule.Version
                        Result = "Update Failed: $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-Host "  ✓ Already installed (use -UpdateExisting to update)" -ForegroundColor Green
                $script:SkippedCount++
                
                return [PSCustomObject]@{
                    Module = $ModuleName
                    Status = "Skipped"
                    Version = $existingModule.Version
                    Result = "Already Installed"
                }
            }
        }
        else {
            Write-Host "  Installing module..." -ForegroundColor Cyan
            Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            
            $installedVersion = (Get-Module -ListAvailable -Name $ModuleName -ErrorAction Stop | 
                Sort-Object Version -Descending | 
                Select-Object -First 1).Version
            Write-Host "  ✓ Installed successfully (v$installedVersion)" -ForegroundColor Green
            $script:SuccessCount++
            
            return [PSCustomObject]@{
                Module = $ModuleName
                Status = "Installed"
                Version = $installedVersion
                Result = "Success"
            }
        }
    }
    catch {
        Write-Host "  ✗ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
        $script:FailureCount++
        
        return [PSCustomObject]@{
            Module = $ModuleName
            Status = "Failed"
            Version = "N/A"
            Result = $_.Exception.Message
        }
    }
}

# Determine which modules to install based on category
$modulesToInstall = @()

if ($Category -eq "All") {
    foreach ($cat in $moduleCategories.Keys) {
        $modulesToInstall += $moduleCategories[$cat]
    }
    $modulesToInstall = $modulesToInstall | Select-Object -Unique
}
else {
    if ($moduleCategories.ContainsKey($Category)) {
        $modulesToInstall = $moduleCategories[$Category]
    }
    else {
        Write-Host "ERROR: Invalid category '$Category'" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Installation Category: $Category" -ForegroundColor Cyan
Write-Host "Total Modules to Process: $($modulesToInstall.Count)`n" -ForegroundColor White

if (-not $SkipPrompts -and -not $Force) {
    Write-Host "This will install the following modules:" -ForegroundColor Yellow
    $modulesToInstall | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
    Write-Host ""
    
    $confirm = Read-Host "Do you want to continue? (Y/N)"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "Installation cancelled.`n" -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "`nStarting module installation...`n" -ForegroundColor Cyan

try {
    $progressCounter = 0
    
    foreach ($moduleName in $modulesToInstall) {
        $progressCounter++
        Write-Progress -Activity "Installing M365 Modules" `
            -Status "Module $progressCounter of $($modulesToInstall.Count): $moduleName" `
            -PercentComplete (($progressCounter / $modulesToInstall.Count) * 100)
        
        Write-Host "Processing: $moduleName" -ForegroundColor Cyan
        
        $result = Install-M365Module -ModuleName $moduleName -UpdateIfExists $UpdateExisting.IsPresent
        $script:InstallResults += $result
        
        Write-Host ""
    }
    
    Write-Progress -Activity "Installing M365 Modules" -Completed
}
catch {
    Write-Host "`nCRITICAL ERROR during installation: $($_.Exception.Message)" -ForegroundColor Red
    Write-Progress -Activity "Installing M365 Modules" -Completed
}
finally {
    # Display summary
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Installation Summary:" -ForegroundColor Green
    Write-Host "  Total Modules Processed: $($modulesToInstall.Count)" -ForegroundColor White
    Write-Host "  Newly Installed: $script:SuccessCount" -ForegroundColor Green
    Write-Host "  Updated: $script:UpdatedCount" -ForegroundColor Cyan
    Write-Host "  Already Installed (Skipped): $script:SkippedCount" -ForegroundColor Yellow
    Write-Host "  Failed: $script:FailureCount" -ForegroundColor Red
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    # Display detailed results
    if ($script:InstallResults.Count -gt 0) {
        Write-Host "Detailed Results:" -ForegroundColor Yellow
        $script:InstallResults | Format-Table Module, Status, Version, Result -AutoSize
    }
    
    # Display failures if any
    if ($script:FailureCount -gt 0) {
        Write-Host "`nFailed Installations:" -ForegroundColor Red
        $script:InstallResults | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
            Write-Host "  ✗ $($_.Module)" -ForegroundColor Red
            Write-Host "    Error: $($_.Result)" -ForegroundColor Yellow
        }
    }
    
    # Provide next steps
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Next Steps:" -ForegroundColor Green
    Write-Host "  1. Modules are now installed and ready to use" -ForegroundColor White
    Write-Host "  2. Run any M365 scripts without module installation prompts" -ForegroundColor White
    Write-Host "  3. To update modules later, run with -UpdateExisting switch" -ForegroundColor White
    Write-Host "`n  Example: .\00-Install-M365Modules.ps1 -UpdateExisting" -ForegroundColor Cyan
    
    # Module categories reference
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Installation by Category:" -ForegroundColor Green
    Write-Host "  All          - Install all modules (default)" -ForegroundColor White
    Write-Host "  Core         - Microsoft Graph core modules" -ForegroundColor White
    Write-Host "  Security     - Security and identity protection" -ForegroundColor White
    Write-Host "  Compliance   - Compliance and governance" -ForegroundColor White
    Write-Host "  Teams        - Microsoft Teams" -ForegroundColor White
    Write-Host "  Exchange     - Exchange Online" -ForegroundColor White
    Write-Host "  SharePoint   - SharePoint Online and PnP" -ForegroundColor White
    Write-Host "  PowerPlatform- Power Apps, Power Automate, Power BI" -ForegroundColor White
    Write-Host "  Intune       - Intune device management" -ForegroundColor White
    Write-Host "  Analytics    - Usage analytics and reporting" -ForegroundColor White
    Write-Host "`n  Example: .\00-Install-M365Modules.ps1 -Category Security" -ForegroundColor Cyan
    
    # Troubleshooting tips
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  • If installation fails, run PowerShell as Administrator" -ForegroundColor White
    Write-Host "  • Ensure you have internet connectivity" -ForegroundColor White
    Write-Host "  • Check PowerShell execution policy: Get-ExecutionPolicy" -ForegroundColor White
    Write-Host "  • Set execution policy if needed: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor White
    Write-Host "  • For corporate environments, check proxy settings with IT" -ForegroundColor White
    
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($script:FailureCount -eq 0) {
        Write-Host "✓ All modules installed successfully!" -ForegroundColor Green
        Write-Host "You're ready to run all M365 PowerShell scripts!`n" -ForegroundColor Green
        exit 0
    }
    else {
        Write-Host "⚠ Some modules failed to install. Review errors above.`n" -ForegroundColor Yellow
        exit 1
    }
}
