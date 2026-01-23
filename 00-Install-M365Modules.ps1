<#
====================================================================================
Script Name: 00-Install-M365Modules.ps1
Description: Master module installer for all 75 M365 PowerShell scripts
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Installs ALL required modules for the complete 75-script collection
• Checks for existing modules and versions
• Updates outdated modules automatically
• Supports selective installation by category
• Validates successful installation
• Provides installation summary and troubleshooting tips
• One-time setup for entire script collection
• Saves time by batch-installing all dependencies

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("All","Core","Security","Compliance","Teams","Exchange","SharePoint","PowerPlatform","Intune","Analytics")]
    [string]$Category = "All",
    
    [Parameter(Mandatory=$false)]
    [switch]$UpdateExisting,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipPrompts
)

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "M365 PowerShell Modules - Master Installer" -ForegroundColor Green
Write-Host "Installing dependencies for all 75 M365 scripts" -ForegroundColor White
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

# Determine which modules to install based on category
$modulesToInstall = @()

if ($Category -eq "All") {
    foreach ($cat in $moduleCategories.Keys) {
        $modulesToInstall += $moduleCategories[$cat]
    }
    # Remove duplicates
    $modulesToInstall = $modulesToInstall | Select-Object -Unique
}
else {
    $modulesToInstall = $moduleCategories[$Category]
}

Write-Host "Installation Category: $Category" -ForegroundColor Cyan
Write-Host "Total Modules to Process: $($modulesToInstall.Count)`n" -ForegroundColor White

if (-not $SkipPrompts) {
    Write-Host "This will install the following modules:" -ForegroundColor Yellow
    $modulesToInstall | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
    Write-Host ""
    
    $confirm = Read-Host "Do you want to continue? (Y/N)"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "Installation cancelled.`n" -ForegroundColor Yellow
        exit
    }
}

Write-Host "`nStarting module installation...`n" -ForegroundColor Cyan

$successCount = 0
$failureCount = 0
$skippedCount = 0
$updatedCount = 0
$installResults = @()

$progressCounter = 0

foreach ($moduleName in $modulesToInstall) {
    $progressCounter++
    Write-Progress -Activity "Installing M365 Modules" -Status "Module $progressCounter of $($modulesToInstall.Count): $moduleName" -PercentComplete (($progressCounter / $modulesToInstall.Count) * 100)
    
    Write-Host "Processing: $moduleName" -ForegroundColor Cyan
    
    try {
        # Check if module already exists
        $existingModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        
        if ($existingModule) {
            Write-Host "  Found existing version: $($existingModule.Version)" -ForegroundColor Yellow
            
            if ($UpdateExisting) {
                Write-Host "  Updating module..." -ForegroundColor Cyan
                try {
                    Update-Module -Name $moduleName -Force -ErrorAction Stop
                    Write-Host "  ✓ Updated successfully" -ForegroundColor Green
                    $updatedCount++
                    $installResults += [PSCustomObject]@{
                        Module = $moduleName
                        Status = "Updated"
                        Version = (Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1).Version
                        Result = "Success"
                    }
                }
                catch {
                    Write-Host "  ⚠ Update failed, keeping existing version" -ForegroundColor Yellow
                    $skippedCount++
                    $installResults += [PSCustomObject]@{
                        Module = $moduleName
                        Status = "Skipped"
                        Version = $existingModule.Version
                        Result = "Update Failed"
                    }
                }
            }
            else {
                Write-Host "  ✓ Already installed (use -UpdateExisting to update)" -ForegroundColor Green
                $skippedCount++
                $installResults += [PSCustomObject]@{
                    Module = $moduleName
                    Status = "Skipped"
                    Version = $existingModule.Version
                    Result = "Already Installed"
                }
            }
        }
        else {
            Write-Host "  Installing module..." -ForegroundColor Cyan
            Install-Module -Name $moduleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            
            $installedVersion = (Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-Host "  ✓ Installed successfully (v$installedVersion)" -ForegroundColor Green
            $successCount++
            $installResults += [PSCustomObject]@{
                Module = $moduleName
                Status = "Installed"
                Version = $installedVersion
                Result = "Success"
            }
        }
    }
    catch {
        Write-Host "  ✗ Installation failed: $_" -ForegroundColor Red
        $failureCount++
        $installResults += [PSCustomObject]@{
            Module = $moduleName
            Status = "Failed"
            Version = "N/A"
            Result = $_.Exception.Message
        }
    }
    
    Write-Host ""
}

Write-Progress -Activity "Installing M365 Modules" -Completed

# Display summary
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Installation Summary:" -ForegroundColor Green
Write-Host "  Total Modules Processed: $($modulesToInstall.Count)" -ForegroundColor White
Write-Host "  Newly Installed: $successCount" -ForegroundColor Green
Write-Host "  Updated: $updatedCount" -ForegroundColor Cyan
Write-Host "  Already Installed (Skipped): $skippedCount" -ForegroundColor Yellow
Write-Host "  Failed: $failureCount" -ForegroundColor Red
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

# Display detailed results
if ($installResults.Count -gt 0) {
    Write-Host "Detailed Results:" -ForegroundColor Yellow
    $installResults | Format-Table Module, Status, Version, Result -AutoSize
}

# Display failures if any
if ($failureCount -gt 0) {
    Write-Host "`nFailed Installations:" -ForegroundColor Red
    $installResults | Where-Object { $_.Status -eq "Failed" } | ForEach-Object {
        Write-Host "  ✗ $($_.Module)" -ForegroundColor Red
        Write-Host "    Error: $($_.Result)" -ForegroundColor Yellow
    }
}

# Provide next steps
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Next Steps:" -ForegroundColor Green
Write-Host "  1. Modules are now installed and ready to use" -ForegroundColor White
Write-Host "  2. Run any of the 75 M365 scripts without module installation prompts" -ForegroundColor White
Write-Host "  3. If you need to update modules later, run with -UpdateExisting switch" -ForegroundColor White
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
Write-Host "  SharePoint   - SharePoint Online & PnP" -ForegroundColor White
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
Write-Host "  • For corporate environments, check with IT about proxy settings" -ForegroundColor White

Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

if ($failureCount -eq 0) {
    Write-Host "✓ All modules installed successfully!" -ForegroundColor Green
    Write-Host "You're ready to run all 75 M365 PowerShell scripts!`n" -ForegroundColor Green
}
else {
    Write-Host "⚠ Some modules failed to install. Review errors above.`n" -ForegroundColor Yellow
}
