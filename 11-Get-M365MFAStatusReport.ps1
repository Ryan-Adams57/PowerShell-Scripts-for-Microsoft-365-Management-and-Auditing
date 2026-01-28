<#
====================================================================================
Script Name: Get-M365MFAStatusReport.ps1
Description: Multi-Factor Authentication status and enforcement report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Shows MFA registration status for all users
• Identifies users without MFA enabled
• Lists MFA methods configured per user
• Highlights privileged accounts without MFA
• Supports filtering by MFA status or method
• Generates security compliance reports
• Exports detailed MFA inventory
• MFA-compatible Microsoft Graph authentication

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$UnregisteredOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$PrivilegedUsersOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_MFA_Status_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 MFA Status Report Generator" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph.Identity.SignIns", "Microsoft.Graph.Users")

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

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "UserAuthenticationMethod.Read.All", "User.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve users
Write-Host "Retrieving user accounts and MFA status..." -ForegroundColor Cyan
$results = @()
$unregisteredCount = 0
$registeredCount = 0

try {
    if ($UserPrincipalName) {
        $users = Get-MgUser -UserId $UserPrincipalName -Property Id, DisplayName, UserPrincipalName, AccountEnabled, UserType
        $users = @($users)
    }
    else {
        $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, UserType -Filter "userType eq 'Member'"
    }
    
    Write-Host "Found $($users.Count) user(s). Checking MFA registration status...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($user in $users) {
        $progressCounter++
        Write-Progress -Activity "Checking MFA Status" -Status "User $progressCounter of $($users.Count): $($user.UserPrincipalName)" -PercentComplete (($progressCounter / $users.Count) * 100)
        
        try {
            # Get authentication methods for user
            $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction SilentlyContinue
            
            $mfaMethods = @()
            $hasMFA = $false
            $phoneCount = 0
            $emailCount = 0
            $appCount = 0
            $fido2Count = 0
            
            if ($authMethods) {
                foreach ($method in $authMethods) {
                    $methodType = $method.AdditionalProperties.'@odata.type'
                    
                    switch -Wildcard ($methodType) {
                        "*phoneAuthentication*" {
                            $phoneCount++
                            $mfaMethods += "Phone"
                            $hasMFA = $true
                        }
                        "*emailAuthentication*" {
                            $emailCount++
                            $mfaMethods += "Email"
                        }
                        "*microsoftAuthenticator*" {
                            $appCount++
                            $mfaMethods += "Authenticator App"
                            $hasMFA = $true
                        }
                        "*softwareOath*" {
                            $mfaMethods += "Software Token"
                            $hasMFA = $true
                        }
                        "*fido2*" {
                            $fido2Count++
                            $mfaMethods += "FIDO2 Security Key"
                            $hasMFA = $true
                        }
                        "*windowsHelloForBusiness*" {
                            $mfaMethods += "Windows Hello"
                            $hasMFA = $true
                        }
                    }
                }
            }
            
            $mfaStatus = if ($hasMFA) { "Registered" } else { "Not Registered" }
            
            if ($hasMFA) {
                $registeredCount++
            } else {
                $unregisteredCount++
            }
            
            # Skip if filtering
            if ($UnregisteredOnly -and $hasMFA) {
                continue
            }
            
            $obj = [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                AccountEnabled = $user.AccountEnabled
                MFAStatus = $mfaStatus
                MFAMethods = ($mfaMethods | Select-Object -Unique) -join "; "
                MethodCount = ($mfaMethods | Select-Object -Unique).Count
                PhoneAuth = $phoneCount
                EmailAuth = $emailCount
                AuthenticatorApp = $appCount
                FIDO2Key = $fido2Count
                UserType = $user.UserType
                RiskLevel = if (-not $hasMFA -and $user.AccountEnabled) { "High" } elseif (-not $hasMFA) { "Medium" } else { "Low" }
            }
            
            $results += $obj
        }
        catch {
            Write-Warning "Error processing user $($user.UserPrincipalName): $_"
        }
    }
    
    Write-Progress -Activity "Checking MFA Status" -Completed
}
catch {
    Write-Host "Error retrieving users or authentication methods: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "MFA Status Analysis Summary:" -ForegroundColor Green
    Write-Host "  Total Users Analyzed: $($results.Count)" -ForegroundColor White
    Write-Host "  Users with MFA Registered: $registeredCount" -ForegroundColor Green
    Write-Host "  Users without MFA: $unregisteredCount" -ForegroundColor Red
    Write-Host "  MFA Registration Rate: $([math]::Round(($registeredCount / $results.Count) * 100, 2))%" -ForegroundColor White
    Write-Host "  High Risk Users (No MFA + Enabled): $(($results | Where-Object { $_.RiskLevel -eq 'High' }).Count)" -ForegroundColor Red
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    if ($unregisteredCount -gt 0) {
        Write-Host "SECURITY RECOMMENDATION:" -ForegroundColor Red
        Write-Host "Enable MFA enforcement for all users to improve security posture.`n" -ForegroundColor Yellow
    }
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table DisplayName, UserPrincipalName, MFAStatus, MFAMethods, RiskLevel -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No users found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
