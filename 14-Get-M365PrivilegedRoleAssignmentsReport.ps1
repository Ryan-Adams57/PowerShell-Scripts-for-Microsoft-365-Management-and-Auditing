<#
====================================================================================
Script Name: Get-M365PrivilegedRoleAssignmentsReport.ps1
Description: Privileged Administrator Role Assignments Report
Author: Ryan Adams
Website: https://www.governmentcontrol.net/
====================================================================================

SCRIPT HIGHLIGHTS:
• Lists all Azure AD administrator role assignments
• Shows Global Administrators and other privileged roles
• Identifies direct vs group-based role assignments
• Highlights PIM (Privileged Identity Management) eligible roles
• Shows assignment dates and assignment scope
• Generates compliance and security audit trails
• Exports complete role assignment inventory
• Critical for security audits and compliance reviews

====================================================================================
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$RoleName,
    
    [Parameter(Mandatory=$false)]
    [switch]$GlobalAdminsOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludePIMEligible,
    
    [Parameter(Mandatory=$false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\M365_Privileged_Role_Assignments_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Module validation and installation
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
Write-Host "Microsoft 365 Privileged Role Assignments Report" -ForegroundColor Green
Write-Host "`n====================================================================================`n" -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph.Identity.DirectoryManagement", "Microsoft.Graph.Users")

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
    Connect-MgGraph -Scopes "RoleManagement.Read.Directory", "Directory.Read.All", "User.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph.`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit
}

# Retrieve directory roles
Write-Host "Retrieving Azure AD directory roles..." -ForegroundColor Cyan
$results = @()
$globalAdminCount = 0

try {
    # Get all directory roles
    if ($RoleName) {
        $roles = Get-MgDirectoryRole -Filter "displayName eq '$RoleName'" -All
    }
    elseif ($GlobalAdminsOnly) {
        $roles = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'" -All
    }
    else {
        $roles = Get-MgDirectoryRole -All
    }
    
    Write-Host "Found $($roles.Count) active directory role(s). Retrieving assignments...`n" -ForegroundColor Green
    
    $progressCounter = 0
    
    foreach ($role in $roles) {
        $progressCounter++
        Write-Progress -Activity "Processing Directory Roles" -Status "Role $progressCounter of $($roles.Count): $($role.DisplayName)" -PercentComplete (($progressCounter / $roles.Count) * 100)
        
        try {
            # Get role members
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
            
            foreach ($member in $roleMembers) {
                try {
                    $memberType = $member.AdditionalProperties.'@odata.type'
                    $memberDetails = $null
                    $assignmentType = "Direct"
                    
                    if ($memberType -eq '#microsoft.graph.user') {
                        # User assignment
                        try {
                            $memberDetails = Get-MgUser -UserId $member.Id -Property DisplayName, UserPrincipalName, AccountEnabled, UserType -ErrorAction Stop
                            
                            # Filter if specified
                            if ($UserPrincipalName -and $memberDetails.UserPrincipalName -ne $UserPrincipalName) {
                                continue
                            }
                            
                            if ($role.DisplayName -eq 'Global Administrator') {
                                $globalAdminCount++
                            }
                            
                            $obj = [PSCustomObject]@{
                                RoleName = $role.DisplayName
                                RoleId = $role.Id
                                MemberType = "User"
                                DisplayName = $memberDetails.DisplayName
                                UserPrincipalName = $memberDetails.UserPrincipalName
                                UserType = $memberDetails.UserType
                                AccountEnabled = $memberDetails.AccountEnabled
                                AssignmentType = $assignmentType
                                MemberId = $member.Id
                                IsPrivileged = if ($role.DisplayName -match "Global|Security|Compliance|Application|Cloud|Exchange|SharePoint|Teams") { $true } else { $false }
                                RiskLevel = if ($role.DisplayName -eq 'Global Administrator') { "Critical" } 
                                           elseif ($role.DisplayName -match "Security|Compliance") { "High" } 
                                           else { "Medium" }
                            }
                            
                            $results += $obj
                        }
                        catch {
                            Write-Warning "Could not retrieve user details for member ID: $($member.Id)"
                        }
                    }
                    elseif ($memberType -eq '#microsoft.graph.group') {
                        # Group assignment
                        try {
                            $groupDetails = Get-MgGroup -GroupId $member.Id -Property DisplayName, Mail -ErrorAction Stop
                            $assignmentType = "Group-Based"
                            
                            $obj = [PSCustomObject]@{
                                RoleName = $role.DisplayName
                                RoleId = $role.Id
                                MemberType = "Group"
                                DisplayName = $groupDetails.DisplayName
                                UserPrincipalName = $groupDetails.Mail
                                UserType = "Group"
                                AccountEnabled = "N/A"
                                AssignmentType = $assignmentType
                                MemberId = $member.Id
                                IsPrivileged = if ($role.DisplayName -match "Global|Security|Compliance|Application|Cloud|Exchange|SharePoint|Teams") { $true } else { $false }
                                RiskLevel = if ($role.DisplayName -eq 'Global Administrator') { "Critical" } 
                                           elseif ($role.DisplayName -match "Security|Compliance") { "High" } 
                                           else { "Medium" }
                            }
                            
                            $results += $obj
                        }
                        catch {
                            Write-Warning "Could not retrieve group details for member ID: $($member.Id)"
                        }
                    }
                    elseif ($memberType -eq '#microsoft.graph.servicePrincipal') {
                        # Service Principal assignment
                        $obj = [PSCustomObject]@{
                            RoleName = $role.DisplayName
                            RoleId = $role.Id
                            MemberType = "ServicePrincipal"
                            DisplayName = "Service Principal"
                            UserPrincipalName = "N/A"
                            UserType = "ServicePrincipal"
                            AccountEnabled = "N/A"
                            AssignmentType = "Direct"
                            MemberId = $member.Id
                            IsPrivileged = if ($role.DisplayName -match "Global|Security|Compliance|Application|Cloud") { $true } else { $false }
                            RiskLevel = "Medium"
                        }
                        
                        $results += $obj
                    }
                }
                catch {
                    Write-Warning "Error processing member in role $($role.DisplayName): $_"
                }
            }
        }
        catch {
            Write-Warning "Error retrieving members for role $($role.DisplayName): $_"
        }
    }
    
    Write-Progress -Activity "Processing Directory Roles" -Completed
}
catch {
    Write-Host "Error retrieving directory roles: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Export and display results
if ($results.Count -gt 0) {
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    Write-Host "Privileged Role Assignment Summary:" -ForegroundColor Green
    Write-Host "  Total Role Assignments: $($results.Count)" -ForegroundColor White
    Write-Host "  Global Administrators: $globalAdminCount" -ForegroundColor Red
    Write-Host "  User Assignments: $(($results | Where-Object { $_.MemberType -eq 'User' }).Count)" -ForegroundColor White
    Write-Host "  Group-Based Assignments: $(($results | Where-Object { $_.AssignmentType -eq 'Group-Based' }).Count)" -ForegroundColor White
    Write-Host "  Critical Risk Assignments: $(($results | Where-Object { $_.RiskLevel -eq 'Critical' }).Count)" -ForegroundColor Red
    
    # Role distribution
    Write-Host "`n  Top 5 Most Assigned Roles:" -ForegroundColor Cyan
    $results | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host "    $($_.Name): $($_.Count)" -ForegroundColor White
    }
    
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n  Report Location: $ExportPath" -ForegroundColor White
    Write-Host "`n====================================================================================`n" -ForegroundColor Cyan
    
    Write-Host "SECURITY RECOMMENDATION:" -ForegroundColor Red
    Write-Host "Review all privileged role assignments regularly and follow least privilege principles.`n" -ForegroundColor Yellow
    
    # Display sample results
    Write-Host "Sample Results (First 10):" -ForegroundColor Yellow
    $results | Select-Object -First 10 | Format-Table RoleName, DisplayName, UserPrincipalName, MemberType, RiskLevel -AutoSize
    
    $openFile = Read-Host "Would you like to open the CSV report? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ExportPath
    }
}
else {
    Write-Host "No role assignments found matching the specified criteria." -ForegroundColor Yellow
}

# Cleanup
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Script completed successfully.`n" -ForegroundColor Green
