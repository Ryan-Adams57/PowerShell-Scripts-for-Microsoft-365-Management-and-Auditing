Try { Connect-MgGraph -Scopes "Policy.Read.All"; Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy | Export-Csv "024_SecurityDefaults.csv" } Catch { Write-Error $_ }
