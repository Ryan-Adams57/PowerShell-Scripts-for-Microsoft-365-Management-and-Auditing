Try { Connect-MgGraph -Scopes "IdentityRiskyUser.Read.All"; Get-MgIdentityRiskyUser | Export-Csv "080_RiskyUsers.csv" } Catch { Write-Error $_ }
