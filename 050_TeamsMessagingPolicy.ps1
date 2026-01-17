Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All | Select-Object UserPrincipalName, ExternalUserState | Export-Csv "050_TeamsUserAudit.csv" } Catch { Write-Error $_ }
