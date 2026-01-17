Try { Connect-MgGraph -Scopes "AuditLog.Read.All"; Get-MgAuditLogSignIn -Filter "userDisplayName eq 'Global Administrator'" | Export-Csv "096_AdminLogins.csv" } Catch { Write-Error $_ }
