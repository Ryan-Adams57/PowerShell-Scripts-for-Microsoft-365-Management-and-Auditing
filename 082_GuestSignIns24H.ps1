Try { Connect-MgGraph -Scopes "AuditLog.Read.All"; Get-MgAuditLogSignIn -Filter "userType eq 'Guest'" | Export-Csv "082_GuestLogins.csv" } Catch { Write-Error $_ }
