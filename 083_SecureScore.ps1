Try { Connect-MgGraph -Scopes "SecurityEvents.Read.All"; Get-MgSecuritySecureScore -Top 1 | Export-Csv "083_SecureScore.csv" } Catch { Write-Error $_ }
