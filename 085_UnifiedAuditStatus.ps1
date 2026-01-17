Try { Connect-ExchangeOnline; Get-AdminAuditLogConfig | Export-Csv "085_AuditLogStatus.csv" } Catch { Write-Error $_ }
