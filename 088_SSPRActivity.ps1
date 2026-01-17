Try { Connect-MgGraph -Scopes "AuditLog.Read.All"; Get-MgAuditLogDirectoryAudit -Filter "category eq 'SelfServicePasswordManagement'" | Export-Csv "088_SSPR.csv" } Catch { Write-Error $_ }
