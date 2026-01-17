Try { Connect-MgGraph -Scopes "AuditLog.Read.All"; Get-MgAuditLogDirectoryAudit -Filter "activityDisplayName eq 'Delete group'" | Export-Csv "084_DeletedGroupsAudit.csv" } Catch { Write-Error $_ }
