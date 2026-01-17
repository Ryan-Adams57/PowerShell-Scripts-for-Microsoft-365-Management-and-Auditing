Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgDirectoryDeletedItemAsGroup | Export-Csv "040_DeletedGroups.csv" } Catch { Write-Error $_ }
