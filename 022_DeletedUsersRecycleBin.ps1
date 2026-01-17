Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgDirectoryDeletedItemAsUser | Export-Csv "022_DeletedUsers.csv" } Catch { Write-Error $_ }
