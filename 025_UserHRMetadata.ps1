Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All | Select-Object DisplayName, Department, JobTitle | Export-Csv "025_UserHRMetadata.csv" } Catch { Write-Error $_ }
