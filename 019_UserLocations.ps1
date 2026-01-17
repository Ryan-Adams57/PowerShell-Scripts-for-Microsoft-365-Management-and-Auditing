Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All | Select-Object DisplayName, UsageLocation | Export-Csv "019_UserLocations.csv" } Catch { Write-Error $_ }
