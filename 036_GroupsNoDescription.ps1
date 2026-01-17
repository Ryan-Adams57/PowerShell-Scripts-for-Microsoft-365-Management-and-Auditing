Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -Filter "description eq null" | Export-Csv "036_GroupsNoDescription.csv" } Catch { Write-Error $_ }
