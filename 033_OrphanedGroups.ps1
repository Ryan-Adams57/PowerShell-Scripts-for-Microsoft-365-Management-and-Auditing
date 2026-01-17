Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -All | ForEach-Object { if(!(Get-MgGroupOwner -GroupId $_.Id)){$_} } | Export-Csv "033_OrphanedGroups.csv" } Catch { Write-Error $_ }
