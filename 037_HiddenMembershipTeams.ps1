Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -Filter "visibility eq 'HiddenMembership'" | Export-Csv "037_HiddenTeams.csv" } Catch { Write-Error $_ }
