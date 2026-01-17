Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -Filter "groupTypes/any(c:c eq 'DynamicMembership')" | Export-Csv "043_DynamicGroups.csv" } Catch { Write-Error $_ }
