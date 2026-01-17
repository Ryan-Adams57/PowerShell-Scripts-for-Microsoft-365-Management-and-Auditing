Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -Filter "mailEnabled eq true" | Select-Object DisplayName, ProxyAddresses | Export-Csv "049_GroupAliases.csv" } Catch { Write-Error $_ }
