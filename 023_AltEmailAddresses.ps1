Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All | Select-Object DisplayName, ProxyAddresses | Export-Csv "023_AltEmails.csv" } Catch { Write-Error $_ }
