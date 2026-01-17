Try { Connect-MgGraph -Scopes "Sites.Read.All"; Get-MgSite -All | Export-Csv "035_SPOAdmins.csv" } Catch { Write-Error $_ }
