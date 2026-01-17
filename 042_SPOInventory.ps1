Try { Connect-MgGraph -Scopes "Sites.Read.All"; Get-MgSite -All | Select-Object DisplayName, WebUrl | Export-Csv "042_SPOInventory.csv" } Catch { Write-Error $_ }
