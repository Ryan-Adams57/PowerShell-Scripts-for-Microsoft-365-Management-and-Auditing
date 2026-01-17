Try { Connect-MgGraph -Scopes "Sites.Read.All"; Get-MgSite -All | Select-Object DisplayName, WebUrl | Export-Csv "031_SPExternalSharing.csv" } Catch { Write-Error $_ }
