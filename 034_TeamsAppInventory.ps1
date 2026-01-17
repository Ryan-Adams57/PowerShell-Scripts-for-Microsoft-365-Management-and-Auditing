Try { Connect-MgGraph -Scopes "AppCatalog.Read.All"; Get-MgAppCatalogTeamApp | Export-Csv "034_TeamsApps.csv" } Catch { Write-Error $_ }
