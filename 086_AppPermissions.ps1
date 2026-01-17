Try { Connect-MgGraph -Scopes "Application.Read.All"; Get-MgServicePrincipal -All | Export-Csv "086_AppPermissions.csv" } Catch { Write-Error $_ }
