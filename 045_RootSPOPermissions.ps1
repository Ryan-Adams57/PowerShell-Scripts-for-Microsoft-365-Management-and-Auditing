Try { Connect-MgGraph -Scopes "Sites.Read.All"; Get-MgSitePermission -SiteId "root" | Export-Csv "045_RootSPOPermissions.csv" } Catch { Write-Error $_ }
