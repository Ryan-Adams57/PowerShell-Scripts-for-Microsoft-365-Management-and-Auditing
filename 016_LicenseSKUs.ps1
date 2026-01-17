Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All | Select-Object DisplayName, @{N="SKUs";E={$_.AssignedLicenses.SkuId}} | Export-Csv "016_LicenseSKUs.csv" } Catch { Write-Error $_ }
