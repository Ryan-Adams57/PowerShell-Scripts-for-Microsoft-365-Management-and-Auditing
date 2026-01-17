Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -All -Property ResourceProvisioningOptions | Export-Csv "039_GroupSource.csv" } Catch { Write-Error $_ }
