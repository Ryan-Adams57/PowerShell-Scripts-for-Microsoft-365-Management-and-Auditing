Try { Connect-MgGraph -Scopes "Organization.Read.All"; Get-MgSubscribedSku | Export-Csv "081_Licenses.csv" } Catch { Write-Error $_ }
