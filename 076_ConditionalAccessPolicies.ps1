Try { Connect-MgGraph -Scopes "Policy.Read.All"; Get-MgIdentityConditionalAccessPolicy | Export-Csv "076_CAPolicies.csv" } Catch { Write-Error $_ }
