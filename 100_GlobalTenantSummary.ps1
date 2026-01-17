Try { Connect-MgGraph -Scopes "Organization.Read.All"; Get-MgOrganization | Export-Csv "100_GlobalSummary.csv" } Catch { Write-Error $_ }
