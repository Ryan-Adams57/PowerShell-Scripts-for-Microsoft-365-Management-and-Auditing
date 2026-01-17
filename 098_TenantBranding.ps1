Try { Connect-MgGraph -Scopes "Organization.Read.All"; Get-MgOrganization | Select-Object DisplayName, VerifiedDomains | Export-Csv "098_TenantInfo.csv" } Catch { Write-Error $_ }
