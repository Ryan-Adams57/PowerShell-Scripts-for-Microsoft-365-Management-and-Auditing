Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All -Property OnPremisesExtensionAttributes | Export-Csv "010_ExtensionAttributes.csv" } Catch { Write-Error $_ }
