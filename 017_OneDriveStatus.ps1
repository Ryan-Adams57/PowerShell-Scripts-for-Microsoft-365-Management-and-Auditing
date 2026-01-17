Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All -Property ProvisionedPlans | Export-Csv "017_OneDriveStatus.csv" } Catch { Write-Error $_ }
