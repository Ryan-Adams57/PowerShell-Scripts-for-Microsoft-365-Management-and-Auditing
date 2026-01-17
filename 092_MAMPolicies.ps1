Try { Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"; Get-MgDeviceManagementManagedAppPolicy | Export-Csv "092_MAM.csv" } Catch { Write-Error $_ }
