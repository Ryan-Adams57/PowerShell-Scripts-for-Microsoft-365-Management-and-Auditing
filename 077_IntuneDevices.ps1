Try { Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"; Get-MgDeviceManagementManagedDevice -All | Export-Csv "077_IntuneInventory.csv" } Catch { Write-Error $_ }
