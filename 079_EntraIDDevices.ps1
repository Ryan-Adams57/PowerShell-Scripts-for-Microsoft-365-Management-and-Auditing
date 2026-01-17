Try { Connect-MgGraph -Scopes "Device.Read.All"; Get-MgDevice -All | Export-Csv "079_EntraDevices.csv" } Catch { Write-Error $_ }
