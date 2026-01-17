Try { Connect-MgGraph -Scopes "DeviceManagementServiceConfig.Read.All"; Get-MgDeviceManagementWindowsAutopilotDeviceIdentity | Export-Csv "099_Autopilot.csv" } Catch { Write-Error $_ }
