Try { Connect-MgGraph -Scopes "TeamworkDevice.Read.All"; Get-MgTeamworkDevice | Export-Csv "048_TeamsDevices.csv" } Catch { Write-Error $_ }
