Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All | Where-Object { $_.ShowInAddressList -eq $false } | Export-Csv "021_HiddenGALUsers.csv" } Catch { Write-Error $_ }
