Try { Connect-ExchangeOnline; Get-MobileDevice | Export-Csv "057_MobileDevices.csv" } Catch { Write-Error $_ }
