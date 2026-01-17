Try { Connect-ExchangeOnline; Get-TransportRule | Export-Csv "062_TransportRules.csv" } Catch { Write-Error $_ }
