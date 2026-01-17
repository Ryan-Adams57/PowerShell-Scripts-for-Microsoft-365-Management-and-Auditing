Try { Connect-ExchangeOnline; Get-CASMailbox -ResultSize Unlimited | Export-Csv "056_LegacyProtocols.csv" } Catch { Write-Error $_ }
