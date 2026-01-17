Try { Connect-ExchangeOnline; Get-AcceptedDomain | Export-Csv "073_Domains.csv" } Catch { Write-Error $_ }
