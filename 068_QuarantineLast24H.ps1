Try { Connect-ExchangeOnline; Get-QuarantineMessage -StartCursorDate (Get-Date).AddDays(-1) | Export-Csv "068_Quarantine.csv" } Catch { Write-Error $_ }
