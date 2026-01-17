Try { Connect-ExchangeOnline; Get-EXOMailbox | Get-MailboxRegionalConfiguration | Export-Csv "074_Regional.csv" } Catch { Write-Error $_ }
