Try { Connect-ExchangeOnline; Get-EXOMailbox | Get-MailboxJunkEmailConfiguration | Export-Csv "065_JunkConfig.csv" } Catch { Write-Error $_ }
