Try { Connect-ExchangeOnline; Get-EXOMailbox | Get-MailboxAutoReplyConfiguration | Where-Object {$_.AutoReplyState -ne "Disabled"} | Export-Csv "066_OOFStatus.csv" } Catch { Write-Error $_ }
