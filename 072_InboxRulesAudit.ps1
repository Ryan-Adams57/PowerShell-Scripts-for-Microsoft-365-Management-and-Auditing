Try { Connect-ExchangeOnline; Get-EXOMailbox | ForEach-Object { Get-InboxRule -Mailbox $_.UPN } | Export-Csv "072_InboxRules.csv" } Catch { Write-Error $_ }
