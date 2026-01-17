Try { Connect-ExchangeOnline; Get-EXOMailboxStatistics -ResultSize Unlimited | Export-Csv "051_MailboxSizes.csv" } Catch { Write-Error $_ }
