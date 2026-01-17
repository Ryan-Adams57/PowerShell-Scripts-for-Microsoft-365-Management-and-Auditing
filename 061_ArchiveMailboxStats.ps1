Try { Connect-ExchangeOnline; Get-EXOMailbox -Archive | Get-EXOMailboxStatistics -Archive | Export-Csv "061_ArchiveStats.csv" } Catch { Write-Error $_ }
