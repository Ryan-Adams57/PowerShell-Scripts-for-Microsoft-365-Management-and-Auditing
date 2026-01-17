Try { Connect-ExchangeOnline; Get-EXOMailbox -SoftDeleted | Export-Csv "063_DeletedMailboxes.csv" } Catch { Write-Error $_ }
