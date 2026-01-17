Try { Connect-ExchangeOnline; Get-EXOMailbox | Get-FocusedInbox | Export-Csv "071_FocusedInbox.csv" } Catch { Write-Error $_ }
