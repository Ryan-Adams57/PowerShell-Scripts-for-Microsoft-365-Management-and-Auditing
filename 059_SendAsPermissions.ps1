Try { Connect-ExchangeOnline; Get-EXOMailbox -ResultSize Unlimited | Get-EXORecipientPermission | Where-Object {$_.Trustee -notlike "*SELF*"} | Export-Csv "059_SendAs.csv" } Catch { Write-Error $_ }
