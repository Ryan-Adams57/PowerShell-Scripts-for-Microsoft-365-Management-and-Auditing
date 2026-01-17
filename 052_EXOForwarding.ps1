Try { Connect-ExchangeOnline; Get-EXOMailbox -ResultSize Unlimited | Where-Object {$_.ForwardingSmtpAddress} | Export-Csv "052_EXOForwarding.csv" } Catch { Write-Error $_ }
