Try { Connect-ExchangeOnline; Get-EXOMailbox | Where-Object {!$_.EmailAddressPolicyEnabled} | Export-Csv "069_ManualPolicy.csv" } Catch { Write-Error $_ }
