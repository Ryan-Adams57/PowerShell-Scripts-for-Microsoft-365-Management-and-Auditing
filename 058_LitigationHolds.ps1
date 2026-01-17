Try { Connect-ExchangeOnline; Get-EXOMailbox -ResultSize Unlimited | Select-Object UserPrincipalName, LitigationHoldEnabled | Export-Csv "058_LitigationHolds.csv" } Catch { Write-Error $_ }
