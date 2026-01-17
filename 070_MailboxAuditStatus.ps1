Try { Connect-ExchangeOnline; Get-EXOMailbox | Select-Object UserPrincipalName, AuditEnabled | Export-Csv "070_AuditStatus.csv" } Catch { Write-Error $_ }
