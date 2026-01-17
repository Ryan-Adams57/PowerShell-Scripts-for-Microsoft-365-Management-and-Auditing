Try { Connect-ExchangeOnline; Get-RetentionCompliancePolicy | Export-Csv "090_Retention.csv" } Catch { Write-Error $_ }
