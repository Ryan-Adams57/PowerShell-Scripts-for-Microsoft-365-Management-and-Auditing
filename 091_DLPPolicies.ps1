Try { Connect-ExchangeOnline; Get-DlpCompliancePolicy | Export-Csv "091_DLP.csv" } Catch { Write-Error $_ }
