Try { Connect-ExchangeOnline; Get-PublicFolder -Recurse | Export-Csv "067_PublicFolders.csv" } Catch { Write-Error $_ }
