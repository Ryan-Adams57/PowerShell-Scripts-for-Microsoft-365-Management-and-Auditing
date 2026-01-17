Try { Connect-ExchangeOnline; Get-DistributionGroup | Where-Object {$_.HiddenFromAddressListsEnabled} | Export-Csv "075_HiddenGroups.csv" } Catch { Write-Error $_ }
