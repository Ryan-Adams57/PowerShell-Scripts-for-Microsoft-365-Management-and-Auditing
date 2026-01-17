Try { Connect-MgGraph -Scopes "Reports.Read.All"; Get-MgReportSharePointSiteUsageDetail -Period "D7" | Export-Csv "030_SPStorage.csv" } Catch { Write-Error $_ }
