Try { Connect-MgGraph -Scopes "Reports.Read.All"; Get-MgReportTeamsUserActivityUserDetail -Period "D30" | Export-Csv "046_TeamsActivity.csv" } Catch { Write-Error $_ }
