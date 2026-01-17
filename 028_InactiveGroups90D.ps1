Try { Connect-MgGraph -Scopes "Reports.Read.All"; Get-MgReportOffice365GroupActivityDetail -Period "D90" | Export-Csv "028_InactiveGroups.csv" } Catch { Write-Error $_ }
