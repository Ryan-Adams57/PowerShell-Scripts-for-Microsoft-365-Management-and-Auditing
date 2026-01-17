Try { Connect-MgGraph -Scopes "Group.Read.All"; Get-MgGroup -All -Property SensitivityLabel | Where-Object {$_.SensitivityLabel} | Export-Csv "047_GroupLabels.csv" } Catch { Write-Error $_ }
