Try { Connect-MgGraph -Scopes "User.Read.All"; Get-MgUser -All -Property LastPasswordChangeDateTime | Export-Csv "018_PasswordChangeDate.csv" } Catch { Write-Error $_ }
