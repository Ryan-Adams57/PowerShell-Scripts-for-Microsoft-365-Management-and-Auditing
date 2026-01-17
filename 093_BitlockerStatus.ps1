Try { Connect-MgGraph -Scopes "BitlockerKey.Read.All"; Get-MgInformationProtectionBitlockerRecoveryKey | Export-Csv "093_Bitlocker.csv" } Catch { Write-Error $_ }
