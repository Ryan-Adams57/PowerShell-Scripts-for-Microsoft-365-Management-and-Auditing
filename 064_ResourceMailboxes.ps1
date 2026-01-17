Try { Connect-ExchangeOnline; Get-EXOMailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox | Export-Csv "064_Resources.csv" } Catch { Write-Error $_ }
