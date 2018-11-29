$EnterUser = Read-Host "Enter Alias"

Set-ADUser -UserPrincipalName $EnterUser@batteriesplus.com -Identity "$EnterUser"

Get-ADUser -Identity "$EnterUser" | Select UserPrincipalName