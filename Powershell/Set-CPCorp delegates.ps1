﻿Import-Csv "C:\Scripts\Powershell\Master.csv" | foreach {Add-MailboxPermission -Identity $_.Name -User tthiede@batteriesplus.com -AccessRights FullAccess -InheritanceType all -AutoMapping $false}
Import-Csv "C:\Scripts\Powershell\Master.csv" | foreach {Add-MailboxPermission -Identity $_.Name -User jwiltsey@batteriesplus.com -AccessRights FullAccess -InheritanceType all -AutoMapping $false}
Import-Csv "C:\Scripts\Powershell\Master.csv" | foreach {Add-ADPermission -Identity $_.Name -User bpluscorp\tthiede -ExtendedRights "Send As"}
Import-Csv "C:\Scripts\Powershell\Master.csv" | foreach {Add-ADPermission -Identity $_.Name -User corp\jwiltsey -ExtendedRights "Send As"}