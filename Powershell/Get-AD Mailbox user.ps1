﻿get-aduser -filter {(Enabled -eq $True) -and (EmailAddress -like "*")} -Properties DisplayName, EmailAddress | select DisplayName, EmailAddress | export-csv C:\Scripts\Email.csv -NoTypeInformation