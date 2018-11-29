Get-ADGroupMember -Identity "Dist86" | Select Name | Export-Csv -Path C:\VUDU\DIST86.CSV -NoTypeInformation
Get-ADGroupMember -Identity "MGR Dist86" | Select Name | Export-Csv -Path C:\VUDU\MGR86.CSV -NoTypeInformation
Get-ADGroupMember -Identity "Dist80" | Select Name | Export-Csv -Path C:\VUDU\DIST80.CSV -NoTypeInformation
Get-ADGroupMember -Identity "MGR Dist80" | Select Name | Export-Csv -Path C:\VUDU\MGR80.CSV -NoTypeInformation
Get-ADGroupMember -Identity "Dist89" | Select Name | Export-Csv -Path C:\VUDU\DIST89.CSV -NoTypeInformation
Get-ADGroupMember -Identity "MGR Dist89" | Select Name | Export-Csv -Path C:\VUDU\MGR89.CSV -NoTypeInformation


Import-Csv "C:\VUDU\Master.csv" | foreach {Remove-MailboxPermission -Identity $_.Name -User bpluscorp\tthiede -AccessRights FullAccess -InheritanceType all -Confirm:$false}
Import-Csv "C:\VUDU\Master.csv" | foreach {Add-MailboxPermission -Identity $_.Name -User bpluscorp\tthiede -AccessRights FullAccess -InheritanceType all -AutoMapping $false}
Import-Csv "C:\VUDU\Master.csv" | foreach {Add-ADPermission -Identity $_.Name -User corp\jwiltsey -ExtendedRights "Send As"}