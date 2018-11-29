Set-Location C:\Scripts

#Program Shortcuts
new-item alias:np -value "C:\Program Files (x86)\Notepad++\Notepad++.exe"
new-item alias:excel -value "C:\Program Files\Microsoft Office\Office15\EXCEL.exe"
new-item alias:RDP -value "c:\scripts\Start-RDP.ps1"

#Module Shortcuts
new-item alias:EXCH -value "c:\scripts\exch.ps1"
new-item alias:OLDEXCH -value "c:\scripts\oldexch.ps1"
new-item alias:ESX -value "c:\scripts\esx.ps1"
new-item alias:Citrix -value "c:\scripts\Citrix.ps1"

#Default Modules
Import-Module ActiveDirectory
clear-host