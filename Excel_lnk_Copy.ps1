$Computer = Get-Content -Path Computers.txt
foreach ($Dest in $Computer) {
	Write-Host "Copying Excel 2016.lnk to "$Dest
	Start-BitsTransfer -Source "\\bp1222\c$\Users\trainee1\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Excel 2016.lnk" -Destination "\\$Dest\c$\Users\trainee1\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\"
	Start-BitsTransfer -Source "\\bp1222\c$\Users\trainee1\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Excel 2016.lnk" -Destination "\\$Dest\c$\Users\trainee1\Desktop"
	}