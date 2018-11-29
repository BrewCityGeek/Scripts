$Computer = Get-Content -Path C:\Batch_Files\GP_Users.txt
foreach ($Dest in $Computer) {
	Write-Host "Copying GP2016Installer.exe to "$Dest
	Start-BitsTransfer -Source "\\hrtd-fsa001\Dept\GP Users\Deployment\GP2016Installer\GP2016Installer.exe" -Destination \\$Dest\c$
	}