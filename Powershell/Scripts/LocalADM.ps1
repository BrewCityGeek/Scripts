#LocalADM.ps1
invoke-command {
	net localgroup administrators | 
	where {$_ -AND $_ -notmatch "command completed successfully"} | 
	select -skip 4
}
