$user = Read-Host "Please enter the name of a user 'Last Name, First Name'"

if (Get-Mailbox $user){
	$dgs = Get-DistributionGroup -ResultSize unlimited
	write-host
	write-host "$user is a member of the following distribution groups (this may take a few minutes):"
	write-host

	$dgs | foreach {if ((Get-DistributionGroupMember $_.Name) -match $user){write-host $_.Name}}
}
else{
	write-host "User not found...run script again. Sober up first." -foregroundcolor red
	write-host
}