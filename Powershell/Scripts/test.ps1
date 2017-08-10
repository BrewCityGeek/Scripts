Add-PSSnapin Citrix.*
Add-PSSnapin VMware.VimAutomation.Core

$servernames = read-host 'Which Server Needs a Reboot?'
$servernames = $servernames.Split(',')
$ESXUN = read-host 'ESX Username'
$ESXSecurePW = read-host 'ESX Password' -AsSecureString
$ESXPW= [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ESXSecurePW))

#Connect to MSN VCenter Server
Write-host 'Connecting to 10.100.1.47 as' $ESXUN
Connect-VIServer -Server 10.100.1.47 -Protocol https -User $ESXUN -password $ESXPW

#Get VM and initiate reboot
foreach ($server in $servernames){
Get-VMGuest $server | Restart-VMGuest
Write-host 'Rebooting' $server
sleep 5
}

write-host "Waiting for VMWare Tools to Start..."

foreach ($server in $serversames){
	do {
	$toolsStatus = (Get-VM $server | Get-View).Guest.ToolsStatus
	write-host $toolsStatus
	sleep 3
	} until ( $toolsStatus -eq "toolsOk" )

	Write-host $server 'is online!'