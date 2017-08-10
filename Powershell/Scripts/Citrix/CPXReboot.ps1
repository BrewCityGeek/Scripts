Add-PSSnapin Citrix.*
Add-PSSnapin VMware.VimAutomation.Core

$servername = read-host 'Which Server Needs a Reboot?'
$ESXUN = read-host 'ESX Username'
$ESXSecurePW = read-host 'ESX Password' -AsSecureString
$ESXPW= [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ESXSecurePW))

#Retrieve Worker Group name from affected server
$WG = Get-XAWorkerGroup -server $servername

#Remove affected server from Worker Group
Write-host 'Removing' $servername 'from'$wg'...'
Remove-XAWorkerGroupServer $wg $servername

#Connect to MSN VCenter Server
Write-host 'Connecting to MSNV-VCN001 as' $ESXUN
Connect-VIServer -Server msnv-vcn001.corp.battplus.co -Protocol https -User $ESXUN -password $ESXPW

#Get VM and initiate reboot
Write-host 'Rebooting' $servername
Get-VMGuest $servername | Restart-VMGuest
sleep 15

#Wait until the VM is back online, with VMWare tools running
write-host "Waiting for VMWare Tools to Start..."
do {
$toolsStatus = (Get-VM $servername | Get-View).Guest.ToolsStatus
write-host $toolsStatus
sleep 3
} until ( $toolsStatus -eq "toolsOk" )

Write-host 'VMWare Tools are running!'
sleep 20

#Add Removed server back to worker group
Write-host 'Adding' $Servername 'back to' $WG
Add-XAWorkerGroupServer $wg $servername