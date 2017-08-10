$Zone1Servers="BP-VM-CPX101","BP-VM-CPX102","BP-VM-CPX103","BP-VM-CPX104","BP-VM-CPX105","BP-VM-CPX106","BP-VM-CPX107","BP-VM-CPX108"
$Zone2Servers="BP-VM-CPX201","BP-VM-CPX202","BP-VM-CPX203","BP-VM-CPX204"
$Zone3Servers="BP-VM-CPX301","BP-VM-CPX302","BP-VM-CPX303","BP-VM-CPX304","BP-VM-CPX305","BP-VM-CPX306","BP-VM-CPX307","BP-VM-CPX308"
$Zone4Servers="BP-VM-CPX401","BP-VM-CPX402","BP-VM-CPX403","BP-VM-CPX404","BP-VM-CPX405","BP-VM-CPX406","BP-VM-CPX407","BP-VM-CPX408"
$Zone5Servers="BP-VM-CPX501","BP-VM-CPX502","BP-VM-CPX503","BP-VM-CPX504","BP-VM-CPX505","BP-VM-CPX506","BP-VM-CPX507","BP-VM-CPX508"
$Zone6Servers="BP-VM-CPX601","BP-VM-CPX602","BP-VM-CPX603","BP-VM-CPX604","BP-VM-CPX605","BP-VM-CPX606","BP-VM-CPX607","BP-VM-CPX608"

$CPServers = "BP-VM-CPX101","BP-VM-CPX102","BP-VM-CPX103","BP-VM-CPX104","BP-VM-CPX105","BP-VM-CPX106","BP-VM-CPX107","BP-VM-CPX108","BP-VM-CPX201","BP-VM-CPX202","BP-VM-CPX203","BP-VM-CPX204","BP-VM-CPX301","BP-VM-CPX302","BP-VM-CPX303","BP-VM-CPX304","BP-VM-CPX305","BP-VM-CPX306","BP-VM-CPX307","BP-VM-CPX308","BP-VM-CPX401","BP-VM-CPX402","BP-VM-CPX403","BP-VM-CPX404","BP-VM-CPX405","BP-VM-CPX406","BP-VM-CPX407","BP-VM-CPX408","BP-VM-CPX501","BP-VM-CPX502","BP-VM-CPX503","BP-VM-CPX504","BP-VM-CPX505","BP-VM-CPX506","BP-VM-CPX507","BP-VM-CPX508","BP-VM-CPX601","BP-VM-CPX602","BP-VM-CPX603","BP-VM-CPX604","BP-VM-CPX605","BP-VM-CPX606","BP-VM-CPX607","BP-VM-CPX608"

$CPUser = Read-Host 'Enter CP User: '
	
	mkdir C:\scripts\CPLogs\$CPuser
	
foreach ($server in $CPServers)
	{
		$StartupLogSource = "\\$server\CounterPoint\CounterPointSQL\Logs\CPLogs\CounterPoint_Startup_$cpuser.log"
		$OperationsLogSource = "\\$server\CounterPoint\CounterPointSQL\Logs\CPLogs\CounterPoint_$cpuser.log"
		$SecurityLogSource = "\\$server\CounterPoint\CounterPointSQL\Logs\CPLogs\Security_$CPUser.log"
		$SessionLogSource = "\\$server\CounterPoint\CounterPointSQL\Logs\CounterPoint_$cpuser.log"
		
		$StartupLogDest = "c:\scripts\CPLogs\$CPuser\'$CPUser'_'$Server'_Startup.log"
		$OperationsLogDest = "c:\scripts\CPLogs\$CPuser\'$CPUser'_'$Server'_Operations.log"
		$SecurityLogDest = "c:\scripts\CPLogs\$CPuser\'$CPUser'_'$Server'_Security.log"
		$SessionLogDest = "c:\scripts\CPLogs\$CPuser\'$CPUser'_'$Server'_Session.log"
		
		Echo "Copying $StartupLogSource to $StartupLogDest"
		copy-item $StartupLogSource $StartupLogDest
		Echo "Copying $OperationsLogSource to $OperationsLogDest"
		copy-item $OperationsLogSource $OperationsLogDest
		Echo "Copying $SecurityLogSource to $SecurityLogDest"
		copy-item $SecurityLogSource $SecurityLogDest
		Echo "Copying $SessionLogSource to $SessionLogDest"
		copy-item $SessionLogSource $SessionLogDest
	}