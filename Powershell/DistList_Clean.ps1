$DistGroup = Read-Host "Enter Dist Group to Clean"

$List = 0
$List = Get-Content -Path C:\Users\agossen\Desktop\CorpStore\Clean\$DistGroup.csv
foreach ($Name in $List) {
	Remove-DistributionGroupMember -BypassSecurityGroupManagerCheck -Identity $DistGroup -Member $Name -Confirm:$False
	}