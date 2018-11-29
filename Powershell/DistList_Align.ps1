$DistGroup = Read-Host "Enter Dist Group to Align"

$List = 0
$List = Get-Content -Path C:\Users\agossen\Desktop\CorpStore\Align\$DistGroup.csv
foreach ($Name in $List) {
	Add-DistributionGroupMember -BypassSecurityGroupManagerCheck -Identity $DistGroup -Member $Name -Confirm:$False
	}