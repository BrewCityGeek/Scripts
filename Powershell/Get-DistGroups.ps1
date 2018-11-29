#This script scrapes all MGR or BP email addresses and gathers what Distribution Groups they are a part of.
#Replace the MGR with BP and the destination directory accordingly.

$storeNumbers = 001 .. 999

foreach ($store in $storeNumbers) {
	$user = dsquery user -samid MGR$store
	if($user -ne $null)	{
		Get-ADUser MGR$store | Get-ADPrincipalGroupMembership | Select -Expand Distinguishedname | Get-DistributionGroup -IgnoreDefaultScope | Select Alias | Export-csv c:\users\agossen\desktop\tammy\managers\MGR$store.csv -NoTypeInformation
	}
}