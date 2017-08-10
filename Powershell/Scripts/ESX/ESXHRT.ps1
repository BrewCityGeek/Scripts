$ESXUN = read-host 'ESX Username'
$ESXSecurePW = read-host 'ESX Password' -AsSecureString
$ESXPW= [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ESXSecurePW))
write-host "Connecting to HRTV-VCN001..."
Connect-VIServer -Server HRTV-vcn001.corp.battplus.co -Protocol https -User $ESXUN -password $ESXPW