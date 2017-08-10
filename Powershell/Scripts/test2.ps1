$password = (Get-Credential).GetNetworkCredential().password
$ESXUN = read-host 'ESX Username'
$ESXPW = ConvertFrom-SecureString $securepw
write-host $esxun
write-host $esxpw