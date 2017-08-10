$UserCredential = Get-Credential -message "Enter Credentials" -username "BATTPLUS\NeusenM"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://hrtv-exch01.battplus.co/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session
Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue
Set-AdServerSettings -ViewEntireForest $True
Set-Location C:\Scripts\Exch
clear-host