$UserCredential = Get-Credential -message "Enter Credentials" -username "bpluscorp\mneusen"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://bph-vm-exch01.batteriesplus.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session
Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue
Set-AdServerSettings -ViewEntireForest $True
Set-Location C:\Scripts\Exch
clear-host