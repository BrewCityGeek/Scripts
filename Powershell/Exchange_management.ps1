$Session = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri http://hrtv-exch01.battplus.co/PowerShell -Authentication Kerberos -Credential battplus\gossena
Import-PSSession $Session -Verbose > $null

Add-PSSnapin *Exchange* -ErrorAction SilentlyContinue > $null

$Command = 0

Set-AdServerSettings -ViewEntireForest $True
$LocalCredentials = Get-Credential corp\agossen2
$RemoteCredentials = Get-Credential bpluscorp\agossen
$Username = Read-Host "Enter Usename to migrate"
.\Prepare-MoveRequest.ps1 -Identity $Username@batteriesplus.com -RemoteForestDomainController bph-mdc02.batteriesplus.com -RemoteForestCredential $RemoteCredentials -LocalForestDomainController hrtp-mdc002.corp.battplus.co -LocalForestCredential $LocalCredentials –LinkedMailUser

do {
    $Command = Read-Host "Enter desired command (Type 'exit' to quit)"
    
    While ($Command -eq "exit")
        {
        Get-PSSession | Remove-PSSession
        Invoke-Expression $Command -ErrorAction SilentlyContinue
        }

    Invoke-Expression $Command -ErrorAction SilentlyContinue
    }
until ($Command -eq "exit")
# SIG # Begin signature block
# MIIFPwYJKoZIhvcNAQcCoIIFMDCCBSwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYqMR0w7cFMWph4vfWG5PFpuv
# ROWgggMIMIIDBDCCAeygAwIBAgIQ532grBRei4lBFW/d28+NTTANBgkqhkiG9w0B
# AQsFADASMRAwDgYDVQQDEwdBbmR5IENBMB4XDTE0MDQwOTE4NTgxOVoXDTM5MTIz
# MTIzNTk1OVowEzERMA8GA1UEAxMIQW5keSBTUEMwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQC/EY3snKPLNV63E3SSjjI/FDgk9gXnMWxg+JQ/erGpXDZi
# jKWHWhsYj97HaDyw4fRchU2Y7ZvmehQTYGRObeBIBpcR0H6xY7gh3KceZq2YR7/M
# fgadD/hWXUfCK4q4QYNhNnTPlN9RkO4Un52yBjlzzaANoMUXKG8SyLBo9uYK4bGr
# BcebZSDlq/pUjezUYpmU2dKVLuPyti2U6YBZ+gocA9d3c4bFzUzbHeSFIdg6rchb
# FM4E1aOwjNtpwU2KbDt8nuC7qvnZdKJ8g7JanaVw60G/Tck/OAeWXqM21iqRw/r5
# AMK3Drtv1LzSIfrjPK4olXSRcy2cFJ88HMvoWr5hAgMBAAGjVTBTMAwGA1UdEwEB
# /wQCMAAwQwYDVR0BBDwwOoAQl+hqDhHt0HfgWugeW/8oFqEUMBIxEDAOBgNVBAMT
# B0FuZHkgQ0GCEJ6hrNjhSV6gTl0doHDck/cwDQYJKoZIhvcNAQELBQADggEBAFqM
# cjPMd7vzWAGn3T+1PItQtS9L6Zrs/jxsw4/NdzoFdEgPs6/hMWLgd4nLjfd8AZ8s
# 3MdnOR52VPbZPStmDYliDNVeKvub2KLFWzfYcYp3G7u92c29QKp1KQMOi/VmvrUS
# 5YltHUH53bt0rQJG3ZjOp0C18KjixgXt7G5A4po1/6Mz4zCq1irpBNLH/EbZTsUk
# 5WyEnWdNVYnR+GmyIkth1R3N56R8n0zWGQBXQ0j09WQ/hgkVD3++Y7UQwVEc9fC8
# kudh6oKgFUa/zknudyIxa9GnRJwRVxBHpAsW/+cE4UXVMy//M4LOU5gxxq9J5Hcu
# af9OCAk952vgTA4WwKExggGhMIIBnQIBATAmMBIxEDAOBgNVBAMTB0FuZHkgQ0EC
# EOd9oKwUXouJQRVv3dvPjU0wCQYFKw4DAhoFAKBSMBAGCisGAQQBgjcCAQwxAjAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMCMGCSqGSIb3DQEJBDEWBBRzfGd8
# RhGxrAV7S7t8/x+Pdtbl/jANBgkqhkiG9w0BAQEFAASCAQAPA6ijs9Cozmsg+Kfc
# O0rmTYAHvedBJpYEAiOFAIdJLwXZ1xY28sG+ZYbK8sF3QDhpo1OQWJMVp0JjYt5a
# FUAyvomwllmP09MS0S+icfbhfdaXspg2boYcexKKyPmhRxU/V1ZZQbtZfK9UNk7n
# I8LviLlglu3BYfgWXzRq0s8tGu2ZU0FFY6HgUIN5DVw7FD/oX8/IYqNFLPLNYZDv
# V/f9wyu6Y0yBfRyaQb6+LEt3HW44VQycHO+RE+dVXCwF8P/Yguq78ihZG0f/L9GK
# H2hO7uJS950YpSfUR4JeIyYESZtxQLU45jcWeulXQwHazX26DeoXUwmfl1afyAIF
# 6nTI
# SIG # End signature block
