$Source = "C:\scripts\TELIUM"
$destination = 'E:\','F:\','G:\','H:\','I:\','J:\','K:\','L:\','M:\','N:\'
#Remove-Item 'E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium','L:\Telium','M:\Telium','N:\Telium' -force
#Remove-Item 'E:\Telium\PK-RGEN31-12010B3','F:\Telium\PK-RGEN31-12010B3','G:\Telium\PK-RGEN31-12010B3','H:\Telium\PK-RGEN31-12010B3','I:\Telium\PK-RGEN31-12010B3','J:\Telium\PK-RGEN31-12010B3','K:\Telium\PK-RGEN31-12010B3','L:\Telium\PK-RGEN31-12010B3','M:\Telium\PK-RGEN31-12010B3','N:\Telium\PK-RGEN31-12010B3' -force
#Remove-Item 'E:\Telium\PK-RGEN31-12012B3','F:\Telium\PK-RGEN31-12012B3','G:\Telium\PK-RGEN31-12012B3','H:\Telium\PK-RGEN31-12012B3','I:\PK-RGEN31-12012B3','J:\Telium\PK-RGEN31-12012B3','K:\Telium\PK-RGEN31-12012B3','L:\Telium\PK-RGEN31-12012B3','M:\Telium\PK-RGEN31-12012B3','N:\Telium\PK-RGEN31-12012B3' -force
#Remove-Item 'E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium','L:\Telium','M:\Telium' -force
#Remove-Item 'E:\PK-RGEN31-12014B3','F:\PK-RGEN31-12014B3','G:\PK-RGEN31-12014B3','H:\PK-RGEN31-12014B3','I:\PK-RGEN31-12014B3','J:\PK-RGEN31-12014B3','K:\PK-RGEN31-12014B3','L:\PK-RGEN31-12012B3','M:\PK-RGEN31-12012B3' -force
$Far = 'E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium','L:\Telium','M:\Telium','N:\Telium'
foreach ($F in $Far)
    {
    if (test-path $F)
        {
        Remove-Item $F -Recurse
        Write-host "removed" $F
        }
        else
        {
        Write-host "Path" $F "is not availible"
        }
    }
foreach ($dest in $destination)
    {
    if (Test-Path $dest)
        {
        Copy-Item $source $dest -Recurse
        #$test = (Get-Itemproperty $dest.LastWriteTime)
        #write-host $test
        Get-ChildItem -Path $dest -Recurse
        write-host "done" $dest
        $Eject =  New-Object -comObject Shell.Application
        $Eject.NameSpace(17).ParseName($dest).InvokeVerb(“Eject”)
        $x=$x+1
        write-host $x
        }
        else
        {
        Write-host "Drive not availible "$dest
        }
    }    
#Write-Host "Press any key to continue ..."

#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")