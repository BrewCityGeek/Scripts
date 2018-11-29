Remove-Item 'D:\Telium','E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium','L:\Telium' -force
$Source = "C:\scripts\TELIUM"

$destination = 'D:\','E:\','F:\','G:\','H:\','I:\','J:\','K:\','L:\'

foreach ($dest in $destination){Copy-Item $source $dest -Recurse}
foreach ($dest in $destination)    
    {
        $TestResult = (Get-Itemproperty 'D:\Telium','E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium','L:\Telium').LastWriteTime
        Write-Output $TestResult
	}