Remove-Item 'D:\Telium','E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium' -force
$Source = "C:\scripts\TELIUM"

$destination = 'D:\','E:\','F:\','G:\','H:\','I:\','J:\','K:\'

foreach ($dest in $destination)
    {
        Copy-Item $source $dest -Recurse
        $Source1 = (Get-ChildItem C:\scripts\Telium).LastWriteTime
        $TestResult = Compare-Object $Source1  (Get-ChildItem 'D:\Telium','E:\Telium','F:\Telium','G:\Telium','H:\Telium','I:\Telium','J:\Telium','K:\Telium').LastWriteTime -IncludeEqual
        Write-Output $TestResult
	}
