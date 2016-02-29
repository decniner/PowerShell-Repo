<#

Add a  user to remote desktop users group of of each PC in order using csv and powershell.

c:\list.csv contains the following:

computername, userid
pcname1, user1
pcname2, user2

#>
$csv = Import-Csv C:\Temp\list.csv

foreach ($line in $csv) {
    $computer = $line.computername
    $user = $line.userid
    
    $ScriptBlockContent = {
        param ($ruser)
        net localgroup "remote desktop users" /del "$ruser"
    }
    
    Invoke-Command -ComputerName $computer -ScriptBlock $ScriptBlockContent -ArgumentList $user
        
}
