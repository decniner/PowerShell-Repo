<#

Add a  user to remote desktop users group of of each PC in order using csv and powershell.

c:\list.csv contains the following:

computername, userid
pcname1, user1
pcname2, user2

#>

$csv = Import-Csv c:\list.csv
foreach ($line in $csv) {
    $computer = $line.computername
    $user = $line.userid
    icm $computer {net localgroup "remote desktop users" /add using:$user
}
