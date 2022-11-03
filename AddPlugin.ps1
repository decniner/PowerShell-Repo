<#
Contact decniner@gmail.com

This script will add plugin/s to Radiate core script folder.

1. ask for ps1 source file
2. copy source to destination > "c:\program filesx86\radiate\core\scripts\"
3. modify Radiate config file
#>

$ErrorActionPreference = "Stop"

$hostname = $env:COMPUTERNAME
$dest = "\\$hostname\C$\Program Files (x86)\Company\Radiate\Core\Scripts\"
$coreFolder = "\\$hostname\C$\Program Files (x86)\Company\Radiate\Core\"
$configcontent = Get-Content $coreFolder\Configuration.xml

#ask for ps1 source file
$ps1path = Read-Host “Path to new script”


if ((Test-Path $ps1path) -eq $true) {
    $error.Clear()
    try{
        
        #copy new plugin to Radiate script folder
        Copy-Item -Path $ps1path -Destination $dest -Force
        
        #update Radiate config file
        $scriptInfo = Get-Item $ps1path
        $ps1Name = $scriptInfo.Name
        $scriptname = $scriptInfo.Name.Split(".")[0]
        $updateConfig = "    Script name $scriptname version 0.1 with filename $ps1Name"
        $configcheck = $configcontent | where {$_ -like "*$ps1Name*"}
        
        if($configcheck.Length -eq 0) {
             $linenumber =  ($configcontent).Count
             $configcontent[$linenumber -3] = "{0}`r`n{1}" -f $updateConfig, $configcontent[$linenumber -3]
             $configcontent | Set-Content $coreFolder\Configuration.xml
             "New plugin added. Please restart Radiate, click Core tab and check ""$scriptname"" plugin."
        }
        else {
             “Config file update not required. New plugin added. Please restart Radiate, click Core tab and check ""$scriptname"" plugin."
        }
    }
    catch{
        $error
    } 
}  

else {
    "File does not exist"
}
