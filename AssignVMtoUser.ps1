This script assigns virtual machines to users or pools.Optionally, it also adds RD Virtualization Host servers to the RD Connection Broker server. While assigning personal virtual desktops, the script prompts for confirmation to continue overriding the current assignment (if any) for the virtual machines.
NOTE: This script should be run on the RD Connection Broker server.
Copy the script to a file and save it as Set-VMAssignment.ps1. Use get-help Set-VMAssignment.ps1 to get usage information and examples.
PowerShell
<#  
.Synopsis  
    Assigns virtual machines to users or pools. The script, optionally also adds RD Virtualization Host servers to the RD Connection Broker server. 
     
.Description  
    Assigns virtual machines to users or pools. The script, optionally also adds RD Virtualization Host servers to the RD Connection Broker server. While assigning personal virtual desktops, the script prompts for confirmation to continue overriding the current assignment (if any) for the virtual machines. 
    NOTE: This script should be run on the RD Connection Broker server. 
 
.Parameter AssignmentFileName 
    Text file containing assignment information in the format '<VMName> [POOL|USER] <User Name/Pool Name>'. User Name should be in 'domain\user' format. 
     
.Parameter HostFileName 
    Text file containing list of RD Virtualization Host servers that need to be added to the connection broker. Each new line seperated entry in the file is treated as a single RD Virtual Host server's name. 
     
.Parameter PoolFileName 
    Text file containing list of pools that need to be created. Each new line seperated entry in the file is treated as the name of a pool. 
 
.Parameter Credential 
    Credentials that have write acces on the Active Directory for Personal Virtual Desktop assignment. 
 
.Parameter Force 
    Switch: The script seeks confirmation from the user in case it is reassigning a virtual machine to another user unless this switch is used. Providing this switch would force continue the reassignment. 
       
.Example  
    PS C:\> Set-VMAssignment.ps1 -AssignmentFileName "C:\Users\Administrator\Desktop\assninfo.txt" 
     
    The above example performs virtual machine assignment as mentioned in 'assninfo.txt'. 
 
.Example  
    PS C:\> Set-VMAssignment.ps1 -AssignmentFileName "C:\Users\Administrator\Desktop\assninfo.txt" -PoolFileName "C:\Users\Administrator\Desktop\pools.txt" 
     
    The above example creates virtual machine pools with names mentioned in 'pools.txt' and performs virtual machine assignment as mentioned in 'assninfo.txt'. 
 
.Example  
    PS C:\> Set-VMAssignment.ps1 -AssignmentFileName "C:\Users\Administrator\Desktop\assninfo.txt" -PoolFileName "C:\Users\Administrator\Desktop\pools.txt" -HostFileName "C:\Users\Administrator\Desktop\hosts.txt" 
     
    The above example adds entries mentioned in 'hosts.txt' as RD Virtualization Host servers to the RD Connection Borker server on which the script is being executed, creates virtual machine pools with names mentioned in 'pools.txt' and performs virtual machine assignment as mentioned in 'assninfo.txt'. 
 
.Notes  
    Name     : Set-VMAssignment.ps1 
     
#> 
param ( 
    [Parameter(Mandatory=$TRUE, HelpMessage="Text file containing lines in the format '<VM Name> [POOL/USER] <Pool Name/User Name>'.")] 
    [string]$AssignmentFileName, 
     
    [Parameter(Mandatory=$FALSE, HelpMessage="Text file containing servers which should be added to the RD Connection Broker server as RD Virtualization Host servers.")] 
    [string]$HostFileName, 
     
    [Parameter(Mandatory=$FALSE, HelpMessage="Text file containing the names of pools that need to be created on the RD Connection Broker server.")] 
    [string]$PoolFileName, 
     
    [Parameter(Mandatory=$FALSE, HelpMessage="Credentials that have write acces on the Active Directory for Personal Virtual Desktop assignment.")] 
    [System.Management.Automation.PSCredential]$Credential, 
     
    [Parameter(Mandatory=$FALSE, HelpMessage="The script seeks confirmation from the user in case it is reassigning a virtual machine to another user unless this switch is used. Providing this switch would force continue the reassignment.")] 
    [Switch]$Force 
) 
 
function Print-Error([System.String]$Msg) 
{ 
    Write-Host $Msg 
    Exit(1) 
} 
 
function Ask-Confirmation() 
{ 
    Write-Warning "The script will not check for current user assignment of Virtual Machines. A virtual machine already assigned to a user might be reassigned to a new user depending on the entries in the 'AssignmentFile'." 
    Write-Host  "Do you want to continue? : [Y]Yes, [N]No" 
    while($TRUE) 
    { 
        $Input = Read-Host 
        switch($Input) 
        { 
            "Yes" { return } 
            "No" {Exit(1)} 
            "N" {Exit(1)} 
            "Y" { return } 
        } 
    } 
} 
 
#Check if there are proper privileges to run the script 
$LoggedInUser = New-Object -TypeName System.Security.Principal.WindowsPrincipal -ArgumentList ([System.Security.Principal.WindowsIdentity]::GetCurrent()) 
if(!($LoggedInUser.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))) 
{ 
    Print-Error "Please execute the script with administrative privileges! Exiting..." 
} 
 
Import-Module ServerManager 
$RoleInstalled = (Get-WindowsFeature RDS-Connection-Broker).Installed 
Remove-Module ServerManager 
 
if(!$RoleInstalled) 
{ 
    Print-Error "This script should be run on the RD Connection Broker server! Exiting..." 
} 
 
 
$LineFormat="<VM Name> [POOL/USER] <Pool Name/User Name>" 
 
# File Format 
# <VMName> [POOL/USER] <farm name/user name> 
 
#Add RD Virtualization Host servers if needed 
function Add-VirtualizationHostServer([string]$FileName) 
{ 
    $servers=(Get-Content -Path $FileName) 
    foreach($server in $servers) 
    { 
        New-Item -Path "RDS:\ConnectionBroker\VirtualDesktops\RDVHostServers" -Name $server 
    } 
     
    Start-Sleep -seconds 5 
} 
 
#Create Farms 
function Add-VmFarm([string]$FileName) 
{ 
    $farms=(Get-Content -Path $FileName) 
    foreach($farm in $farms) 
    { 
        New-Item -Path RDS:\ConnectionBroker\VirtualDesktops\Pools -Name $farm -DisplayName $farm 
    } 
} 
 
#Assign function 
function Assign-VirtualMachines([string]$FileName) 
{ 
    if(!$Force.IsPresent) 
    { 
        Ask-Confirmation 
    } 
     
    $i=1 
    foreach($Line in (Get-Content -Path $FileName)) 
    { 
        $tokens = $Line.Trim().Split() 
        if(($tokens) -and ($tokens.Length -eq 3)) 
        { 
            if($tokens[1] -eq "POOL") 
            { 
                New-Item -Path "RDS:\ConnectionBroker\VirtualDesktops\Pools\$($tokens[2])\VirtualMachines" -Name $tokens[0] 
                if(!$?) 
                { 
                    Write-Error "ERROR in Line '$i': Could not create a pool with name $($tokens[0])" 
                } 
            } 
            else 
            { 
                if($tokens[1] -eq "USER") 
                { 
                    if($Credential) 
                    { 
                        Set-PersonalVirtualDesktop -User $tokens[2] -VirtualDesktop $tokens[0] -Credential $Credential 
                    } 
                    else 
                    { 
                        Set-PersonalVirtualDesktop -User $tokens[2] -VirtualDesktop $tokens[0] 
                    } 
                     
                    if($?) 
                    { 
                        $VMObject = Get-VirtualDesktop -User $tokens[2] 
                        Write-Host "Virtual machine: '$($VMObject.Name)' assigned to: '$($tokens[2])'" 
                    } 
                    else 
                    { 
                        Write-Error "ERROR in Line '$i': Could not assign virtual machine: '$($tokens[0])' to '$($tokens[2])'" 
                        Write-Error "You may want to use the script, Verify-VMConfiguration.ps1, to verify the configuration of virtual machine so that it can be used as a virtual desktop in RemoteApp and Desktop Connection." 
                    } 
                } 
                else 
                { 
                    Write-Error "Error in Line $i: Line in unexpected format. Expected Format: $LineFormat" 
                } 
            } 
        } 
        else 
        { 
            if(($tokens) -and ($tokens.Length -gt 0)) 
            { 
                Write-Error "Error in Line $i: Line in unexpected format. Expected Format: $LineFormat" 
            } 
        } 
        $i++ 
    } 
} 
 
Import-Module RemoteDesktopServices 
 
#Add Virtualization Host Servers 
if($HostFileName) 
{ 
    Add-VirtualizationHostServer($HostFileName) 
} 
 
if($PoolFileName) 
{ 
    Add-VmFarm($PoolFileName) 
} 
 
if($AssignmentFileName) 
{ 
    Assign-VirtualMachines($AssignmentFileName) 
} 
 
