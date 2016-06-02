function get-diskinfo ($hostname) {
    
    $driveC = Get-CimInstance -ComputerName $hostname Win32_LogicalDisk -Filter "DriveType=3" 
    $driveC | select @{n="Hostname";e={$_.PScomputername}}, 
              @{n="Drive";e={$_.deviceid}},
              @{n="Size(GB)";e={'{0:N2}'-f ($_.size/1gb)}},
              @{n="FreeSpace(GB)";e={'{0:N2}'-f ($_.freespace/1gb)}}
}

function get-uptime ($hostname) {
    
    $lastbootuptime = (gcim -ClassName Win32_OperatingSystem).Lastbootuptime
    $boot = New-TimeSpan -Start $lastbootuptime
    $uptime = $boot.Hours
    "$uptime hours"
}
