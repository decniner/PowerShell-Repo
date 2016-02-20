function Get-ComputerInfo {
    
    begin {}
    
    process {
        [string]$computername = $_
 
        $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computername
        $bios = Get-WmiObject -Class Win32_BIOS -ComputerName $computername
        $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ComputerName $computername

        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $computername
        $obj | Add-Member -MemberType NoteProperty -Name OsVersion -Value $os.caption
        $obj | Add-Member -MemberType NoteProperty -Name FreeRAM -Value $os.FreePhysicalMemory
        $obj | Add-Member -MemberType NoteProperty -Name Status -Value $os.Status
        $obj | Add-Member -MemberType NoteProperty -Name OSArchitecture $os.OSArchitecture
        $obj | Add-Member -MemberType NoteProperty -Name OSSerialNumber $os.SerialNumber
        $obj | Add-Member -MemberType NoteProperty -Name HardwareManufacturer $bios.Manufacturer
        $obj | Add-Member -MemberType NoteProperty -Name HardwareSerialNumber $bios.SerialNumber
        $obj | Add-Member -MemberType NoteProperty -Name BIOSVersion $bios.SMBIOSBIOSVersion
        $obj | Add-Member -MemberType NoteProperty -Name DiskSize $disk.Size
        $obj | Add-Member -MemberType NoteProperty -Name FreeDiskSpace $disk.FreeSpace

        Write-Output $obj

        }

    end {}
    
}

Get-Content .\VM.txt | Get-ComputerInfo | ogv