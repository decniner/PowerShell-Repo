Get-WmiObject -Class Win32_ComputerSystem | ? { $_.rename("new-hostname") }
