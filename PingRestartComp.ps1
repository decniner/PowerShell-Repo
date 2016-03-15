1  function Ping-Address {
2    PROCESS {
3      $ping = $false
4      $results = Get-WmiObject -query `
5      "SELECT * FROM Win32_PingStatus WHERE Address = '$_'"
6      foreach ($result in $results) {
7        if ($results.StatusCode -eq 0) {
8          $ping = $true
9        }
10      }
11      if ($ping -eq $true) {
12        Write-Output $_
13      }
14    }
15  }
16
17  function Restart-Computer {
18    PROCESS {
19      $computer = Get-WmiObject Win32_OperatingSystem -computer $_
20      $computer.Reboot()
21    }
22  }
23
24  Get-Content c:\computers.txt | Ping-Address | Restart-Computer
