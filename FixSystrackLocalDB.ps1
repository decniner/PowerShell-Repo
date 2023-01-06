$ErrorActionPreference = "Stop"

# Define the hostnames
$hostnames = "host1", "host2", "host3", "host4", "host5", "host6", "host7", "host8", "host9", "host10"

# Define the path to the collect.mdb and collect.ldb files
$mdbPath = "C:\Program Files (x86)\Systrack\lsiagent\MDB\"

# Define the path where the transcript will be saved
$transcriptPath = "C:\TEMP"

# Start a transcript
Start-Transcript -Path "$transcriptPath\transcript.txt"

# Loop through each hostname
foreach ($hostname in $hostnames) {
    # Try to stop the lsiagent service on the current host
    Try {
        Stop-Service -Name "lsiagent" -ComputerName $hostname
    }
    Catch {
        # If an error occurs, write the error to the transcript
        Write-Error "Error stopping lsiagent service on $hostname: $($_.Exception.Message)"
    }

    # Try to rename the collect.mdb and collect.ldb files on the current host
    Try {
        Rename-Item -Path "$mdbPath\collect.mdb" -NewName "collect.mdb.bak" -Force
        Rename-Item -Path "$mdbPath\collect.ldb" -NewName "collect.ldb.bak" -Force
    }
    Catch {
        # If an error occurs, write the error to the transcript
        Write-Error "Error renaming collect.mdb and collect.ldb files on $hostname: $($_.Exception.Message)"
    }

    # Try to start the lsiagent service on the current host
    Try {
        Start-Service -Name "lsiagent" -ComputerName $hostname
    }
    Catch {
        # If an error occurs, write the error to the transcript
        Write-Error "Error starting lsiagent service on $hostname: $($_.Exception.Message)"
    }
}

# Stop the transcript
Stop-Transcript
