$ErrorActionPreferrence = 'Stop'
$tempPath = c:\TEMP
$archivePath = c:\Archive\temp.zip
$Error.Clear()
Try{
  $rawlog = gci tempPath
  If ($rawlog.count -gt 100){
    Compress-Archive -Path $tempPath -DestinationPath $archivePath
  }
}
Catch{
  $Error
}

# Create a zip file with the contents of C:\Stuff\
Compress-Archive -Path C:\Stuff -DestinationPath archive.zip

# Add more files to the zip file
# (Existing files in the zip file with the same name are replaced)
Compress-Archive -Path C:\OtherStuff\*.txt -Update -DestinationPath archive.zip

# Extract the zip file to C:\Destination\
Expand-Archive -Path archive.zip -DestinationPath C:\Destination
