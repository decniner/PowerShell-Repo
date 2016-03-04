<#
In PowerShell, you will need to use one of the
Registry virtual drives to read from the Windows Registry as it is - not always intuitive. 
Here are two functions to make your life easier.
#>

# Get-RegistryValues lists all values stored in a registry key:

function Get-RegistryValues($key) {
    (Get-Item $key).GetValueNames()
}
Get-RegistryValues HKLM:\software\microsoft\windows\currentversion

# Get-RegistryValue reads any value in any key and returns just the value:


function Get-RegistryValue($key, $value) {
    (Get-ItemProperty $key $value).$value
}
Get-RegistryValue 'HKLM:\Sofware\Microsoft\Windows NT\CurrentVersion' `
  RegisteredOwner
