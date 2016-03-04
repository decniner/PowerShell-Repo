<#

Exiting a Function

To exit a function immediately, use the return statement. 
The next function expects a name (including wildcards) and lists all matching processes. 
If no name is specified, the function outputs a warning message and exits the function using return:
#>

function Get-NamedProcess($name=$null) {
    if ($name -eq $null) {
    Write-Host -foregroundcolor Red 'Specify a name!'
    return
}

    Get-Process $name
}
