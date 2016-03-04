<#
Get-Acl is a convenient Cmdlet to expose NTFS file and folder settings. 
For example, to get a list of ownerships for a folder content, do this:
#>

Dir | Get-Acl

<#
To find out which "Identities" (Users or groups) have specific permissions granted, 
access the Access property of each Acl to see the actual permissions, 
then group by IdentityReference:
#>

dir | Get-Acl | ForEach-Object { $_.Access  } | Group-Object IdentityReference
