clear   
  
#--Change the name with directory you want to visit --#    
$dir_to_look="C:\Users\admin\Desktop"    
  
#--You may change the number of days of your choice --#   
$seven_days_backdate=$(Get-Date).AddDays(-7)    
  
#--Find the files which are modified or created within last 7 days --#    
Get-Childitem $dir_to_look -Recurse | `   
where-object {!($_.psiscontainer)} | `   
where { $_.LastWriteTime -gt $seven_days_backdate } | `   
foreach {  Write-Host "$($_.LastWriteTime) :: $($_.Fullname) "  }   
