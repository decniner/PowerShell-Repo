$result = @()
$osinfo = get-computerinfo
$csv = get-computerinfo | export-csv .\result.csv -notypeinformation

foreach ($ent in $osinfo){
  If (ent.osversion -eq $source.osversion){
    $myosinfopsobject = New-Object -Type PSCustomObject -Property @{
               'freePercent' = $freePercent
               'freeGB'      = $freeGB
               'system'      = $system
               'disk'        = $disk
             }
 }
}
