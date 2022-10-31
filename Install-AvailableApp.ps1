Function Install-AvailableApp ($hostname){
$Applications = (Get-CimInstance -ClassName CCM_Application -Namespace “root\ccm\clientSDK” -ComputerName $hostname)

ForEach ($Application in $Applications) {
$Args = @{EnforcePreference = [UINT32] 0
Id = “$($Application.id)”
IsMachineTarget = $Application.IsMachineTarget
IsRebootIfNeeded = $False
Priority = ‘High’
Revision = “$($Application.Revision)”}
Invoke-CimMethod -Namespace “root\ccm\clientSDK” -ClassName CCM_Application -ComputerName $hostname -MethodName Install -Arguments $Args}
}
