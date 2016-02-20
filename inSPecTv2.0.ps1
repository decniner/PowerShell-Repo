#========================================================================
# inSPecT v2.0
# 10/16/2014
# Contact: densan@ms.com ; jelmerl@ms.com
#========================================================================
#----------------------------------------------
#region Application Functions
#----------------------------------------------
function OnApplicationExit{
$script:ExitCode = 0} #Set the exit code for the Packager
#endregion Application Functions
#----------------------------------------------
# Form Function
#----------------------------------------------

function Call_LazyAdminTool{
	Import-Module ActiveDirectory
	$ErrorActionPreference= 'silentlycontinue'

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies#
	#----------------------------------------------
	#region Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formMain = New-Object System.Windows.Forms.Form
	$tabControl = New-Object System.Windows.Forms.TabControl
    $tabPage1 = New-Object System.Windows.Forms.TabPage
	$tabPage2 = New-Object System.Windows.Forms.TabPage
	$panel1 = New-Object System.Windows.Forms.Panel
	$groupTools = New-Object System.Windows.Forms.GroupBox
	$btnLBConsole = New-Object System.Windows.Forms.Button
	$btnRestart = New-Object System.Windows.Forms.Button
	$btnResize = New-Object System.Windows.Forms.Button
	$btnRestore = New-Object System.Windows.Forms.Button
	$btnWinRS = New-Object System.Windows.Forms.Button
	$btnBAQ = New-Object System.Windows.Forms.Button
	$btnFileVersion = New-Object System.Windows.Forms.Button
	$btnCDrive = New-Object System.Windows.Forms.Button
	$btnRA = New-Object System.Windows.Forms.Button
	$btnRDP = New-Object System.Windows.Forms.Button
	$groupInfo = New-Object System.Windows.Forms.GroupBox
	$btnServices = New-Object System.Windows.Forms.Button
	$btnGetError = New-Object System.Windows.Forms.Button
	$btnGetLargeFiles = New-Object System.Windows.Forms.Button
	$btnCatalyst = New-Object System.Windows.Forms.Button
	$btnTopProcess = New-Object System.Windows.Forms.Button
	$btnSystemInfo = New-Object System.Windows.Forms.Button
	$richTextBox = New-Object System.Windows.Forms.RichTextBox
	$btnPing = New-Object System.Windows.Forms.Button
	$txtComputer = New-Object System.Windows.Forms.TextBox
	$menu = New-Object System.Windows.Forms.MenuStrip
	$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuFileConnect = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuView = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewEventVwr = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewServices = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewUser = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewWSUS = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewWSUSReport = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewWSUSUpdate = New-Object System.Windows.Forms.ToolStripMenuItem
	$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
	$cmsProcEnd = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsStartupRemove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAdminAdd = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAdminRemove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAppUninstall = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsSelect = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
	$SB = New-Object System.Windows.Forms.StatusStrip
	$SBPContact = New-Object System.Windows.Forms.ToolStripStatusLabel
	$SBPStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
	$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
	$btnCheck_Addins = New-Object System.Windows.Forms.Button
	$btnC2CLogs = New-Object System.Windows.Forms.Button
	$btnAppHist = New-Object System.Windows.Forms.Button
	$btnAppFinder = New-Object System.Windows.Forms.Button
	$btnLogOnHist = New-Object System.Windows.Forms.Button
	$btnProcessCheck = New-Object System.Windows.Forms.Button
	$btnLicenseQuery = New-Object System.Windows.Forms.Button
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$handler_tabControl_SelectedIndexChanged={
		if($tabControl.SelectedTab -eq $tabPage2){
			$panel1.Controls.Add($richTextBox)
			$tabPage2.Controls.Add($txtComputer)
			$tabPage2.Controls.Add($panel1)
			$tabPage2.Controls.Add($btnCheck_Addins)
			$tabPage2.Controls.Add($btnC2CLogs)
			$tabPage2.Controls.Add($btnAppHist)
			$tabPage2.Controls.Add($btnAppFinder)
			$tabPage2.Controls.Add($btnLogOnHist)
			$tabPage2.Controls.Add($btnProcessCheck)
			$tabPage2.Controls.Add($btnLicenseQuery)
		}
		else{
			$panel1.Controls.Add($richTextBox)
			$tabPage1.Controls.Add($txtComputer)
			$tabPage1.Controls.Add($groupTools)
			$tabPage1.Controls.Add($groupInfo)
			$tabPage1.Controls.Add($panel1)
			$tabPage1.Controls.Add($btnPing)
		}
	}

    $handler_richTextBox_SelectionChanged={
		if ($richTextBox.Text -like "*Catalyst*" -or $richTextBox.Text -like "*SYSTEM INFORMATION*"){
			if ($richTextBox.SelectedText -ne ""){
				$txtComputer.text = $richTextBox.SelectedText.Trim()
			}
		}
	}
	
	$menuViewEventVwr_Click={
		Get-ComputerName
		EventVwr $ComputerName
	}
		
	$menuViewServices_Click={
		Get-ComputerName
		Services.msc /Computer:$ComputerName
	}
	
	$menuViewUser_Click={
		Get-ComputerName
		LUsrMgr.msc /Computer:$ComputerName
	}
			
	$menuHelpAbout_Click={
    	MsgBox "Contact:  abc@xyz.com"
    }	
    
    $menuFileExit_Click={
		$formMain.Close()
	}
		
	$btnPing_Click={
		Get-ComputerName
		Clear_Screen
		$SBPStatus.Text = "Running Ping Check..."
		$IP = $null
		$IP = [System.Net.Dns]::GetHostAddresses("$Computername")|select -expandproperty IPAddressToString
		$richTextBox.Text += "*Pinging $ComputerName [$IP]","`n","`n"
		if ($ComputerName -ne ""){
			if (Test-Connection $ComputerName -quiet -Count 1){
				AppendText $richTextBox Green "$ComputerName is ONLINE."
			}
			else{
				AppendText $richTextBox Red "$ComputerName does NOT have network connection"
			}
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnSystemInfo_Click={
		Get-ComputerName
		Clear_Screen
		$SBPStatus.Text = "Checking System Information..."
		if ($ComputerName -ne ""){
			if (Test-Connection $ComputerName -quiet -Count 1){
				$richTextBox.Text += System_Info | Format-Table -AutoSize | Out-String
				Monitor_Check
			}
			else{
				AppendText $richTextBox Red "$ComputerName is inaccessible or does NOT have network connection"
			}
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnCatalyst_Click={
		Clear_Screen
		if ($txtComputer.Text -eq ""){ 
			$richTextBox.Text += "`n","Please enter search keyword (i.e., desk location, machine owner or PC name)" 
		}
		else{
			$SBPStatus.Text = "Checking catalyst records..."
			$UserInput = $txtComputer.Text
			$CatResult = Import-Csv "\\msad\root\JP\TK\lib\it\eus\Tools\CAT\catalyst.csv" | Select Location, Machine_Name, Machine_Owner | Where-Object {$_.Machine_Owner -like "$UserInput" -or $_.Machine_Name -like "$UserInput" -or $_.Location -like "*$UserInput*"} | format-table -AutoSize | Out-String
			$CatDate = Get-Date
			$richTextBox.Text += "*Catalyst Report on",$CatDate.ToShortDateString(),"`n","$CatResult"
			$progressBar.Value = 100
		}
		if ($CatResult -eq ""){
			$richTextBox.Text = ""
			AppendText $richTextBox Red "`n","Sorry, no matching results found!" 
		}
		$richTextBox.add_SelectionChanged($handler_richTextBox_SelectionChanged)
		$SBPStatus.Text = "Ready"
	}

	$btnTopProcess_Click={
		Get-ComputerName
		Clear_Screen
		if (Test-Connection $ComputerName -quiet -Count 1){
			$SBPStatus.Text = "Checking running processes..."
			$richTextBox.Text += "Processes running on $ComputerName...","`n"
			$richTextBox.Text += Get-Process -ComputerName $ComputerName | 
			Sort WS -desc | 
			Select Name,ID,StartTime,@{n="CPU(%)";e={$_.cpu/1000 -as [int]}},@{n="MEM Usage (MB)";e={$_.ws/1MB -as [int]}} | 
			Format-Table -AutoSize | 
			Out-String
		}
		else{
			AppendText $richTextBox Red "$ComputerName is inaccessible or does NOT have network connection"
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnGetError_Click={
		Get-ComputerName
		Clear_Screen
		if (Test-Connection $ComputerName -quiet -Count 1){
			$SBPStatus.Text = "Checking for errors..."
			$richTextBox.Text += "List of Errors on $ComputerName...","`n"
			$richTextBox.Text += Get-EventLog -ComputerName $ComputerName -LogName Application -EntryType Error,Warning -Newest 200 |
			Where-object {$_.InstanceID -ne 3221225473}|
			select TimeWritten, Source, Message| 
			fl|
			Out-String
		}
		else {
			AppendText $richTextBox Red "$ComputerName is inaccessible or does NOT have network connection"
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnFileVersion_click={
		Clear_Screen
		$objExcel = new-object -comobject excel.application  
		$objExcel.Visible = $True  
		$objWorkbook = $objExcel.Workbooks.Open("\\msad\root\JP\TK\lib\it\eus\TOOLS\Lazy.Admin.Tool\Win7FileVersionCheck.xls") 
  	}
	
	$btnBAQ_click={
		Clear_Screen
		Invoke-Item -Path "\\msad\root\JP\TK\lib\it\eus\Tools\Lazy.Admin.Tool\BAQ.xls"
	}
  
  	$btnGetLargeFiles_Click={
		Clear_Screen
		Get-ComputerName
		$RealValue = $null
		if (Test-Connection $ComputerName -quiet -Count 1){
			$SBPStatus.Text = "Files are being scanned, please wait..."
			$WhoIsLoggedOn = Invoke-Command -Computername $ComputerName {quser}
			$TestMatch = $WhoIsLoggedOn[1] -match "[a-zA-Z0-9_]+"
			if ($TestMatch -eq $False){
				$TestMatch = $WhoIsLoggedOn[4] -match "[a-zA-Z0-9_]+"
			}
			$RealValue = $matches.values
			if ($ComputerName.Text -ne ""){
				
				$Path1 = "\\$ComputerName\C$\Users\$RealValue\AppData\Local\Temp"
				$Path2 = "\\$ComputerName\C$\Windows\Temp"
				$Path3 = "\\$ComputerName\C$\Users\$RealValue\AppData\Local\Microsoft\Outlook" 
				$Path4 = "\\$ComputerName\C$"
				$files = get-childitem -path "$Path1", "$Path2", "$Path3", "$Path4"
			 	$files | % -begin { $i=0 } -process {($i/$files.count*100); $i+=1 ;$progressBar.value = ($i/$files.count*100)}
				$TopFiles = Get-ChildItem -Path "$Path1" -Recurse | 
				Sort-Object Length -Descending |
				Select-Object Name, @{Name="Size (MB)";Expression={"{0:N2}" -f ($_.Length/1mb)}} -First 5
				$TopFiles2 = Get-ChildItem -Path "$Path2" -Recurse |
				Sort-Object Length -Descending |
				Select-Object Name, @{Name="Size (MB)";Expression={"{0:N2}" -f ($_.Length/1mb)}} -First 5
				$TopFiles3 = Get-ChildItem -Path "$Path3" |
				Sort-Object Length -Descending |
				Select-Object Name, @{Name="Size (MB)";Expression={"{0:N2}" -f ($_.Length/1mb)}} -First 5 
				$TopFiles4 = Get-ChildItem -Path "$Path4" |
				Sort-Object Length -Descending |
				Select-Object Name, @{Name="Size (MB)";Expression={"{0:N2}" -f ($_.Length/1mb)}} -First 5
				$FileCount = Get-ChildItem -Path "\\$ComputerName\C$\Users\$RealValue\AppData\Local\Temp" -Recurse -Force | 
				measure
				$richTextBox.Text = "Found a total of " + $FileCount.count + " temporary files on " + $ComputerName `
				+ "`n" + "----------------------------------------------------------------------------" + "`n" `
				+ "`n" + "Top files according to size:" + "`n" + "`n" + "Directory: " + "\\$ComputerName\C$\Users\$RealValue\AppData\Local\Temp" + "`n"
				$richTextBox.Text += $TopFiles | 
				Format-Table | 
				Out-String
				$richTextBox.Text += "`n" + "`n" + "Directory: " + "\\$ComputerName\C$\Windows\Temp" + "`n"
				$richTextBox.Text += $TopFiles2 | ft -HideTableHeaders | 
				Out-String
				$richTextBox.Text += "`n" + "`n" + "Directory: " + "\\$ComputerName\C$\Users\$RealValue\AppData\Local\Microsoft\Outlook" + "`n"
				$richTextBox.Text += $TopFiles3 | ft -HideTableHeaders | 
				Out-String	
				$richTextBox.Text += "`n" + "`n" + "Directory: " + "\\$ComputerName\C$" + "`n"
				$richTextBox.Text += $TopFiles4 | ft -HideTableHeaders | 
				Out-String	
			}
		}
		else{
			$richTextBox.Text = ""
			AppendText $richTextBox Red "$ComputerName is inaccessible or does NOT have network connection"
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnCDrive_Click={
		Get-ComputerName
		$ViewCDrive = "\\$ComputerName\C$"
		Explorer.exe $ViewCDrive
	}
	
	$btnWinRS_Click={
        Get-ComputerName
        Start-Process "cmd" -ArgumentList " /k winrs -r:$ComputerName cmd /K cd %binpath%\SoftwareDelivery\MyInstaller"
	}
	
    $btnServices_Click={
		Get-ComputerName
		Services.msc /Computer:$ComputerName
	}
    
    $btnRDP_Click={
	   	Get-ComputerName
	   	MSTSC.exe /v:$ComputerName
	}
    
    $btnRA_Click={
       	Get-ComputerName
        $batchPath = "\\msad\ROOT\JP\TK\lib\it\eus\Tools\Lazy.Admin.Tool\Remote_Assist.bat"
		Clear-Content "$batchPath"
		Add-Content -Path "$batchPath" "@ECHO Running Remote Assistance", 'C:\Windows\System32\runas.exe /savecred /user:msad\%USERNAME:_sup=% "cmd /c C:\Windows\System32\msra.exe /offerra computername"'
		(Get-Content -Path "$batchPath") | 
		foreach-object {$_ -replace "computername", "$ComputerName"} | 
		Set-Content -Path "$batchPath"
		Start-Process "$batchPath"
	}
	
	$btnRestart_Click={
		Get-ComputerName
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | 
		Out-Null
		$VBMsg = New-Object -COMObject WScript.Shell
		$VBRestart = $VBMsg.popup("Are you sure you want to restart " + $ComputerName.ToUpper() + "?",0,"Restart " + $ComputerName.ToUpper() + "?",4)
		Switch ($VBRestart){
			6 {Restart-Computer -Force -Computername $ComputerName}
			7 {}
		}
	}	
	
	$btnResize_Click={
		$formMain.ClientSize = '148, 646'
	}	
	
	$btnRestore_Click={
		$formMain.ClientSize = '780, 646'
	}
	
	$btnLBConsole_Click={
		Start-Process "\\msad\root\JP\TK\lib\it\eus\Tools\LBConsole\Densan.BAT"
	}
	
	$btnCheck_Addins_Click={
		Get-ComputerName
		Clear_Screen
		$UserName = $null
		if ($ComputerName -ne ""){
			if (Test-Connection $ComputerName -quiet -Count 1){
				$SBPStatus.Text = "Checking for known add-in issues..."
				$ODEInstalledCheck = Get-MSRemoteRegistry -computername $ComputerName -path "HKLM\SOFTWARE\Wow6432Node\ABC\ODE" | Where {$_.Name -eq "Excel_version"} 
				if ($ODEInstalledCheck -eq $Null){
					AppendText $richTextBox Red "-> Excel ODE does not seem to be installed. Please install it via MyApps."
				}
				else{
					$z = Get-Process -ComputerName $ComputerName -Name Excel
					If ($z -eq $Null){
						AppendText $richTextBox Red "-> Please make sure Excel is running on the target machine before running this tool.","`n","`n"
					}	
					else{	
						$richTextBox.Text = "Checking for known issues. Please wait...","`n","`n" 
						$SBPStatus.Text = "Checking for known add-in issues..."
						$ComputerSystem = Get-WMIObject -computername $ComputerName -class Win32_ComputerSystem
						$WhoIsLoggedOn = Invoke-Command -Computername $ComputerSystem.Name {quser}
						$TestMatch = $WhoIsLoggedOn[1] -match "[a-zA-Z0-9_]+"
						if ($TestMatch -eq $False){
							$TestMatch = $WhoIsLoggedOn[4] -match "[a-zA-Z0-9_]+"
						}
						$UserName = $matches.values
						$User = New-Object System.Security.Principal.NTAccount("$UserName")
						$SID = $User.Translate([System.Security.Principal.SecurityIdentifier])
						$xlentAddinLocate = Get-MSRemoteRegistry -computername $ComputerName -path HKEY_USERS\$SID\Software\Microsoft\Office\Excel\AddIns\denStanley.DeskApps.XLent.XLentODEPlugin | 
						select Name, Data | 
						where {$_.name -eq "LoadBehavior"}
						Start-Sleep -m 1000
						$progressBar.Value = 10
						if ($xlentAddinLocate.Data -eq $null){
							AppendText $richTextBox Red "-> ODE is not loaded","`n"
						}
						else {
							AppendText $richTextBox Green "-> ODE status is Green!","`n","`n"
							Start-Sleep -m 1000
							$progressBar.Value = 20
							AppendText $richTextBox Black "Checking Com Add-ins...","`n"
							Start-Sleep -m 1000
							$progressBar.Value = 30
							AppendText $richTextBox Green "The following Com add-ins are loaded in excel:","`n","`n"
							$Addins_loaded = Get-MSRemoteRegistry -computername $ComputerName -path `
							HKEY_USERS\$SID\Software\Microsoft\Office\Excel\AddIns `
							| Select Name | ft -HideTableHeaders | Out-String
							AppendText $richTextBox Green "$Addins_loaded"
							$progressBar.Value = 40
						} 
						$y = Get-MSRemoteRegistry -computername $ComputerName `
						-path HKEY_USERS\$SID\Software\Microsoft\Office\Excel\AddIns\ `
						| Select name | Where {$_.name -eq "zzdenStanley.Deskapps.ODE.MSODEHost_Excel"}
						if ($y.Name -eq "zzdenStanley.Deskapps.ODE.MSODEHost_Excel"){
							AppendText $richTextBox Red "-> ODE won't load in Excel during start-up","`n","`n",`
							"Please run the below command via command prompt then restart excel:","`n",`
							"REG DELETE HKEY_USERS\$SID\Software\Microsoft\Office\Excel\AddIns\zzdenStanley.Deskapps.ODE.MSODEHost_Excel /f","`n","`n"
						}
						$Edis = Get-MSRemoteRegistry -computername $ComputerName -path HKEY_USERS\$SID\Software\Microsoft\Office\14.0\Excel\Resiliency\DisabledItems
						$Odis = Get-MSRemoteRegistry -computername $ComputerName -path HKEY_USERS\$SID\Software\Microsoft\Office\14.0\Outlook\Resiliency\DisabledItems
						$Wdis = Get-MSRemoteRegistry -computername $ComputerName -path HKEY_USERS\$SID\Software\Microsoft\Office\14.0\Word\Resiliency\DisabledItems
						$Pdis = Get-MSRemoteRegistry -computername $ComputerName -path HKEY_USERS\$SID\Software\Microsoft\Office\14.0\Powerpoint\Resiliency\DisabledItems
						AppendText $richTextBox Black "Looking for disabled add-ins...","`n","`n"
						$progressBar.Value = 50
						Start-Sleep -m 1000
						if ($Edis.Name){
							AppendText $richTextBox Red "-> Found some disabled item/s in Excel. Please check.","`n"
						}
						if ($Odis.Name){
							AppendText $richTextBox Red "-> Found some disabled item/s in Outlook. Please check.","`n"
						}
						if ($Wdis.Name){
							AppendText $richTextBox Red "-> Found some disabled item/s in Word. Please check.","`n"
						}
						if ($Pdis.Name){
							AppendText $richTextBox Red "-> Found some disabled item/s in Powerpoint. Please check.","`n"
						}
						
						if (!$Edis.Name -and !$Odis.Name -and !$Wdis.Name -and !$Pdis.Name){
							AppendText $richTextBox Green "-> All Green."
						}
					}	
				}				
			}
			else{
				AppendText $richTextBox Red "$ComputerName is inaccessible or does NOT have network connection"
			}
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnC2CLogs_Click={
	
	function New-Zip {
		param([string]$zipfilename)
		set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		(dir $zipfilename).IsReadOnly = $false
		}

#		#usage: new-zip c:\demo\myzip.zip
		function Add-Zip {
		param([string]$zipfilename)

		if(-not (test-path($zipfilename)))
		{
			set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
			(dir $zipfilename).IsReadOnly = $false	
		}
	
		$shellApplication = new-object -com shell.application
		$zipPackage = $shellApplication.NameSpace($zipfilename)
	
		foreach($file in $input) 
		{ 
            $zipPackage.CopyHere($file.FullName)
            Start-sleep -milliseconds 500
		}
		}

	#usage: dir c:\demo\files\*.* -Recurse | add-Zip c:\demo\myzip.zip
	
		Get-ComputerName
		Clear_Screen
		$UserName = $null
		$matches = $null
		$SBPStatus.Text = "Running C2C Log Grabber..."
		if (Test-Connection $ComputerName -quiet -Count 1){
			try {
				$ComputerSystem = Get-WMIObject -computername $ComputerName -class Win32_ComputerSystem
				$WhoIsLoggedOn = Invoke-Command -Computername $ComputerSystem.Name {quser}
				$TestMatch = $WhoIsLoggedOn[1] -match "[a-zA-Z0-9_]+"
				if ($TestMatch -eq $False){
					$TestMatch = $WhoIsLoggedOn[4] -match "[a-zA-Z0-9_]+"
				}
				$UserName = $matches.values
				$ConvertToText = $matches | Out-String
				$C2DSourceDir = "\\$ComputerName\C$\ProgramData\den jelmer\ShimC2DService\"
				$C2DDestDir = New-Item -ItemType directory -Path "\\offidbackup\temp\C2C_logs\$UserName\$ComputerName\" -force
				$LatestC2DLogs = Get-ChildItem -Path $C2DSourceDir -ErrorAction Stop | where-object {$_.lastwritetime -gt (get-date).addDays(-7)} | Foreach-Object { $_.FullName } 
				
				#Below code will create 3 zip files in the $C2DDestDir folder
				New-Zip $C2DDestDir\LocalStore.zip 
				New-Zip $C2DDestDir\ShimC2DServiceLogs.zip
				New-Zip $C2DDestDir\Click2Dial.zip
#				New-Zip $C2DDestDir\BlastMail.zip
				
				for($i=0; $i -le $LatestC2DLogs.count -1; $i++){
					$C2DLogPath = $LatestC2DLogs[$i] 	
					dir $C2DLogPath -Recurse | Add-Zip $C2DDestDir\LocalStore.zip 	
					Start-Sleep -s 2
				}

				$CRMSourceDir = "\\$ComputerName\C$\Users\$UserName\AppData\Roaming\CRMClientTracker\Local Store\"
				$LatestCRMLogs = Get-ChildItem -Path $CRMSourceDir -ErrorAction Stop | where-object {$_.lastwritetime -gt (get-date).addDays(-7)} | Foreach-Object { $_.FullName } 
					
				for($i=0; $i -le $LatestCRMLogs.count -1; $i++){
					$CRMLogPath = $LatestCRMLogs[$i] 	
					dir $CRMLogPath -Recurse | Add-Zip $C2DDestDir\ShimC2DServiceLogs.zip
					Start-Sleep -s 2
				}
				
#				$BMSourceDir = "\\$ComputerName\C$\SysAdmin\Log\OutlookBlastmail\$UserName\"
#				$LatestBMLogs = Get-ChildItem -Path $BMSourceDir -ErrorAction Stop | where-object {$_.lastwritetime -gt (get-date).addDays(-7)} | Foreach-Object { $_.FullName } 
#					
#				for($i=0; $i -le $LatestBMLogs.count -1; $i++){
#					$BMLogPath = $LatestBMLogs[$i] 	
#					dir $BMLogPath -Recurse | Add-Zip $C2DDestDir\BlastMail.zip
#					Start-Sleep -s 2
#				}
				
				AppendText $richTextBox Green "Latest logs are saved under \\offidbackup\temp\C2C_logs\$UserName\$ComputerName\"
				
			}	
			
			catch {
  				[system.exception]
				AppendText $richTextBox Red "-> Unable to locate the log files.", "`n"
			}
		}	
		else{
			AppendText $richTextBox Red "$ComputerName does NOT have network connection."
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100
	}
	
	$btnAppHist_Click={
		Get-ComputerName
		Clear_Screen
		$richTextBox.text += "Recent Application Installation History on $ComputerName","`n"
		$richTextBox.text += "===================================================================","`n"
		$richTextBox.text += get-eventlog -computername $ComputerName -logname application -source msiinstaller -InstanceId 1033 -Newest 20 |
				Select @{n="InstalledOn";e={$_.TimeWritten}}, 
				   @{n="Application Name";
				   	 e={ 
						  $Result = %{
							$_.Message -match "(?<=Windows Installer installed the product. Product Name: )(.*)(?=. Product language)"
						  }
							$matches[0]
					  } 
					} | 
			ft -AutoSize |
			Out-String
		$richTextBox.text += get-eventlog -computername $ComputerName -logname application -source msiinstaller -InstanceID 1034 -Newest 20 | 
			Select @{n="UninstalledOn";e={$_.TimeWritten}}, 
				   @{n="Application Name";
				   	 e={ 
						  $Result2 = %{
							$_.Message -match "(?<=Windows Installer removed the product. Product Name: )(.*)(?=. Product language)"
						  }
							$matches[0]
					  } 
					} | 
			ft -AutoSize |
			Out-String
		
		$richTextBox.text += "Script Used","`n"
		$richTextBox.text += "-----------","`n"
		$richTextBox.text += get-eventlog -computername $ComputerName -logname application -source msiinstaller -InstanceId 1040 | 
			Select message |
			Where {% {$_.Message -like "*\\*"}} | 
			%{
			$x = $_.Message -match "(?<=Beginning a Windows Installer transaction: )(.*)(?=source)"
			$matches[0]
			} | Select -First 20 |
			Out-String
			
#		$richTextBox.text += get-hotfix -ComputerName $ComputerName | 
#			Sort InstalledOn -Descending | 
#			Select InstalledOn, HotFixID, Description -First 10 | 
#			ft -AutoSize |
#			Out-String		
	}	
		
	$btnAppFinder_Click= {
		Get-ComputerName
		Clear_Screen
		$SBPStatus.Text = "Running App Finder..."
		if (Test-Connection $ComputerName -quiet -Count 1){
			Add-Type -AssemblyName Microsoft.VisualBasic
			$App = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Application name', 'Application Name')
			$RichTextBox.Text = "Searching for $App on $ComputerName, please wait...", "`n" 
			
			if ($App){
				[Array]$CountItems = Get-ADComputer $ComputerName | Get-InstalledSoftware | where { $_.Name -like "*$App*" }
				$AppFound = $CountItems | ft -AutoSize | Out-String
				$RichTextBox.Text = "Found " + $CountItems.Count + " entrie/s of $App on $ComputerName.","`n" 
				$RichTextBox.Text += $AppFound
				
				if (!$AppFound){
					Clear_Screen
					AppendText $richTextBox Red "Cannot find $App on $ComputerName."
				}
			}
			
			else{
				Clear_Screen
				$RichTextBox.Text = "You did not enter a valid application name."
			}
		}	
			
		else{
			AppendText $richTextBox Red "$ComputerName does NOT have network connection."
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100	
	}
	
	$btnLogOnHist_Click= {
		Get-ComputerName
		Clear_Screen
		$SBPStatus.Text = "Checking Log On History..."
		if (Test-Connection $ComputerName -quiet -Count 1){
			Add-Type -AssemblyName Microsoft.VisualBasic
			$NoOfDays = [Microsoft.VisualBasic.Interaction]::InputBox('Enter number of days', 'No. of Days')
			
			if ($NoOfDays){
				Get-Logonhistory -Computer $ComputerName -Days "$NoOfDays"

			else{
					Clear_Screen
					$RichTextBox.Text = "You did not enter a value for no. of days."
				}
			}	
		}	
			
		else{
			AppendText $richTextBox Red "$ComputerName does NOT have network connection."
		}
		$SBPStatus.Text = "Ready"
		$progressBar.Value = 100	
	}
	$btnProcessCheck_Click={
	Clear_Screen
	function Get-ProcessStatus {
	<#
  .SYNOPSIS
   	GetProcessStatus will check the status of processes
  .DESCRIPTION
    Check Process on a list of computers.
	The function will show the computername, logged on user, Process, Start Time, Uptime

  .EXAMPLE
  Get-ProcessStatus -ComputerName OFITDENSAN7
  .EXAMPLE
  Get-ProcessStatus -ComputerName OFITDENSAN7, OFIT275728D
  .PARAMETER computername
  OFIT275728D
  .PARAMETER logname
  The name of a file to write failed computer names to. Defaults to errors.txt.
#>
  	[CmdletBinding()]
  	param
  	(
    	[Parameter(Mandatory=$False,
    	ValueFromPipeline=$True,
    	ValueFromPipelineByPropertyName=$True)]
    	[string[]]$ComputerName
  	)

  	begin {
      $ErrorActionPreference= 'silentlycontinue'
  	}
  
  	process { 
		$ComputerName =  ""
    	Add-Type -AssemblyName Microsoft.VisualBasic
		$ComputerName = [Microsoft.VisualBasic.Interaction]::InputBox('Insert machine list in txt format', 'ComputerName')
		$ComputerName = $ComputerName -replace '[""]',''
		$Process = [Microsoft.VisualBasic.Interaction]::InputBox('Enter process name', 'Process Name')
		$Process = $Process.Trim()
#     $ComputerName = "OFFID291674W"
#     $ComputerName = Get-Content "\\offidbackup\shared\Jelmer\fidtraders.txt"
		if ($ComputerName -like "*.txt*"){
		$ComputerName = Get-Content -Path "$ComputerName"}
		else {$ComputerName = $txtComputer.Text}	
		$hash = @()
      	foreach ($computer in $ComputerName) {
                   $C2CStatus = ""
                   $ShimC2DStatus = ""
                   $user = ""
                   $StartTime = ""
                   $Uptime = ""
                   $days = ""
                   $hours = ""
                   $minutes = ""
				   $Version = ""
				   $path = ""
            $Ping = Test-Connection -ComputerName $Computer -Quiet -Count 1
            if ($Ping -eq "True"){ 
				$EmptyString = "" |Out-String
                $computer = $computer.Trim()
                $user = (Get-WmiObject Win32_Process -ErrorAction SilentlyContinue -Computername $Computer -filter 'name="explorer.exe"').GetOwner() | Select Domain,User |select -ExpandProperty User
                $ProcessName = gps -ComputerName $Computer -Name $Process*  |select -expandproperty ProcessName
				if ($ProcessName.count -lt 2){
				$ProcessName = $ProcessName}
				else {$ProcessName = $ProcessName[0]}
                #$ShimC2DStatus = invoke-command -ComputerName $computer {get-service -ErrorAction SilentlyContinue -DisplayName ShimC2DService}  |select -ExpandProperty Status
                         if ($ProcessName.length -gt 2){
                         $StartTime = gwmi win32_process -computername $computer| ? { $_.name -like "$processName*"}
						 $Version = Invoke-Command -ComputerName $computer -ScriptBlock {param($ProcessName) get-process | Where{$_.Name -like "*$ProcessName*"} | select -ExpandProperty FileVersion} -ArgumentList $processName
                         	if ($StartTime -ne ""){
						 			$StartTime =  $StartTime | % { $_.ConvertToDateTime( $_.CreationDate )}
									$StartTime = $StartTime[0]
                                    $Uptime = (Get-Date) - $StartTime
                                    $days = $Uptime |select -ExpandProperty Days 
                                    $hours = $Uptime |select -ExpandProperty Hours
                                    $minutes = $Uptime |select -ExpandProperty Minutes
                                    $Uptime = "$days days $hours hours $minutes min"
									} 
                         }
                         else {$ProcessName = "Process $process not found"}
                        # if ($ShimC2DStatus.length -gt 5){$ShimC2DStatus = $ShimC2DStatus}
                        # else {$ShimC2DStatus = "ShimC2D Not running"}
                         
                   		$hashProperties = @{  
                         "ComputerName"                 = $Computer
                         "User"                         = $user
                         #"ShimC2D_Status"               = $ShimC2DStatus 
                         "Process"               		= $ProcessName
						 "Path"               			= $Path
                         "StartTime"                    = $StartTime
                         "Uptime"                 		= $Uptime
						 "Version"                 		= $Version
                         } 
            $HashObject = New-Object PSObject -Property $hashProperties
            $hash = $hash += $HashObject 
			$HashTracker = $hash | Out-String
			$RichTextBox.Text += "Completed scan on $Computer....$EmptyString"
            }  
			else {
				$EmptyString = "" |Out-String
				$RichTextBox.Text += "$computer not pingable $EmptyString"
				$ProcessName = "Machine not pingable"
				$hashProperties = @{  
                         "ComputerName"                 = $Computer
						 "Process"               		= $ProcessName
						 "Version"               		= $Error}
				$HashObject = New-Object PSObject -Property $hashProperties
            	$hash = $hash += $HashObject 
			}
          }
		  $EmptyString = "" |Out-String
		  $richTextBox.text += "$EmptyString"
		  $RichTextBox.Text += "Check the gridview for results $EmptyString"
		  $hash | select ComputerName, User, Process, Version, Path, StartTime, Uptime |sort Process |Out-GridView
     
      #$hash | select ComputerName, User, ShimC2D_Status, C2CPro_Process, StartTime, Uptime |sort User |Export-Csv -Path \\offidbackup\shared\jelmer\c2c.csv -Encoding ascii -NoTypeInformation
      }
	}
	Get-ProcessStatus
	}
	
	$btnLicenseQuery_Click={
		Get-ComputerName
		Clear_Screen
		$Line = ""
		$username = ""
		$EmptyString = "" |Out-String
		$user = $ComputerName
		$UserName = Get-ADUser $user -Properties * |select -expandproperty mscn
		$richTextBox.text += "Quering Licenses for user: $UserName ($user)  $emptystring $emptystring"
		$richTextBox.text += "RESULTS: $emptystring"
		$UserLicenses = Get-ADUser $user -Properties * |select -expandproperty MemberOf
		if ($username.Length -gt 2){
			foreach ($Line in $UserLicenses) {
			$Bool1 = $Line.Contains("ww-")
			$Bool2 = $Line.Contains("-lic")
			if ($Bool1 -eq "True" -and $Bool2 -eq "True"){
				$Line =$Line.Replace("CN=","")
				$Line = $Line -replace ",.*" -replace ",.*"
				$richTextBox.text += $Line + $EmptyString}
			}
			
		}
		else {$richTextBox.text += "Unable to find user $user"}
	}
		
	$SBPContactLink_Click={
		$Email = "mailto:tksatfls@denstanley.com"
		Start $Email
	}
		
	
	
	# --End User Script--
	#----------------------------------------------
	#region Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load={
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formMain.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed={
		#Remove all event handlers from the controls
		try {
			$btnLBConsole.remove_Click($btnLBConsole_Click)
			$btnRestart.remove_Click($btnRestart_Click)
			$btnResize.remove_Click($btnResize_Click)
			$btnRestore.remove_Click($btnRestore_Click)
			$btnWinRS.remove_Click($btnWinRS_Click)
			$btnBAQ.remove_Click($btnBAQ_Click)
			$btnFileVersion.remove_Click($btnFileVersion_Click)
			$btnCDrive.remove_Click($btnCDrive_Click)
			$btnRA.remove_Click($btnRA_Click)
			$btnRDP.remove_Click($btnRDP_Click)
			$btnServices.remove_Click($btnServices_Click)
			$btnGetError.remove_Click($btnGetError_Click)
			$btnGetLargeFiles.remove_Click($btnGetLargeFiles_Click)
			$btnCatalyst.remove_Click($btnCatalyst_Click)
			$btnTopProcess.remove_Click($btnTopProcess_Click)
			$btnSystemInfo.remove_Click($btnSystemInfo_Click)
			$btnPing.remove_Click($btnPing_Click)
			$btnCheck_Addins.remove_Click($btnCheck_Addins_Click)
			$btnC2CLogs.remove_Click($btnC2CLogs_Click)
			$btnAppHist.remove_Click($btnAppHist_Click)
			$btnAppFinder.remove_Click($btnAppFinder_Click)
			$btnLogOnHist.remove_Click($btnLogOnHist_Click)
			$btnProcessCheck.remove_Click($btnProcessCheck_Click)
			$btnLicenseQuery.remove_Click($btnLicenseQuery_Click)
			
			$formMain.remove_Load($formMain_Load)
			$menuFileConnect.remove_Click($menuFileConnect_Click)
			$menuFileExit.remove_Click($menuFileExit_Click)
			$menuViewEventVwr.remove_Click($menuViewEventVwr_Click)
			$menuViewServices.remove_Click($menuViewServices_Click)
			$menuViewUser.remove_Click($menuViewUser_Click)
			$menuViewWSUSReport.remove_Click($menuViewWSUSReport_Click)
			$menuViewWSUSUpdate.remove_Click($menuViewWSUSUpdate_Click)
			$cmsProcEnd.remove_Click($cmsProcEnd_Click)
			$cmsStartupRemove.remove_Click($cmsStartupRemove_Click)
			$cmsAdminAdd.remove_Click($cmsAdminAdd_Click)
			$cmsAdminRemove.remove_Click($cmsAdminRemove_Click)
			$cmsAppUninstall.remove_Click($cmsAppUninstall_Click)
			$cmsSelect.remove_Click($cmsSelect_Click)
			$menuHelpAbout.remove_Click($menuHelpAbout_Click)
			$formMain.remove_Load($Form_StateCorrection_Load)
			$formMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Events
	#----------------------------------------------
	#region Form Code
	#----------------------------------------------
	#
	# formMain
	#
	#$Image = [system.drawing.image]::FromFile("$($Env:Public)\Pictures\Sample Pictures\Penguins.jpg")
	#$formMain.BackgroundImage = $Image
	#$formMain.BackgroundImageLayout = "None"
	#$formMain.Width = $Image.Width
	#$formMain.Height = $Image.Height
	#$formMain.TopMost = $true
	#$formMain.AutoScale = $true
	#$File = (get-item '\\Msad\root\JP\TK\lib\it\eus\TOOLS\Lazy.Admin.Tool\grey-wallpaper-36.jpg')
	#$Image = [System.Drawing.Image]::Fromfile($file);
	#$tabPage1.BackgroundImage = $Image
	#$tabPage1.BackgroundImageLayout = "None"
	#$tabPage1.Width = $Image.Width
	#$tabPage1.Height = $Image.Height
	#$tabPage1.TopMost = $true
	#$tabPage1.AutoScale = $true
	$formMain.Controls.Add($tabControl)
	$formMain.Controls.Add($SB)
	$formMain.Controls.Add($menu)
	#$formMain.ClientSize = '822, 696'
	#$formMain.ClientSize = '808, 661'
	$formMain.ClientSize = '814, 635'
	$formMain.MainMenuStrip = $menu
	$formMain.Name = "formMain"
	$formMain.StartPosition = 'CenterScreen'
	$currentUser = $env:username.tolower()
	$formMain.Text = "inSPecT v2.0 running as $currentUser ¯\_(ツ)_/¯"
	#$formMain.AutoSize = $True
	#$formMain.AutoSizeMode = "GrowOnly"
	#$formMain.FormBorderStyle = "Fixed3D"
	$formMain.SizeGripStyle = "Hide"
	$formMain.add_Load($formMain_Load)
	#
	# tabControl
	#
	$tabControl.Controls.Add($tabPage1)
	$tabControl.Controls.Add($tabPage2)
	$tabControl.Anchor = 'Top, Bottom, Left, Right'
	$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
	$tabControl.SelectedIndex = 0
	$tabControl.TabIndex = 0
	$tabControl.Name = 'tabControl1'
	$tabControl.Location = '4, 34'
	#$tabControl.Size = '813, 634'
	$tabControl.Size = '809, 575'
	$tabControl.add_SelectedIndexChanged($handler_tabControl_SelectedIndexChanged)
	#
	# tabPage1
	#
	$panel1.Controls.Add($richTextBox)
	$tabPage1.Controls.Add($txtComputer)
	$tabPage1.Controls.Add($groupTools)
	$tabPage1.Controls.Add($groupInfo)
	$tabPage1.Controls.Add($btnPing)
	$tabPage1.Controls.Add($panel1)
	$tabPage1.Controls.Add($btnLBConsole)
	$tabPage1.Controls.Add($btnRestart)
	$tabPage1.Controls.Add($btnWinRS)
	$tabPage1.Controls.Add($btnBAQ)
	$tabPage1.Controls.Add($btnFileVersion)
	#$tabPage1.Controls.Add($btnResize)
	#$tabPage1.Controls.Add($btnRestore)
	$tabPage1.UseVisualStyleBackColor = $True
	$tabPage1.Text = 'Home'
	$tabPage1.DataBindings.DefaultDataSourceUpdateMode = 0
	$tabPage1.TabIndex = 0
	$tabPage1.Name = 'tabPage1'
	$tabPage1.Padding = '3, 3, 3, 3'
	$tabPage1.Location = '4, 22'
	$tabPage1.Size = '805, 608'
	#
	# groupTools
	#
	$groupTools.Controls.Add($btnLBConsole)
	$groupTools.Controls.Add($btnRestart)
	$groupTools.Controls.Add($btnWinRS)
	$groupTools.Controls.Add($btnBAQ)
	$groupTools.Controls.Add($btnFileVersion)
	$groupTools.Controls.Add($btnCDrive)
	$groupTools.Controls.Add($btnRA)
	$groupTools.Controls.Add($btnRDP)
	#$groupTools.Location = '0, 301'
	$groupTools.Location = '0, 278'
	$groupTools.Name = "groupTools"
	$groupTools.Size = '126, 270'
	$groupTools.TabIndex = 8
	$groupTools.TabStop = $False
	$groupTools.Text = "Tools"
	#
	# btnRestore
	#
	$btnRestore.Location = '75,590'
	$btnRestore.Name = "btnRestore"
	$btnRestore.Size = '50, 25'
	$btnRestore.TabIndex = 12
	$btnRestore.Text = "+"
	$btnRestore.UseVisualStyleBackColor = $True
	$btnRestore.add_Click($btnRestore_Click)
	#
	# btnResize
	#
	$btnResize.Location = '19,590'
	$btnResize.Name = "btnResize"
	#$btnResize.Size = '110, 25'
	$btnResize.Size = '50, 25'
	$btnResize.TabIndex = 12
	$btnResize.Text = "-"
	$btnResize.UseVisualStyleBackColor = $True
	$btnResize.add_Click($btnResize_Click)
	#
	#btnLBConsole
	#
	$btnLBConsole.Location = '9, 205'
	$btnLBConsole.Name = "btnLBConsole"
	$btnLBConsole.Size = '110, 25'
	$btnLBConsole.TabIndex = 12
	$btnLBConsole.Text = "LBConsole"
	$btnLBConsole.UseVisualStyleBackColor = $True
	$btnLBConsole.FlatStyle = 0
	$btnLBConsole.add_Click($btnLBConsole_Click)
	#
	# btnRestart
	#
	#$btnRestart.Location = '9, 174'
	#$btnRestart.Location = '9, 205'
	$btnRestart.Location = '9, 236'
	$btnRestart.Name = "btnRestart"
	$btnRestart.Size = '110, 25'
	$btnRestart.TabIndex = 12
	$btnRestart.Text = "Restart Computer"
	$btnRestart.UseVisualStyleBackColor = $True
	$btnRestart.FlatStyle = 0
	$btnRestart.add_Click($btnRestart_Click)
	$btnRestart.Enabled = $False
	if ($currentUser -eq 'densan_sup') {$btnRestart.Enabled = $True}
	#
	# btnWinRS
	#
	$btnWinRS.Location = '9, 143'
	$btnWinRS.Name = "btnWinRS"
	$btnWinRS.Size = '110, 25'
	$btnWinRS.TabIndex = 12
	$btnWinRS.Text = "WinRS"
	$btnWinRS.UseVisualStyleBackColor = $True
	$btnWinRS.FlatStyle = 0
	$btnWinRS.add_Click($btnWinRS_Click)
	#
	# btnDesktopLocator
	#
	$btnBAQ.Location = '9, 174'
	$btnBAQ.Name = "btnBAQ"
	$btnBAQ.Size = '110, 25'
	$btnBAQ.TabIndex = 12
	$btnBAQ.Text = "BAQ"
	$btnBAQ.UseVisualStyleBackColor = $True
	$btnBAQ.FlatStyle = 0
	$btnBAQ.add_Click($btnBAQ_Click)
	#
	# btnFileVersion
	#
	$btnFileVersion.Location = '9, 112'
	$btnFileVersion.Name = "btnMsg"
	$btnFileVersion.Size = '110, 25'
	$btnFileVersion.TabIndex = 11
	$btnFileVersion.Text = "File Version Check"
	$btnFileVersion.UseVisualStyleBackColor = $True
	$btnFileVersion.FlatStyle = 0
	$btnFileVersion.add_Click($btnFileVersion_Click)
	#
	# btnCDrive
	#
	$btnCDrive.Location = '9, 81'
	$btnCDrive.Name = "btnCDrive"
	$btnCDrive.Size = '110, 25'
	$btnCDrive.TabIndex = 10
	$btnCDrive.Text = "View C Drive"
	$btnCDrive.UseVisualStyleBackColor = $True
	$btnCDrive.FlatStyle = 0
	$btnCDrive.add_Click($btnCDrive_Click)
	#
	# btnRA
	#
	$btnRA.Location = '9, 50'
	$btnRA.Name = "btnRA"
	$btnRA.Size = '110, 25'
	$btnRA.TabIndex = 9
	$btnRA.Text = "Remote Assist"
	$btnRA.UseVisualStyleBackColor = $True
	$btnRA.FlatStyle = 0
	$btnRA.add_Click($btnRA_Click)
	#
	# btnRDP
	#
	$btnRDP.Location = '9, 19'
	$btnRDP.Name = "btnRDP"
	$btnRDP.Size = '110, 25'
	$btnRDP.TabIndex = 8
	$btnRDP.Text = "Remote Desktop"
	$btnRDP.UseVisualStyleBackColor = $True
	$btnRDP.FlatStyle = 0
	$btnRDP.add_Click($btnRDP_Click)
	#
	# groupInfo
	#
	$groupInfo.Controls.Add($btnServices)
	$groupInfo.Controls.Add($btnGetError)
	$groupInfo.Controls.Add($btnGetLargeFiles)
	$groupInfo.Controls.Add($btnCatalyst)
	$groupInfo.Controls.Add($btnTopProcess)
	$groupInfo.Controls.Add($btnSystemInfo)
	#$groupInfo.Location = '0, 83'
	$groupInfo.Location = '0, 66'
	$groupInfo.Name = "groupInfo"
	$groupInfo.Size = '126, 208'
	$groupInfo.TabIndex = 7
	$groupInfo.TabStop = $False
	$groupInfo.Text = "Information"
	#
	# btnServices
	#
	$btnServices.Location = '9, 174'
	$btnServices.Name = "btnServices"
	$btnServices.Size = '110, 25'
	$btnServices.TabIndex = 7
	$btnServices.Text = "Services"
	$btnServices.UseVisualStyleBackColor = $True
	$btnServices.FlatStyle = 0
	$btnServices.add_Click($btnServices_Click)
	#
	# btnGetError
	#
	$btnGetError.Location = '9, 143'
	$btnGetError.Name = "btnGetError"
	$btnGetError.Size = '110, 25'
	$btnGetError.TabIndex = 6
	$btnGetError.Text = "Errors / Warnings"
	$btnGetError.UseVisualStyleBackColor = $True
	$btnGetError.FlatStyle = 0
	$btnGetError.add_Click($btnGetError_Click)
	#
	# btnGetLargeFiles
	#
	$btnGetLargeFiles.Location = '9, 112'
	$btnGetLargeFiles.Name = "$btnGetLargeFiles"
	$btnGetLargeFiles.Size = '110, 25'
	$btnGetLargeFiles.TabIndex = 5
	$btnGetLargeFiles.Text = "Large Files"
	$btnGetLargeFiles.UseVisualStyleBackColor = $True
	$btnGetLargeFiles.FlatStyle = 0
	$btnGetLargeFiles.add_Click($btnGetLargeFiles_Click)
	#
	# btnApplications
	#
	#$btnCatalyst.Location = '9, 81'
	$btnCatalyst.Location = '9, 50'
	$btnCatalyst.Name = "btnCatalyst"
	$btnCatalyst.Size = '110, 25'
	$btnCatalyst.TabIndex = 4
	$btnCatalyst.Text = "Catalyst"
	$btnCatalyst.UseVisualStyleBackColor = $True
	$btnCatalyst.FlatStyle = 0
	$btnCatalyst.add_Click($btnCatalyst_Click)
	#
	# btnLocalAdmins
	#
	#$btnTopProcess.Location = '9, 50'
	$btnTopProcess.Location = '9, 81'
	$btnTopProcess.Name = "btnTopProcess"
	$btnTopProcess.Size = '110, 25'
	$btnTopProcess.TabIndex = 3
	$btnTopProcess.Text = "Top Processes"
	$btnTopProcess.UseVisualStyleBackColor = $True
	$btnTopProcess.FlatStyle = 0
	$btnTopProcess.add_Click($btnTopProcess_Click)
	#
	# btnSystemInfo
	#
	$btnSystemInfo.Location = '9, 19'
	$btnSystemInfo.Name = "btnSystemInfo"
	$btnSystemInfo.Size = '110, 25'
	$btnSystemInfo.TabIndex = 2
	$btnSystemInfo.Text = "System Info"
	$btnSystemInfo.UseVisualStyleBackColor = $True
	$btnSystemInfo.FlatStyle = 0
	$btnSystemInfo.add_Click($btnSystemInfo_Click)
	#
	# panel1
	#
	$panel1.Anchor = 'Top, Bottom, Left, Right'
	$panel1.DataBindings.DefaultDataSourceUpdateMode = 0
	$panel1.BorderStyle = "FixedSingle"
	$panel1.TabIndex = 1
	$panel1.Name = 'panel1'
	#$panel1.Size = '642, 582'
	$panel1.Size = '671, 600'
	$panel1.Location = '130, 5'
	#
	# richTextBox
	#
	$richTextBox.Anchor = 'Top, Bottom, Left, Right'
	#$richTextBox.Dock = 'Fill'
	$richTextBox.ContextMenuStrip = $contextMenu
	$richTextBox.Name = "richTextBox"
	$richTextBox.Location = '8,8'
	$richTextBox.Font = "Consolas,10"
	#$richTextBox.Size = '621, 552'
	$richTextBox.Size = '656, 584'
	$richTextBox.TabIndex = 13
	$richTextBox.ReadOnly = $True
	$richTextBox.BackColor = "Window" #http://www.imagemagick.org/script/color.php
	$richTextBox.BorderStyle = "None"
	#$richTextBox.AppendText($textOut)  
	#$richTextBox.WordWrap = $False
	#$richTextBox.ScrollBars = $True
	#
	# btnPing
	#
	#$btnPing.Location = '10,47'
	$btnPing.Location = '10,32'
	$btnPing.Name = "btnPing"
	$btnPing.Size = '110, 25'
	$btnPing.TabIndex = 1
	$btnPing.Text = "Ping Check"
	$btnPing.UseVisualStyleBackColor = $True
	$btnPing.FlatStyle = 0
	$btnPing.add_Click($btnPing_Click)
	#
	# txtComputer
	#
	#$txtComputer.Location = '10, 28'
	#$txtComputer.Location = '0, 14'
	$txtComputer.Location = '0, 5'
	$txtComputer.Name = "txtComputer"
	$txtComputer.Size = '126, 20'
	$txtComputer.TabIndex = 0
	$txtComputer.AutoCompleteSource = 'CustomSource'
    $txtComputer.AutoCompleteMode='SuggestAppend'
    $txtComputer.AutoCompleteCustomSource=$autocomplete
	import-csv -Path "\\Msad\root\JP\TK\lib\it\eus\TOOLS\CAT\catalyst.csv" `
	| % {$txtComputer.AutoCompleteCustomSource.AddRange($_.Machine_Owner)}
	#
	# menu
	#
	[void]$menu.Items.Add($menuFile)
	[void]$menu.Items.Add($menuView)
	[void]$menu.Items.Add($menuHelp)
	$menu.Location = '0, 0'
	$menu.Name = "menu"
	$menu.Size = '780, 24'
	$menu.TabIndex = 0
	$menu.Text = "menuMain"
	#
	# menuFile
	#
	#[void]$menuFile.DropDownItems.Add($menuFileConnect)
	[void]$menuFile.DropDownItems.Add($menuFileExit)
	$menuFile.Name = "menuFile"
	$menuFile.Size = '37, 20'
	$menuFile.Text = "File"
	#
	# menuFileConnect
	#
	$menuFileConnect.Name = "menuFileConnect"
	$menuFileConnect.Size = '186, 22'
	$menuFileConnect.Text = "Connect to domain..."
	$menuFileConnect.add_Click($menuFileConnect_Click)
	#
	# menuFileExit
	#
	$menuFileExit.Name = "menuFileExit"
	$menuFileExit.Size = '186, 22'
	$menuFileExit.Text = "Exit"
	$menuFileExit.add_Click($menuFileExit_Click)
	#
	# menuView
	#
	[void]$menuView.DropDownItems.Add($menuViewEventVwr)
	[void]$menuView.DropDownItems.Add($menuViewServices)
	[void]$menuView.DropDownItems.Add($menuViewUser)
	#[void]$menuView.DropDownItems.Add($menuViewWSUS)
	$menuView.Name = "menuView"
	$menuView.Size = '44, 20'
	$menuView.Text = "View"
	#
	# menuViewEventVwr
	#
	$menuViewEventVwr.Name = "menuViewEventVwr"
	$menuViewEventVwr.Size = '166, 22'
	$menuViewEventVwr.Text = "Event Viewer"
	$menuViewEventVwr.add_Click($menuViewEventVwr_Click)
	#
	# menuViewServices
	#
	$menuViewServices.Name = "menuViewServices"
	$menuViewServices.Size = '166, 22'
	$menuViewServices.Text = "Services"
	$menuViewServices.add_Click($menuViewServices_Click)
	#
	# menuViewUser
	#
	$menuViewUser.Name = "menuViewUser"
	$menuViewUser.Size = '166, 22'
	$menuViewUser.Text = "Users and Groups"
	$menuViewUser.add_Click($menuViewUser_Click)
	#
	# menuViewWSUS
	#
	[void]$menuViewWSUS.DropDownItems.Add($menuViewWSUSReport)
	[void]$menuViewWSUS.DropDownItems.Add($menuViewWSUSUpdate)
	$menuViewWSUS.Name = "menuViewWSUS"
	$menuViewWSUS.Size = '166, 22'
	$menuViewWSUS.Text = "WSUS Logs"
	#
	# menuViewWSUSReport
	#
	$menuViewWSUSReport.Name = "menuViewWSUSReport"
	$menuViewWSUSReport.Size = '126, 22'
	$menuViewWSUSReport.Text = "Reporting"
	$menuViewWSUSReport.add_Click($menuViewWSUSReport_Click)
	#
	# menuViewWSUSUpdate
	#
	$menuViewWSUSUpdate.Name = "menuViewWSUSUpdate"
	$menuViewWSUSUpdate.Size = '152, 22'
	$menuViewWSUSUpdate.Text = "Updates"
	$menuViewWSUSUpdate.add_Click($menuViewWSUSUpdate_Click)
	#
	# contextMenu
	#
	[void]$contextMenu.Items.Add($cmsProcEnd)
	[void]$contextMenu.Items.Add($cmsStartupRemove)
	[void]$contextMenu.Items.Add($cmsAdminAdd)
	[void]$contextMenu.Items.Add($cmsAdminRemove)
	[void]$contextMenu.Items.Add($cmsAppUninstall)
	[void]$contextMenu.Items.Add($cmsSelect)
	$contextMenu.Name = "contextMenu"
	$contextMenu.Size = '188, 114'
	#
	# cmsProcEnd
	#
	$cmsProcEnd.Name = "cmsProcEnd"
	$cmsProcEnd.Size = '187, 22'
	$cmsProcEnd.Text = "End Process"
	$cmsProcEnd.Visible = $False
	$cmsProcEnd.add_Click($cmsProcEnd_Click)
	#
	# cmsStartupRemove
	#
	$cmsStartupRemove.Name = "cmsStartupRemove"
	$cmsStartupRemove.Size = '187, 22'
	$cmsStartupRemove.Text = "Remove Startup Item"
	$cmsStartupRemove.Visible = $False
	$cmsStartupRemove.add_Click($cmsStartupRemove_Click)
	#
	# cmsAdminAdd
	#
	$cmsAdminAdd.Name = "cmsAdminAdd"
	$cmsAdminAdd.Size = '187, 22'
	$cmsAdminAdd.Text = "Add Local Admin"
	$cmsAdminAdd.Visible = $False
	$cmsAdminAdd.add_Click($cmsAdminAdd_Click)
	#
	# cmsAdminRemove
	#
	$cmsAdminRemove.Name = "cmsAdminRemove"
	$cmsAdminRemove.Size = '187, 22'
	$cmsAdminRemove.Text = "Remove Local Admin"
	$cmsAdminRemove.Visible = $False
	$cmsAdminRemove.add_Click($cmsAdminRemove_Click)
	#
	# cmsAppUninstall
	#
	$cmsAppUninstall.Name = "cmsAppUninstall"
	$cmsAppUninstall.Size = '187, 22'
	$cmsAppUninstall.Text = "Uninstall"
	$cmsAppUninstall.Visible = $False
	$cmsAppUninstall.add_Click($cmsAppUninstall_Click)
	#
	# cmsSelect
	#
	$cmsSelect.Name = "cmsSelect"
	$cmsSelect.Size = '187, 22'
	$cmsSelect.Text = "Select Computer"
	$cmsSelect.Visible = $False
	$cmsSelect.add_Click($cmsSelect_Click)
	#
	# menuHelp
	#
	[void]$menuHelp.DropDownItems.Add($menuHelpAbout)
	$menuHelp.Name = "menuHelp"
	$menuHelp.Size = '44, 20'
	$menuHelp.Text = "Help"
	#
	# menuHelpAbout
	#
	$menuHelpAbout.Name = "menuHelpAbout"
	$menuHelpAbout.Size = '152, 22'
	$menuHelpAbout.Text = "Contact us"
	$menuHelpAbout.add_Click($menuHelpAbout_Click)
	#
    #Status Bar
	#
	$SB.Text = "Ready"
	$SB.BringToFront()
	#$SB.SizingGrip = $False 
	$imageSBPContact = [System.Drawing.Image]::FromFile('\\Msad\root\JP\TK\lib\it\eus\TOOLS\Lazy.Admin.Tool\icon_follow_mail.png')
	$SB.Items.Add('',$imageSBPContact,$SBPContactLink_Click)
	$SB.Items.Add($SBPContact)
	$SB.Items.Add($SBPStatus)
	$SB.Items.Add($progressBar)
	#
	#Status Bar Contact
	#
	$SBPContact.Text = "Contact Us                   "
	$SBPContact.IsLink = $True
	$SBPContact.add_Click($SBPContactLink_Click)
	$SBPContact.LinkBehavior = 'NeverUnderline'
	$SBPContact.BorderSides = 'Right'
	#
	#Status
	#
	$SBPStatus.Text = "Ready"
	$SBPStatus.Spring = $True
	#$SBPStatus.BorderSides = 'Left'
	#
	#ProgressBar
	#
	#$progressBar.Value = 100
	$progressBar.Name = 'progressBar'
	#
	# tabPage2
	#
	$tabPage2.UseVisualStyleBackColor = $True
	$tabPage2.Text = 'Other Tools'
	$tabPage2.DataBindings.DefaultDataSourceUpdateMode = 0
	$tabPage2.TabIndex = 1
	$tabPage2.Name = 'tabPage2'
	$tabPage2.Padding = '3, 3, 3, 3'
	$tabPage2.Location = '4, 22'
	$tabPage2.Size = '192, 74'
	
	#
	# btnAddCheck
	#
	#$btnCheck_Addins.Location = '10,47'
	$btnCheck_Addins.Location = '10,32'
	$btnCheck_Addins.Name = "btnAddCheck"
	$btnCheck_Addins.Size = '110, 25'
	$btnCheck_Addins.TabIndex = 1
	$btnCheck_Addins.Text = "Add-in Checker"
	$btnCheck_Addins.UseVisualStyleBackColor = $True
	$btnCheck_Addins.FlatStyle = 0
	$btnCheck_Addins.add_Click($btnCheck_Addins_Click)
	
	#
	# btnC2CLogs
	#
	$btnC2CLogs.Location = '10,63'
	$btnC2CLogs.Name = "btnC2CLogs"
	$btnC2CLogs.Size = '110, 25'
	$btnC2CLogs.TabIndex = 1
	$btnC2CLogs.Text = "C2C Log Grabber"
	$btnC2CLogs.UseVisualStyleBackColor = $True
	$btnC2CLogs.FlatStyle = 0
	$btnC2CLogs.add_Click($btnC2CLogs_Click)
	#
	# btnAppHist
	#
	$btnAppHist.Location = '10,94'
	$btnAppHist.Name = "btnAppHist"
	$btnAppHist.Size = '110, 25'
	$btnAppHist.TabIndex = 1
	$btnAppHist.Text = "AppInst History"
	$btnAppHist.UseVisualStyleBackColor = $True
	$btnAppHist.FlatStyle = 0
	$btnAppHist.add_Click($btnAppHist_Click)
	#
	# btnAppFinder
	#
	$btnAppFinder.Location = '10,125'
	$btnAppFinder.Name = "$btnAppFinder"
	$btnAppFinder.Size = '110, 25'
	$btnAppFinder.TabIndex = 1
	$btnAppFinder.Text = "App Finder"
	$btnAppFinder.UseVisualStyleBackColor = $True
	$btnAppFinder.FlatStyle = 0
	$btnAppFinder.add_Click($btnAppFinder_Click)
	
	#
	# btnLogOnHist
	#
	$btnLogOnHist.Location = '10,156'
	$btnLogOnHist.Name = "$btnLogOnHist"
	$btnLogOnHist.Size = '110, 25'
	$btnLogOnHist.TabIndex = 1
	$btnLogOnHist.Text = "Log On History"
	$btnLogOnHist.UseVisualStyleBackColor = $True
	$btnLogOnHist.FlatStyle = 0
	$btnLogOnHist.add_Click($btnLogOnHist_Click)
	
	#
	# btnProcessCheck
	#
	$btnProcessCheck.Location = '10,187'
	$btnProcessCheck.Name = "$btnProcessCheck"
	$btnProcessCheck.Size = '110, 25'
	$btnProcessCheck.TabIndex = 1
	$btnProcessCheck.Text = "Process Checker"
	$btnProcessCheck.UseVisualStyleBackColor = $True
	$btnProcessCheck.FlatStyle = 0
	$btnProcessCheck.add_Click($btnProcessCheck_Click)
	
	$btnLicenseQuery.Location = '10,218'
	$btnLicenseQuery.Name = "$btnProcessCheck"
	$btnLicenseQuery.Size = '110, 25'
	$btnLicenseQuery.TabIndex = 1
	$btnLicenseQuery.Text = "License Query"
	$btnLicenseQuery.UseVisualStyleBackColor = $True
	$btnLicenseQuery.FlatStyle = 0
	$btnLicenseQuery.add_Click($btnLicenseQuery_Click)
	
	
	#endregion Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formMain.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formMain.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formMain.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formMain.ShowDialog()

} 

function Monitor_Check{
	Try {
		[array]$Monitors = Get-WmiObject Win32_PnPEntity -ComputerName $ComputerName -ErrorAction Stop | Where {$_.Service -eq "monitor"}
		$i = 0
        If ($Monitors){
			$richTextBox.Text += "`n","($($Monitors.Count)) MONITOR/S INSTALLED","`n"
			$richTextBox.Text += "============================================================","`n"
			$richTextBox.Text += Get-WmiObject -cn $ComputerName WmiMonitorID -Namespace root\wmi | 
			Select @{n="Model                  ";e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}},`
			@{n="Serial #               ";e={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00)}} | 
			ft -AutoSize | 
			Out-String
		} 
		Else{
			$richTextBox.Text += "`n","NO MONITOR/S DETECTED","`n", 
			"===========================================================","`n" 
		}
	} 
	Catch {
		Clear_Screen
		AppendText $richTextBox Red "`n","Access denied, Machine Organization Unit (OU) is not set to Tokyo!"
	}
}
	
function Clear_Screen{
	$richTextBox.clear()
	$progressBar.Value = 0
}

function Get-ComputerName{
	if($txtComputer.Text -eq "." -or $txtComputer.Text -eq "localhost" -or $txtComputer.Text -eq "" -or $txtComputer.Text -eq $null){$txtComputer.Text = hostname}
	$Script:ComputerName = $txtComputer.Text.Trim()
}	

function System_Info{
	Get-Computername
	$AFS = $null
	$matches = $null
	$RealValue = $null
	$EmptyString = "" |Out-String
	$richTextBox.Text += "Getting System Information for $Computername..$EmptyString"
	$ComputerSystem = Get-WMIObject -computername $ComputerName -class Win32_ComputerSystem
	$WhoIsLoggedOn = Invoke-Command -Computername $ComputerSystem.Name {quser}
	$TestMatch = $WhoIsLoggedOn[1] -match "[a-zA-Z0-9_]+"
	If ($TestMatch -eq $False)
		{
		$TestMatch = $WhoIsLoggedOn[4] -match "[a-zA-Z0-9_]+"
		}
	$RealValue = $matches.values
	$ComputerNameUp = ("$ComputerName").ToUpper()
	$richTextBox.Text += "Collecting monitor information..$EmptyString"
	$ComputerOS = Get-WMIObject -computername $ComputerName -class Win32_OperatingSystem
	$ComputerBIOS = Get-WMIObject -computername $ComputerName -class Win32_BIOS
	$ComputerCPU = Get-WmiObject -computername $ComputerName Win32_Processor
	$LastBootTime = (Get-WmiObject -Class Win32_OperatingSystem -computername $ComputerName).LastBootUpTime
	$Sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime) 
	$richTextBox.Text += "Calculating CPU/MEMORY/DISK USAGE..$EmptyString"
	$AdUserInfo = Get-ADuser "$RealValue" -Properties *
    $PWExpire = (get-date) - ($AdUserInfo.PasswordLastSet)
    $ExpireDate = 90 - $PWExpire.Days
    $MyWorkSpace = Get-MSRemoteRegistry -ComputerName $ComputerName -Path "HKEY_LOCAL_MACHINE\SOFTWARE\den jelmer\P8Build\Current" |Select Name, Data |where {$_.Name -like "*Current_Release*"}|select -ExpandProperty data |Out-String
	$MyWorkSpace = $MyWorkSpace.Trim()
	$AFSVersion = Get-MSRemoteRegistry -computername $ComputerName -path "HKEY_LOCAL_MACHINE\SOFTWARE\OpenAFS\Versions" | where {$_.Name -like "Current"} |select -ExpandProperty Data 
	$AFSMode = Get-MSRemoteRegistry -computername $ComputerName -path "HKEY_LOCAL_MACHINE\SOFTWARE\OpenAFS\Client" | where {$_.Name -like "Mode"} | select -ExpandProperty Data 
	$AFS = $AFSVersion + $AFSMode
	$richTextBox.clear()
	
	"*SYSTEM INFORMATION - " + $ComputerNameUp
	"============================================================"
	""
	"Logon User              : " + $matches.Values
	"Last Logon              : " + $AdUserInfo.LastLogonDate
	"PW Last Set             : " + $AdUserInfo.PasswordLastSet
    "Days until PW Expiration: " + $ExpireDate + " Days"
	#"Hostname                : " + $ComputerSystem.Caption
	"Model                   : " + $ComputerSystem.Model
	"OS                      : " + $ComputerOS.Caption
	"Last Reboot             : " + $ComputerOS.ConvertToDateTime($ComputerOS.LastBootUptime)
	"Uptime                  : " + $Sysuptime.days + " days " + $Sysuptime.hours + " hours " + $Sysuptime.minutes + " min"
	#"Serial                  : " + $ComputerBIOS.SerialNumber
	"AFS Version             : " + $AFS
	"Build Version           : " + $MyWorkSpace
	""
	$ComputerSystemR = Get-WmiObject -ComputerName $ComputerName -Class Win32_operatingsystem -Property CSName, TotalVisibleMemorySize, FreePhysicalMemory
 	$HardDiskSpaceCheck = get-WmiObject win32_logicaldisk -Filter "DeviceID='C:'" -computername $ComputerName -Property FreeSpace, Size
	$MachineName = $ComputerSystemR.CSName
 	$FreePhysicalMemory = ($ComputerSystemR.FreePhysicalMemory) / (1mb)
 	$TotalVisibleMemorySize = ($ComputerSystemR.TotalVisibleMemorySize) / (1mb)
 	$TotalVisibleMemorySizeR = "{0:N2}" -f $TotalVisibleMemorySize
 	$TotalFreeMemPerc = ($FreePhysicalMemory/$TotalVisibleMemorySize)*100
 	$TotalFreeMemPercR = "{0:N2}" -f $TotalFreeMemPerc
 	$TotalCPULoad = $ComputerCPU.LoadPercentage
	$FreeSpace = ($HardDiskSpaceCheck.FreeSpace)/(1gb)
	$FreeSpaceR = "{0:N2}" -f $FreeSpace
	$DiskSize = ($HardDiskSpaceCheck.Size)/(1gb)
	$DiskSizeR = "{0:N2}" -f $DiskSize
	""
	"*CPU/MEMORY/DISK USAGE"
	"============================================================"
	""
	"RAM Installed           : $TotalVisibleMemorySizeR GB"
	"Free Physical Memory    : $TotalFreeMemPercR %"
	"CPU Load                : $TotalCPULoad %"
	"Total Disk Size         : $DiskSizeR GB"
	"Free Disk Space         : $FreeSpaceR GB"
	""
	}		

function System_Info2{
	Get-Computername
	$User = $null; $CaptureDefPrint = $null; $matches = $null; $CaptureDefPrint = $null
	$User = (Get-WmiObject Win32_Process -Computername $ComputerName -filter 'name="explorer.exe"').GetOwner() | Select Domain,User |select -ExpandProperty User
	$objectSid = get-aduser $user -property * |select objectSid
	$SID = ($objectSID.ObjectSid).Value
	$DefaultPrinter = Get-MSRemoteRegistry -ComputerName $ComputerName -Path "HKEY_USERS\$SID\Software\Microsoft\Windows NT\CurrentVersion\Windows" |Select Data
	$PrintMatch = $DefaultPrinter[1] -match "[^,]*"
	if ($matches -eq $null){
			$PrintMatch = $DefaultPrinter[1] -match "[^,]*"
	}
	$CaptureDefPrint = ($matches[0].ToString()).SubString(7)
	if ($CaptureDefPrint -eq "1}"){
	$CaptureDefPrint = $DefaultPrinter[0] | select data | select -ExpandProperty Data
	}
	$ComputerSystem = Get-WMIObject -computername $ComputerName -class Win32_ComputerSystem
	$ComputerNameUp = ("$ComputerName").ToUpper()
	$ComputerOS = Get-WMIObject -computername $ComputerName -class Win32_OperatingSystem
	$ComputerBIOS = Get-WMIObject -computername $ComputerName -class Win32_BIOS
	$ComputerCPU = Get-WmiObject -computername $ComputerName Win32_Processor
	$LastBootTime = (Get-WmiObject -Class Win32_OperatingSystem -computername $ComputerName).LastBootUpTime
	$Sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime) 
	$AdUserInfo = Get-ADuser "$user" -Properties *
	
	$MyWorkSpace = Get-MSRemoteRegistry -ComputerName $ComputerName -Path "HKEY_LOCAL_MACHINE\SOFTWARE\den jelmer\P8Build\Current" |Select Name, Data |where {$_.Name -like "*Current_Release*"}|select -ExpandProperty data |Out-String
	$MyWorkSpace = $MyWorkSpace.Trim()
	
	$AFSVersion = Get-MSRemoteRegistry -computername $ComputerName -path "HKEY_LOCAL_MACHINE\SOFTWARE\OpenAFS\Versions" | where {$_.Name -like "Current"} |select -ExpandProperty Data 
	$AFSMode = Get-MSRemoteRegistry -computername $ComputerName -path "HKEY_LOCAL_MACHINE\SOFTWARE\OpenAFS\Client" | where {$_.Name -like "Mode"} | select -ExpandProperty Data 
	$AFS = $AFSVersion + $AFSMode
	
	
	"*SYSTEM INFORMATION - " + $ComputerNameUp
	"============================================================"
	""
	"Logon User              : " + $User
	"Last Logon              : " + $AdUserInfo.LastLogonDate
	"PW Last Set             : " + $AdUserInfo.PasswordLastSet
	"Hostname                : " + $ComputerSystem.Caption
	"Model                   : " + $ComputerSystem.Model
	"OS                      : " + $ComputerOS.Caption
	"Last Reboot             : " + $ComputerOS.ConvertToDateTime($ComputerOS.LastBootUptime)
	"Uptime                  : " + $Sysuptime.days + " days " + $Sysuptime.hours + " hours " + $Sysuptime.minutes + " min"
	#"Serial                  : " + $ComputerBIOS.SerialNumber
	"Default Printer         : " + $CaptureDefPrint
	"AFS Version             : " + $AFS
	"Build Version           : " + $MyWorkSpace
	""
	$ComputerSystemR = Get-WmiObject -ComputerName $ComputerName -Class Win32_operatingsystem -Property CSName, TotalVisibleMemorySize, FreePhysicalMemory
 	$HardDiskSpaceCheck = get-WmiObject win32_logicaldisk -Filter "DeviceID='C:'" -computername $ComputerName -Property FreeSpace, Size
	$MachineName = $ComputerSystemR.CSName
 	$FreePhysicalMemory = ($ComputerSystemR.FreePhysicalMemory) / (1mb)
 	$TotalVisibleMemorySize = ($ComputerSystemR.TotalVisibleMemorySize) / (1mb)
 	$TotalVisibleMemorySizeR = "{0:N2}" -f $TotalVisibleMemorySize
 	$TotalFreeMemPerc = ($FreePhysicalMemory/$TotalVisibleMemorySize)*100
 	$TotalFreeMemPercR = "{0:N2}" -f $TotalFreeMemPerc
 	$TotalCPULoad = $ComputerCPU.LoadPercentage
	$FreeSpace = ($HardDiskSpaceCheck.FreeSpace)/(1gb)
	$FreeSpaceR = "{0:N2}" -f $FreeSpace
	$DiskSize = ($HardDiskSpaceCheck.Size)/(1gb)
	$DiskSizeR = "{0:N2}" -f $DiskSize
	""
	"*CPU/MEMORY/DISK USAGE"
	"============================================================"
	""
	"RAM Installed           : $TotalVisibleMemorySizeR GB"
	"Free Physical Memory    : $TotalFreeMemPercR %"
	"CPU Load                : $TotalCPULoad %"
	"Total Disk Size         : $DiskSizeR GB"
	"Free Disk Space         : $FreeSpaceR GB"
	""
}	

function Get-InstalledSoftware {
	Param
	(
		[Alias('Computer','ComputerName','HostName')]
		[Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
		[string[]]$Name = $env:COMPUTERNAME
	)
	Begin
	{
		$LMkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		$LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
		$CUkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
		$CUtype = [Microsoft.Win32.RegistryHive]::CurrentUser
		
	}
	Process
	{
		ForEach($Computer in $Name)
		{
			$MasterKeys = @()
			If(!(Test-Connection -ComputerName $Computer -count 1 -quiet))
			{
				Write-Error -Message "Unable to contact $Computer. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $Computer
				Break
			}
			$CURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($CUtype,$computer)
			$LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype,$computer)
			ForEach($Key in $LMkeys)
			{
				$RegKey = $LMRegKey.OpenSubkey($key)
				If($RegKey -ne $null)
				{
					ForEach($subName in $RegKey.getsubkeynames())
					{
						foreach($sub in $RegKey.opensubkey($subName))
						{
							$MasterKeys += (New-Object PSObject -Property @{
							"ComputerName" = $Computer
							"Name" = $sub.getvalue("displayname")
							"SystemComponent" = $sub.getvalue("systemcomponent")
							"ParentKeyName" = $sub.getvalue("parentkeyname")
							"Version" = $sub.getvalue("DisplayVersion")
							"UninstallCommand" = $sub.getvalue("UninstallString")
							})
						}
					}
				}
			}
			ForEach($Key in $CUKeys)
			{
				$RegKey = $CURegKey.OpenSubkey($Key)
				If($RegKey -ne $null)
				{
					ForEach($subName in $RegKey.getsubkeynames())
					{
						foreach($sub in $RegKey.opensubkey($subName))
						{
							$MasterKeys += (New-Object PSObject -Property @{
							"ComputerName" = $Computer
							"Name" = $sub.getvalue("displayname")
							"SystemComponent" = $sub.getvalue("systemcomponent")
							"ParentKeyName" = $sub.getvalue("parentkeyname")
							"Version" = $sub.getvalue("DisplayVersion")
							"UninstallCommand" = $sub.getvalue("UninstallString")
							})
						}
					}
				}
			}
			$MasterKeys = ($MasterKeys | Where {$_.Name -ne $Null -AND $_.SystemComponent -ne "1" -AND $_.ParentKeyName -eq $Null} | select Name,Version | sort Name)
			$MasterKeys
		}
	}
	End
	{
		
	}
	
	
}

function Get-Logonhistory {
Param (
 [string]$Computer = (Read-Host Remote computer name),
 [int]$Days = 10
 )
 cls
 $Result = @()
 $richTextBox.text += "Gathering Event Logs, this can take a while...","`n"
 $ELogs = Get-EventLog System -Source Microsoft-Windows-WinLogon -After (Get-Date).AddDays(-$Days) -ComputerName $Computer
 If ($ELogs)
 { $richTextBox.text += "Processing...","`n"
 ForEach ($Log in $ELogs)
 { If ($Log.InstanceId -eq 7001)
   { $ET = "Logon"
   }
   ElseIf ($Log.InstanceId -eq 7002)
   { $ET = "Logoff"
   }
   Else
   { Continue
   }
   $Result += New-Object PSObject -Property @{
    Time = $Log.TimeWritten
    'Event Type' = $ET
    User = (New-Object System.Security.Principal.SecurityIdentifier $Log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])
   }
 }
 $Result | Select Time,"Event Type",User | Sort Time -Descending | Out-GridView
 $richTextBox.text += "Done.","`n"
 }
 Else
 { $richTextBox.text += "Problem with $ComputerName.","`n"
 $richTextBox.text += "If you see a 'Network Path not found' error, try starting the Remote Registry service on that computer.","`n"
 $richTextBox.text += "Or there are no logon/logoff events (XP requires auditing be turned on)","`n"
 }
}

function AppendText($richTextBox, $color, $text){
	$start = $richTextBox.TextLength
	$richTextBox.AppendText($text)
	$end = $richTextBox.TextLength
	$richTextBox.Select($start, $end – $start)
	$richTextBox.SelectionColor = $color
	$richTextBox.SelectionLength = 0 #clear
}

#End Function

#Load the form
	
	Call_LazyAdminTool | Out-Null
		
	#Perform cleanup
	OnApplicationExit
