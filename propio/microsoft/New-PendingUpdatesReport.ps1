#########################################################
#                                                       #
# Monitoring Windows Updates and Pending Restarts       #
#                                                       #
#########################################################

#########################################################
# List of computers to be monitored
#########################################################
$Servers = Get-Content .\Machines.txt
$CN_NUM = 1
#########################################################

$results = @()
foreach ($Computer in $Servers) 
{ 
    Write-Progress -Activity “Cargando reporte de actualizaciones” -status “Procesando equipo $Computer” -percentComplete ($CN_NUM / $Servers.count*100)
	try 
  	{ 
	  	$service = Get-WmiObject Win32_Service -Filter 'Name="wuauserv"' -ComputerName $Computer -Ea 0
		$WUStartMode = $service.StartMode
		$WUState = $service.State
		$WUStatus = $service.Status
  	
		try{
			if (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
			{ 
				#check if the server is the same where this script is running
				if($Computer -eq "$env:computername.$env:userdnsdomain")
				{
					$UpdateSession = New-Object -ComObject Microsoft.Update.Session
				}
				else { $UpdateSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computer)) }
				$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
				$SearchResult = $UpdateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0")
				$Critical = $SearchResult.updates | where { $_.MsrcSeverity -eq "Critical" }
				$important = $SearchResult.updates | where { $_.MsrcSeverity -eq "Important" }
				$other = $SearchResult.updates | where { $_.MsrcSeverity -eq $null }
				# Get windows updates counters
				$totalUpdates = $($SearchResult.updates.count)
				$totalCriticalUp = $($Critical.count)
				$totalImportantUp = $($Important.count)
				
				if($totalUpdates -gt 0)
				{
					$updatesToInstall = "SÍ"
				}
				else { $updatesToInstall = "NO" }
			}
			else
			{
				# if cannot connected to the server the updates are listed as not defined
				$totalUpdates = "N/A"
				$totalCriticalUp = "N/A"
				$totalImportantUp = "N/A"
			}
		}
		catch 
        { 
			# if an error occurs the updates are listed as not defined
        	Write-Warning "$Computer`: $_" 
         	$totalUpdates = "N/A"
			$totalCriticalUp = "N/A"
			$totalImportantUp = "N/A"
			$updatesToInstall = "NO"
        }
  
        # Querying WMI for build version 
        $WMI_OS = Get-WmiObject -Class Win32_OperatingSystem -Property BuildNumber, CSName -ComputerName $Computer -Authentication PacketPrivacy -Impersonation Impersonate

        # Making registry connection to the local/remote computer 
        $RegCon = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"LocalMachine",$Computer) 
         
        # If Vista/2008 & Above query the CBS Reg Key 
        If ($WMI_OS.BuildNumber -ge 17763)
        { 
            $RegSubKeysCBS = $RegCon.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\").GetSubKeyNames() 
            $CBSRebootPend = $RegSubKeysCBS -contains "RebootPending" 
        }
		else{
			$CBSRebootPend = $false
		}
           
        # Query WUAU from the registry 
        $RegWUAU = $RegCon.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\") 
        $RegSubKeysWUAU = $RegWUAU.GetSubKeyNames() 
        $WUAURebootReq = $RegSubKeysWUAU -contains "RebootRequired" 
		
		If($CBSRebootPend –OR $WUAURebootReq)
		{
			$machineNeedsRestart = "SÍ"
		}
		else
		{
			$machineNeedsRestart = "NO"
		}
         
        # Closing registry connection 
        $RegCon.Close() 
		
		if($machineNeedsRestart -or $updatesToInstall -or ($WUStartMode -eq "Manual") -or ($totalUpdates -eq "N/A"))
		{
			$GR_TMP = New-Object PSObject 
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Hostname" -Value $WMI_OS.CSName 
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Estado actual" -Value $WUStatus
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Servicio - Estado" -Value $WUState
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Servicio - Inicio" -Value $WUStartMode 
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "# Updates" -Value $totalUpdates
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "# Updates por instalar" -Value $updatesToInstall
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "# Updates críticas" -Value $totalCriticalUp
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "# Updates importantes" -Value $totalImportantUp
            Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Reinicio pendiente" -Value $machineNeedsRestart
            $results += $GR_TMP       	
		}
  	}
	Catch 
 	{ 
    	Write-Warning "$Computer`: $_" 
  	}
  $CN_NUM++
}

#########################################################
# Formating result
#########################################################

$tableFragment = $results | ConvertTo-HTML -fragment

$tableFragment = $tableFragment -Replace "<table>",'<table class="CanviaTable">'
$HTMLmessage = $null

# HTML Format for Output 
$HTMLmessage = @"

<style type=""text/css"">
body {
 font-family: "Segoe UI",sans-serif;
}
table.CanviaTable {
  background-color: #FFFFFF;
  width: 100%;
  text-align: center;
  border-collapse: collapse;
}
table.CanviaTable td {
  font-size: 11px;
  color: #000000;
  border: 1px solid #AAAAAA;
}
table.CanviaTable th {
  font-size: 12px;
  border: 1px solid #AAAAAA;
  font-weight: bold;
  color: #FFFFFF;
  text-align: center;
  border-left: 2px solid #D3E8F9;
  background: #FF7400;
  background: -moz-linear-gradient(top, #ff9740 0%, #ff8219 66%, #FF7400 100%);
  background: -webkit-linear-gradient(top, #ff9740 0%, #ff8219 66%, #FF7400 100%);
  background: linear-gradient(to bottom, #ff9740 0%, #ff8219 66%, #FF7400 100%);
  border-bottom: 2px solid #444444
}
table.CanviaTable th:first-child {
  border-left: none;
}
h1.CanviaTitulo {
  background-color: #FFFFFF;
  text-align: center;
  font-size: 22px
}
p.CanviaIntro {
  background-color: #FFFFFF;
  text-align: center;
  font-size: 16px
}
</style>

<h1 class="CanviaTitulo">Reporte de actualizaciones y reinicios pendientes</h1>
<p class="CanviaIntro">Este informe se generó porque los servidores que se enumeran a continuación tienen las actualizaciones de Windows listas para ser instaladas, las actualizaciones de Windows configuradas para ser comprobadas manualmente o los servidores que requieren un reinicio. Los servidores que no se encuentren bajo estas condiciones no serán listados.</p>
<body>
$tableFragment
</body>
"@
$HTMLmessage | Out-File .\Report.html