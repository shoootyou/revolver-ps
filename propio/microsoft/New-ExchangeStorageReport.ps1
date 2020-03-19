#########################################################
# List of computers to be monitored
#########################################################
$Servers = Get-ExchangeServer | Select Name
$CN_NUM = 1
#########################################################

$GR_OUT = "<html>
<body>
<font size=""1"" face=""Segoe UI,Arial,sans-serif"">
<h2 align=""center"">Exchange Environment Report</h3>
<h4 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>
<table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
<tr bgcolor=""#009900"">
<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Segoe UI,Arial,sans-serif"">
<tr align=""center"" bgcolor=""#FFD700"">
<th>Server</th>
<th>Nombre</th>
<th>Unidad</th>
<th>Capacidad total</th>
<th>Capacidad usada</th>
<th>Espacio libre</th>
<th>Espacio libre %</th>
</tr>
"
foreach ( $FR_CPN in $Servers) 
{ 
    $Computer = $FR_CPN.Name
    Write-Progress -Activity “Cargando reporte de actualizaciones” -status “Procesando equipo $Computer” -percentComplete ($CN_NUM / $Servers.count*100)
	$CP_LCL = Get-WmiObject win32_logicaldisk -Filter "Drivetype=3" -ComputerName $Computer -Ea 0
    foreach($CP_DSK in $CP_LCL){
        $GR_OUT+="<tr"
		if ($AlternateRow)
		{
			$Output+=" style=""background-color:#dddddd"""
			$AlternateRow=0
		} else
		{
			$AlternateRow=1
		}
        $Output+="><td>$Computer</td>
        
        
        "
        $GR_TMP = New-Object PSObject 
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Hostname" -Value $Computer
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Nombre" -Value $CP_DSK.VolumeName
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Unidad" -Value $CP_DSK.DeviceID
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Capacidad total" -Value ("{0:N1}" -f( $CP_DSK.Size / 1gb))
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Capacidad usada" -Value ("{0:N1}" -f( ($CP_DSK.Size - $CP_DSK.Freespace) / 1gb))
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Espacio libre" -Value ("{0:N1}" -f( $CP_DSK.Freespace / 1gb ))
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Espacio libre %" -Value ("{0:P0}" -f ($CP_DSK.freespace/$CP_DSK.size))
        $GR_OUT += $GR_TMP
  	}

  $CN_NUM++
}

#########################################################
# Formating result
#########################################################

$tableFragment = $GR_OUT | ConvertTo-HTML -fragment

 $GR_OUT | Out-File .\Report.html

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

<h1 class="CanviaTitulo">Reporte de almacenamiento de Exchange</h1>
<body>
$tableFragment
</body>
"@
$HTMLmessage | Out-File .\New-ExchangeStorageReport.html