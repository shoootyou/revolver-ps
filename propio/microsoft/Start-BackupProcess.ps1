function Get-RunTime { [string](Get-Date).Year + (Get-Date).Month + (Get-Date).Day + "," + (Get-Date).Hour + (Get-Date).Minute + (Get-Date).Second }
function Get-FileTime {   If((Get-Date).Month -lt 10){ return [string](Get-Date).Year + "0" + (Get-Date).Month} else{return [string](Get-Date).Year + (Get-Date).Month}}
$CP_WMI = Get-WmiObject -Class Win32_OperatingSystem
$CP_FIL = "E:\QlikSense\00_Backup\DataBase_Bkp\QSR_backup_" + (Get-FileTime) + ".tar"
$LO_PAT = "E:\QlikSense\00_Backup\DataBase_Bkp_Log\LogEjecucionBackup.txt"

New-Item $LO_PAT -ItemType File -ErrorAction SilentlyContinue | Out-Null


[string] (Get-RunTime) + ",----------------------------------------------------------------------------------" | Out-File $LO_PAT -Append utf8
[string] (Get-RunTime) + ",Deteniendo servicio QlikSenseEngineService" | Out-File $LO_PAT -Append utf8
net stop "QlikSenseEngineService" 
[string] (Get-RunTime) + ",Servicio QlikSenseEngineService detenido" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Deteniendo servicio QlikSenseProxyService" | Out-File $LO_PAT -Append utf8
net stop "QlikSenseProxyService"
[string] (Get-RunTime) + ",Servicio QlikSenseProxyService detenido" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Deteniendo servicio QlikSensePrintingService" | Out-File $LO_PAT -Append utf8
net stop "QlikSensePrintingService"
[string] (Get-RunTime) + ",Servicio QlikSensePrintingService detenido" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Deteniendo servicio QlikSenseSchedulerService" | Out-File $LO_PAT -Append utf8
net stop "QlikSenseSchedulerService"
[string] (Get-RunTime) + ",Servicio QlikSenseSchedulerService detenido" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Deteniendo servicio QlikSenseServiceDispatcher" | Out-File $LO_PAT -Append utf8
net stop "QlikSenseServiceDispatcher"
[string] (Get-RunTime) + ",Servicio QlikSenseServiceDispatcher detenido" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Deteniendo servicio QlikSenseRepositoryService" | Out-File $LO_PAT -Append utf8
net stop "QlikSenseRepositoryService"
[string] (Get-RunTime) + ",Servicio QlikSenseRepositoryService detenido" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando proceso de Backup" | Out-File $LO_PAT -Append utf8
& "C:\Program Files\Qlik\Sense\Repository\PostgreSQL\9.6\bin\pg_dump.exe" -h localhost -p 4432 -U postgres -b -F t -f $CP_FIL QSR
[string] (Get-RunTime) + ",Proceso de backup finalizado" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando servicio QlikSenseRepositoryService" | Out-File $LO_PAT -Append utf8
do{
    net start "QlikSenseRepositoryService"
}
until((Get-Service -Name "QlikSenseRepositoryService").Status -eq "Running")
[string] (Get-RunTime) + ",Servicio iniciado QlikSenseRepositoryService" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando servicio QlikSenseEngineService" | Out-File $LO_PAT -Append utf8
do{
    net start "QlikSenseEngineService"
}
until((Get-Service -Name "QlikSenseEngineService").Status -eq "Running")
[string] (Get-RunTime) + ",Servicio iniciado QlikSenseEngineService" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando servicio QlikSenseProxyService" | Out-File $LO_PAT -Append utf8
do{
    net start "QlikSenseProxyService"
}
until((Get-Service -Name "QlikSenseEngineService").Status -eq "Running")
[string] (Get-RunTime) + ",Servicio iniciado QlikSenseProxyService" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando servicio QlikSensePrintingService" | Out-File $LO_PAT -Append utf8
do{
    net start "QlikSensePrintingService"
}
until((Get-Service -Name "QlikSenseEngineService").Status -eq "Running")
[string] (Get-RunTime) + ",Servicio iniciado QlikSensePrintingService" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando servicio QlikSenseSchedulerService" | Out-File $LO_PAT -Append utf8
do{
    net start "QlikSenseSchedulerService"
}
until((Get-Service -Name "QlikSenseSchedulerService").Status -eq "Running")
[string] (Get-RunTime) + ",Servicio iniciado QlikSenseSchedulerService" | Out-File $LO_PAT -Append utf8

[string] (Get-RunTime) + ",Iniciando servicio QlikSenseServiceDispatcher" | Out-File $LO_PAT -Append utf8
do{
    net start "QlikSenseServiceDispatcher"
}
until((Get-Service -Name "QlikSenseServiceDispatcher").Status -eq "Running")
[string] (Get-RunTime) + ",Servicio iniciado QlikSenseServiceDispatcher" | Out-File $LO_PAT -Append utf8