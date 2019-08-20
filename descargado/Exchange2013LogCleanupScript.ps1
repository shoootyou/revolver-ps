Set-Executionpolicy RemoteSigned
Write-Host '------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host 'Especifica la cantidad de días para atrás contando desde hoy, que quieres guardar los logs' -ForegroundColor Green
Write-Host '------------------------------------------------------------------------------------------' -ForegroundColor Green
[int]$PersonalDays = Read-Host
if(!$PersonalDays){
    $days=30 #You can change the number of days here 
}
else{
    $days = $PersonalDays
}
 
# Modify the drive and paths as needed
$ExchangeInstallRoot = "C"
$IISLogPath="inetpub\logs\LogFiles\"
$ExchangeLoggingPath="Program Files\Microsoft\Exchange Server\V15\Logging\"

Write-Host "Estamos removiendo los Logs de Exchange y IIS; se mantendrán los de los últimos" $days "días"
 
Function CleanLogfiles($TargetFolder)
{
    $TargetServerFolder = "\\$E15Server\$ExchangeInstallRoot$\$TargetFolder"
    Write-Host $TargetServerFolder
    if (Test-Path $TargetServerFolder) {
        $Now = Get-Date
        $LastWrite = $Now.AddDays(-$days)
        $Files = Get-ChildItem $TargetServerFolder -Include *.* -Recurse | Where {$_.LastWriteTime -le "$LastWrite"} 
        foreach ($File in $Files)
            {
               # Write-Host "Deleting file $File" -ForegroundColor "Red" 
                Remove-Item $File -ErrorAction SilentlyContinue | out-null}
        }
Else {
    Write-Host "la carpeta $TargetServerFolder no existe! ¡Verifica la ruta!" -ForegroundColor "red"
    }
}
 
$Ex2013 = Get-ExchangeServer | Where {$_.IsE15OrLater -eq $true}
foreach ($E15Server In $Ex2013) {
    CleanLogfiles($IISLogPath)
    CleanLogfiles($ExchangeLoggingPath)
    }


