Write-Host
Write-Host "Escribir la sección de amarillo"
Write-Host "ej. : " -NoNewline
Write-Host "salasistemas" -ForegroundColor Yellow -NoNewline
Write-Host "@aris.com.pe"
Write-Host
Write-Host "¿Cual es la sala que se requiere habilitar el permiso?" 
$Usuario = Read-Host
$CAL_ES = '@aris.com.pe:\Calendario'
$CAL_US = '@aris.com.pe:\Calendar'

$USR_CAL_ES = $Usuario + $CAL_ES
$USR_CAL_US = $Usuario + $CAL_US

$CUR_ES = Get-MailboxFolderPermission $USR_CAL_ES -ErrorAction SilentlyContinue
$CUR_US = Get-MailboxFolderPermission $USR_CAL_US -ErrorAction SilentlyContinue
if((!$CUR_ES) -and (!$CUR_US)){
    Write-Host
    Write-Warning "Sala no encontrada, por favor verifique el usuario de la sala según el ejemplo"
}
else{
    if(!$CUR_ES){
        Set-MailboxFolderPermission -Identity $USR_CAL_US -User Default -AccessRights LimitedDetails
        Write-Host
        Write-Host "Permiso de la sala modificado correctamente" -ForegroundColor Green
    }
    elseif(!$CUR_US){
        Set-MailboxFolderPermission -Identity $USR_CAL_ES -User Default -AccessRights LimitedDetails -WarningAction SilentlyContinue
        Write-Host
        Write-Host "Permiso de la sala modificado correctamente" -ForegroundColor Green
    }
}