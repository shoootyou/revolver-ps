Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
$DB_LST = Get-MailboxDatabase | ? { $_.Name -like "CNVEXMDBC*" -and $_.ReplicationType -eq "None"}
$DB_SRC = "CNVC1VWMCA00"
$DB_DST = "CNVC4VWMCA00"
$CN_GPR = 1
foreach($DB in $DB_LST){
    $DB_NAM = $DB.Name
    Write-Progress -Activity “Actualizando Bases de Datos de Exchange" -status “Revisando la DB $DB_NAM” -percentComplete ($CN_GPR / $DB_LST.count*100)
    Set-MailboxDatabase $DB_NAM -CircularLoggingEnabled $false -Confirm:$false
    Dismount-Database $DB_NAM -Confirm:$false
    $VAR_00 = (Get-MailboxDatabase $DB_NAM -Status  | % { eseutil /mh $_.edbfilepath }  | Select-String -Pattern "State:").ToString().Substring(19)
    if($VAR_00 -eq "Clean Shutdown"){
        $VAR_01 = Get-MailboxDatabase $DB_NAM -Status | Select LogFolderPath
        $VAR_02 = New-Item ("D:\Temporal\" + $DB_NAM ) -ItemType Directory -Force
        Get-ChildItem $VAR_01.LogFolderPath  | % { Move-Item -Path $_.FullName -Destination ($VAR_02.FullName + "\" + $_.Name) -Force}
        Start-Sleep 10
        Mount-Database $DB_NAM -Confirm:$false
        Add-MailboxDatabaseCopy -Identity $DB_NAM -MailboxServer $DB_DST
        Start-Sleep 10
        Set-MailboxDatabase $DB_NAM -CircularLoggingEnabled $true -Confirm:$false
        Dismount-Database $DB_NAM -Confirm:$false
        Start-Sleep 10
        Mount-Database $DB_NAM -Confirm:$false
        Write-Host "Finalizando proceso"
        Start-Sleep 10
    }
    else{
        Write-Host "No se pudo confirmar la salud de la DB $DB_NAM" -ForegroundColor Yellow
    }
    $CN_GPR++
}