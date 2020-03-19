$DB_DIR = Get-ChildItem \\10.73.1.123\e$\obcq16122013\CUENTAS_PAGAR -Directory | Sort Name
$I = 1
foreach($FLD in $DB_DIR){
    Write-Progress -Id 1 -Activity “Copiando Información” -status (“Trabajando en " + $FLD.Name) -percentComplete ($i / $DB_DIR.count*100)
    if($FLD.Name -ne "V1" -and $FLD.Name -ne "V2" -and $FLD.Name -ne "V3" -and$FLD.Name -ne "V4" -and $FLD.Name -ne "V5"-and $FLD.Name -ne "V10"){
        $DIR_SRC = "\\10.73.1.123\e$\obcq16122013\CUENTAS_PAGAR\" + $FLD.Name
        $DIR_DES = "F:\obcq16122013\CUENTAS_PAGAR\" + $FLD.Name
        $LOG_TXT = "C:\cuentas_pagar_" + $FLD.Name + ".log"
        ROBOCOPY $DIR_SRC $DIR_DES /MIR /COPYALL /Z /R:2 /W:2 /eta /tee /MT:64 /LOG:$LOG_TXT
        #pause
    }
    $I++
}

