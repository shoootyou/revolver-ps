$DB_GPR = Import-CSV .\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\AUSA-ListadoGrupos.csv
$TS_OUT = '.\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\AdicionUsuariosGrupos.log'
#Start-Transcript -path $TS_OUT -append
$DB_ALL = @()
foreach($IT_GPR in $DB_GPR){
    if($IT_GPR.MiembrosdelGrupo){
        #Write-Host '-----------------------------------------------------------------------------------------------'
        $IT_GPR_NAM = $IT_GPR.'Nombre del Grupo'
        $DB_MBM = $IT_GPR.MiembrosdelGrupo -split ','
        foreach($IT_MBM_NAM in $DB_MBM){
            if(($IT_MBM_NAM | Select-String -Pattern '[@]' -AllMatches).Matches.Count -gt 1){
                $IT_GPR.'Nombre del Grupo'
                
                $IT_MBM_NAM
            }

        }
    }
}
#Stop-Transcript