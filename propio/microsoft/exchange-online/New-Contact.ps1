$DB_GPR = Import-CSV "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\AUSA-ListadoGrupos.csv"
$DB_ALL = @()
foreach($IT_GPR in $DB_GPR){
    $DB_MBM = $IT_GPR.'MiembrosdelGrupo' -split ','
    foreach($IT_MBM in $DB_MBM){ 
        if(($IT_MBM.Trim()) -and ($IT_MBM.Trim() -like '*@*') -and ($IT_MBM.Trim() -notlike '*ausasolucioneslogisticas*')){
            $IT_CNT_MAL = $IT_MBM.Trim()
            $DB_ALL += $IT_CNT_MAL
        }
    }
}

$DB_ALL = $DB_ALL | Sort-Object | Get-Unique

foreach($IT_CNT in $DB_ALL){ 
    $IT_CNT_MAL = $IT_CNT
    $IT_CNT_NAM = $IT_CNT_MAL.Substring(0,$IT_CNT_MAL.IndexOf('@')) + (($IT_CNT_MAL.Substring($IT_CNT_MAL.IndexOf('@'))).Substring(0,($IT_CNT_MAL.Substring($IT_CNT_MAL.IndexOf('@'))).IndexOf('.'))).Replace('@','-')
    $IT_CNT_ALI = $IT_CNT_NAM
    $IT_CNT_NAM

    New-MailContact -Name $IT_CNT_NAM -ExternalEmailAddress $IT_CNT_MAL -Alias $IT_CNT_ALI
}