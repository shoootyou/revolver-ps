$DB_GPR = Import-CSV .\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\AUSA-ListadoGrupos.csv
$TS_OUT = '.\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\Creaciongrupos.log'
Start-Transcript -path $TS_OUT -append
$DB_ALL = @()
foreach($IT_GPR in $DB_GPR){
    $IT_GPR_NAM = $IT_GPR.'Nombre del Grupo'
    $IT_GPR_DIR = $IT_GPR.'Direccion de correp'
    if(!$IT_GPR_NAM){
        $IT_GPR_NAM = $IT_GPR_DIR.Substring(0,$IT_GPR_DIR.IndexOf('@')) + (($IT_GPR_DIR.Substring($IT_GPR_DIR.IndexOf('@'))).Substring(0,($IT_GPR_DIR.Substring($IT_GPR_DIR.IndexOf('@'))).IndexOf('.'))).Replace('@','-')
    }
   New-DistributionGroup -Name $IT_GPR_NAM -DisplayName $IT_GPR_NAM -PrimarySmtpAddress $IT_GPR_DIR -Type Distribution
}
Stop-Transcript