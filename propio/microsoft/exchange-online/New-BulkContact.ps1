Clear
$DB_GPR = Import-CSV "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\AUSA-ListadoGrupos.csv"
$DB_ALL = @()
foreach($IT_GPR in $DB_GPR){
    $DB_MBM = $IT_GPR.MiembrosdelGrupo -split ','
    $DB_GPR = $IT_GPR.'Direccion de correp'
    foreach($IT_MBM in $DB_MBM){ 
        #$DB_ALL += $IT_MBM
        $INT_REC = Get-Recipient $IT_MBM -ErrorAction SilentlyContinue
        if(!$INT_REC){
            If($IT_MBM -like '*@ausa.com.pe'){
                'Interno, Fallido, ' + $DB_GPR + ', ' + $IT_MBM
                'Interno, Fallido, ' + $DB_GPR + ', ' + $IT_MBM | Export-Csv "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\Registro.csv" -Append
            }
            else{
                'Externo, Fallido, ' + $DB_GPR + ', ' + $IT_MBM
                'Externo, Fallido, ' + $DB_GPR + ', ' + $IT_MBM | Export-Csv "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\Registro.csv" -Append
            }
        }
        else{
            If($IT_MBM -like '*@ausa.com.pe'){
                'Interno, Correcto, ' + $DB_GPR + ', ' + $IT_MBM
                'Interno, Correcto, ' + $DB_GPR + ', ' + $IT_MBM | Export-Csv "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\Registro.csv" -Append
            }
            else{
                'Externo, Correcto, ' + $DB_GPR + ', ' + $IT_MBM
                'Externo, Correcto, ' + $DB_GPR + ', ' + $IT_MBM | Export-Csv "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\Registro.csv" -Append
            }
        }
    }
}