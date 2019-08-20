Clear
$DB_GRP_PAT = "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\Documentacion\CSV-20190426-ListadoGrupos.csv"
$DB_GRP_CSV = Import-Csv $DB_GRP_PAT
$CON_GPR = 1
foreach($FOR_GRP_CSV in $DB_GRP_CSV){
    
    If($FOR_GRP_CSV.MiembrosdelGrupo){
        
        $IT_GPR_NAM = $FOR_GRP_CSV.'Nombre del Grupo'
        $IT_GPR_DIR = $FOR_GRP_CSV.'Direccion de correp'
        Write-Progress -Activity “Agregando usuarios a los grupos" -status “Procesando el grupo $IT_GPR_DIR” -percentComplete ($CON_GPR / $DB_GRP_CSV.count*100) -Id 500
        if(!$IT_GPR_NAM){
            $IT_GPR_NAM = $IT_GPR_DIR.Substring(0,$IT_GPR_DIR.IndexOf('@')) + (($IT_GPR_DIR.Substring($IT_GPR_DIR.IndexOf('@'))).Substring(0,($IT_GPR_DIR.Substring($IT_GPR_DIR.IndexOf('@'))).IndexOf('.'))).Replace('@','-')
        }
        
        $DB_GRP_MBM = $FOR_GRP_CSV.MiembrosdelGrupo -split ','
        $CON_MBM = 1
        foreach($FOR_MBM in $DB_GRP_MBM){
            Write-Progress -Activity “Agregando usuarios al grupo $IT_GPR_DIR" -status “Adicionando usuario $FOR_MBM” -percentComplete ($CON_MBM / $DB_GRP_MBM.count*100) -ParentId 500
            #Write-Host 'Agregando a usuario ' $FOR_MBM -ForegroundColor Gray
            #Invoke-Expression "Add-DistributionGroupMember -Identity $IT_GPR_DIR -Member $FOR_MBM" -ErrorAction SilentlyContinue -ErrorVariable ErrorDistribution 
            #$ErrorDistribution
            if($FOR_MBM -like '*ausa.com.pe'){
                $VAL = Get-Mailbox $FOR_MBM -ErrorAction SilentlyContinue
            }
            else{
                $VAL = Get-Recipient $FOR_MBM -ErrorAction SilentlyContinue
            }
            #$VAL
            If($VAL){
                Add-DistributionGroupMember -Identity $IT_GPR_DIR -Member $FOR_MBM -ErrorAction SilentlyContinue
            }
            else{
                #Write-Host '------------------------------------------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
                Write-Host $FOR_MBM -ForegroundColor Yellow
            }
            $CON_MBM++
        }
        <#
        $VAL_ONL_CON = Get-DistributionGroupMember $IT_GPR_DIR
        If(!$VAL_ONL_CON.Count -and !$VAL_ONL_CON[1]){
            $VAL_ONL_CON = $VAL_ONL_CON[0].Name
        }
        If($VAL_ONL_CON.Count -ne $DB_GRP_MBM.Count){
            Write-Host '==================================================================================================================================================================' -ForegroundColor Cyan
            Write-Host $IT_GPR_NAM -ForegroundColor Cyan
            Write-Host $IT_GPR_DIR -ForegroundColor Cyan
            Write-Host '------------------------------------------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
            Write-Host 'Alerta, algo sucedió. Aquí la lista de los usuarios' -ForegroundColor Yellow
            Write-Host '------------------------------------------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
            Write-Host 'Lista de usuarios en nube : ' $VAL_ONL_CON -ForegroundColor Yellow
            Write-Host '------------------------------------------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
            Write-Host 'Lista de usuarios en CSV  : '$DB_GRP_MBM -ForegroundColor Yellow
        }
        else{
            #Write-Host '------------------------------------------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
            #Write-Host 'Completado correctamente' -ForegroundColor Green
        }#>
        $CON_GPR++
    }
}