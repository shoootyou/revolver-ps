function Get-MailboxFolderSize{
    <#
        .DESCRIPTION
        Verifica el tamaño de las diversas carpetas de un buzón específico o de todos los miembros de la organización que posean licencia.

        .PARAMETER ExportFile
        Valida si se exportará o no, un csv con la información obtenida.

        .PARAMETER UserPrincipalName
        Si en caso especifica un UPN en particular, se mostrará la información de dicho usuario.

        .EXAMPLE
        Get-MailboxFolderSize -ExportFile:$true 
        
        Sirve para poder autorizar la exportación en un archivo CSV
        el mismo será ubicado en el escritorio de su usuario actual.
        
        .EXAMPLE
        Get-MailboxFolderSize -UserPrincipalName prueba@midominio.com 

        Sirve para poder mostrar la información de un usuario en particular,
        si en caso desea obtener la información de todos, no especifique el UPN.
               
    #>


    param(
            [bool]$ExportFile = $false,
            [bool]$ViewGrid = $true,
            [string]$UserPrincipalName,
            [string]$Filename = 'Exportacion',
            [bool]$Archive = $false
    )
    if(!$UserPrincipalName){
        $GBL_EXP_USR = Get-MsolUser | Where-Object {($_.isLicensed -eq 'True')} | select *
    }
    else{

        $GBL_EXP_USR = Get-MsolUser | Where-Object {$_.UserPrincipalName -eq $UserPrincipalName}| Where-Object {($_.isLicensed -eq 'True')} | select *
    
    }
    
    $OBJ_OUT = @()
    foreach($IN_EXP_USR in $GBL_EXP_USR){

        if($Archive){
            $IN_DET_USR = Get-MailboxFolderStatistics -Archive -Identity $IN_EXP_USR.UserPrincipalName | 
            Select Name,FolderPath,ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,Identity
        }
        else{

            $IN_DET_USR = Get-MailboxFolderStatistics -Identity $IN_EXP_USR.UserPrincipalName | 
            Select Name,FolderPath,ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,Identity
    
        }

            foreach($IN_FDR_MBX in $IN_DET_USR){
    
            $OBJ_OUT_PRO = New-Object PSObject

            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Nombre' -Value $IN_FDR_MBX.Name
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Ruta' -Value $IN_FDR_MBX.FolderPath
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Elementos' -Value $IN_FDR_MBX.ItemsInFolder

#-----------------------------------------------------------------------------------------------------------------------------------------------------------
        $SRC_01_SIZE = $IN_FDR_MBX.FolderSize

        $SRC_01_SIZE_STRG = $SRC_01_SIZE.ToString()
        $SRC_01_SIZE_STRG_POS_1 = $SRC_01_SIZE_STRG.IndexOf('(')
        $SRC_01_SIZE_STRG_PH_1 = $SRC_01_SIZE_STRG.Substring($SRC_01_SIZE_STRG_POS_1+1)
        $SRC_01_SIZE_STRG_POS_2 = $SRC_01_SIZE_STRG_PH_1.IndexOf(')')
        $SRC_01_SIZE_STRG_PH_2 = $SRC_01_SIZE_STRG_PH_1.Substring(0,$SRC_01_SIZE_STRG_POS_2)
        $SRC_LEN = $SRC_01_SIZE_STRG_PH_2.length
        if($SRC_LEN -gt 7){
            $SRC_01_SIZE_STRG_POS_3 = $SRC_01_SIZE_STRG_PH_2.IndexOf('bytes')
            [string]$SRC_01_SIZE_STRG_PH_3 = $SRC_01_SIZE_STRG_PH_2.Substring(0, $SRC_01_SIZE_STRG_POS_3-1)

            $SRC_01_PART_01 = [string]$SRC_01_SIZE_STRG_PH_3.split(',')[0]
            $SRC_01_PART_02 = [string]$SRC_01_SIZE_STRG_PH_3.split(',')[1]
            $SRC_01_PART_03 = [string]$SRC_01_SIZE_STRG_PH_3.split(',')[2]
            $SRC_01_PART_04 = [string]$SRC_01_SIZE_STRG_PH_3.split(',')[3]
            $SRC_01_SIZE_STRG_PH_4 =$SRC_01_PART_01 + $SRC_01_PART_02 + $SRC_01_PART_03 + $SRC_01_PART_04

            [int64]$SRC_01_SIZE_INTE = [convert]::ToInt64($SRC_01_SIZE_STRG_PH_4, 10)
            $SRC_01_SIZE_OUT = [int64]$SRC_01_SIZE_INTE
        }
        else{
            $SRC_01_SIZE_STRG_POS_3 = $SRC_01_SIZE_STRG_PH_2.IndexOf('bytes')
            [string]$SRC_01_SIZE_STRG_PH_3 = $SRC_01_SIZE_STRG_PH_2.Substring(0, $SRC_01_SIZE_STRG_POS_3-1)
            [int]$SRC_01_SIZE_INTE = [convert]::ToInt32($SRC_01_SIZE_STRG_PH_3, 10)
            $SRC_01_SIZE_OUT = [int]$SRC_01_SIZE_INTE
        }
                  
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Tamaño' -Value $SRC_01_SIZE_OUT
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Elementos en carpeta y subcarpeta' -Value $IN_FDR_MBX.ItemsInFolderAndSubfolders
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
        $SRC_02_SIZE = $IN_FDR_MBX.FolderAndSubfolderSize

        $SRC_02_SIZE_STRG = $SRC_02_SIZE.ToString()
        $SRC_02_SIZE_STRG_POS_1 = $SRC_02_SIZE_STRG.IndexOf('(')
        $SRC_02_SIZE_STRG_PH_1 = $SRC_02_SIZE_STRG.Substring($SRC_02_SIZE_STRG_POS_1+1)
        $SRC_02_SIZE_STRG_POS_2 = $SRC_02_SIZE_STRG_PH_1.IndexOf(')')
        $SRC_02_SIZE_STRG_PH_2 = $SRC_02_SIZE_STRG_PH_1.Substring(0,$SRC_02_SIZE_STRG_POS_2)
        $SRC_LEN = $SRC_02_SIZE_STRG_PH_2.length
        if($SRC_LEN -gt 7){
            $SRC_02_SIZE_STRG_POS_3 = $SRC_02_SIZE_STRG_PH_2.IndexOf('bytes')
            [string]$SRC_02_SIZE_STRG_PH_3 = $SRC_02_SIZE_STRG_PH_2.Substring(0, $SRC_02_SIZE_STRG_POS_3-1)

            $SRC_02_PART_01 = [string]$SRC_02_SIZE_STRG_PH_3.split(',')[0]
            $SRC_02_PART_02 = [string]$SRC_02_SIZE_STRG_PH_3.split(',')[1]
            $SRC_02_PART_03 = [string]$SRC_02_SIZE_STRG_PH_3.split(',')[2]
            $SRC_02_PART_04 = [string]$SRC_02_SIZE_STRG_PH_3.split(',')[3]
            $SRC_02_SIZE_STRG_PH_4 =$SRC_02_PART_01 + $SRC_02_PART_02 + $SRC_02_PART_03 + $SRC_02_PART_04

            [int64]$SRC_02_SIZE_INTE = [convert]::ToInt64($SRC_02_SIZE_STRG_PH_4, 10)
            $SRC_02_SIZE_OUT = [int64]$SRC_02_SIZE_INTE
        }
        else{
            $SRC_02_SIZE_STRG_POS_3 = $SRC_02_SIZE_STRG_PH_2.IndexOf('bytes')
            [string]$SRC_02_SIZE_STRG_PH_3 = $SRC_02_SIZE_STRG_PH_2.Substring(0, $SRC_02_SIZE_STRG_POS_3-1)
            [int]$SRC_02_SIZE_INTE = [convert]::ToInt32($SRC_02_SIZE_STRG_PH_3, 10)
            $SRC_02_SIZE_OUT = [int]$SRC_02_SIZE_INTE
        }
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Tamaño de Carpeta y subcarpeta' -Value $SRC_02_SIZE_OUT
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Identidad' -Value $IN_EXP_USR.UserPrincipalName
            
            $OBJ_OUT += $OBJ_OUT_PRO
 
            }

    }

    if($ViewGrid){
        $OBJ_OUT | Out-GridView -Title "Mailbox's folders size"
    }
    else{
    }
    
    if($ExportFile){
        if($Archive){
            $IN_USR_PRO = $env:USERPROFILE
            $OUT_CSV = $IN_USR_PRO + '\Desktop\' + $Filename + ' archive.csv'
            $OBJ_OUT | Export-Csv $OUT_CSV
        }
        else{
            $IN_USR_PRO = $env:USERPROFILE
            $OUT_CSV = $IN_USR_PRO + '\Desktop\' + $Filename + '.csv'
            $OBJ_OUT | Export-Csv $OUT_CSV
        }

        
    }
    else{
    }

}

Get-MailboxFolderSize -UserPrincipalName lsarayasi@s10peru.com -ViewGrid:$true -Archive:$false # -ExportFile:$false -Filename 'lsarayasi' 