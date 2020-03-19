function Get-MailboxFolderSize{
    param(
            [bool]$ExportFile = $false,
            [string]$UserPrincipalName 
    )
    if(!$UserPrincipalName){
        $GBL_EXP_USR = Get-MsolUser | Where-Object {($_.isLicensed -eq 'True')} | select *
    }
    else{

        $GBL_EXP_USR = Get-MsolUser | Where-Object {$_.UserPrincipalName -eq $UserPrincipalName}| Where-Object {($_.isLicensed -eq 'True')} | select *
    
    }
    
    $OBJ_OUT = @()
    foreach($IN_EXP_USR in $GBL_EXP_USR){

        $IN_DET_USR = Get-MailboxFolderStatistics -Identity $IN_EXP_USR.UserPrincipalName | 
        Select Name,FolderPath,ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,Identity
        
            foreach($IN_FDR_MBX in $IN_DET_USR){
    
            $OBJ_OUT_PRO = New-Object PSObject

            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Nombre' -Value $IN_FDR_MBX.Name
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Ruta' -Value $IN_FDR_MBX.FolderPath
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Elementos' -Value $IN_FDR_MBX.ItemsInFolder
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Tamaño' -Value $IN_FDR_MBX.FolderSize
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Elementos en carpeta y subcarpeta' -Value $IN_FDR_MBX.ItemsInFolderAndSubfolders
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Tamaño de Carpeta y subcarpeta' -Value $IN_FDR_MBX.FolderAndSubfolderSize
            Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Identidad' -Value $IN_EXP_USR.UserPrincipalName
            
            $OBJ_OUT += $OBJ_OUT_PRO
 
            }

    }

    $OBJ_OUT | Out-GridView -Title "Información Personal"
    if($ExportFile){
        $IN_USR_PRO = $env:USERPROFILE
        $OUT_CSV = $IN_USR_PRO + '\Desktop\Exportacion.csv'
        $OBJ_OUT | Export-Csv $OUT_CSV
    }
    else{
    }

}

Get-MailboxFolderSize -UserPrincipalName mramirez@s10peru.com