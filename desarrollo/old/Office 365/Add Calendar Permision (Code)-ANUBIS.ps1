$CSV_PATH = Import-CSVPath
$GBL_IMP_CSV = Import-Csv -Path $CSV_PATH -Delimiter ';'
foreach($GBL_IMP_USR in $GBL_IMP_CSV){
    
    $IF_TMP_01 = Get-MailboxFolderStatistics -Identity $GBL_IMP_USR.Emisor | Where-Object {$_.Name -eq 'Calendario'} | Select Name
    if($IF_TMP_01){
        Write-Host "----------------------------------------------------------------------------------------------------------------"
        Write-Host "Escribiendo permiso para " $RCP_Identity "en el usuario" $EMS_Identity "| Español"
        Write-Host "----------------------------------------------------------------------------------------------------------------"
        $EMS_Identity = $GBL_IMP_USR.Emisor + ':\Calendario'
        $RCP_Identity = $GBL_IMP_USR.Receptor + '@s10peru.com'
        Set-MailboxFolderPermission -Identity $EMS_Identity -User $RCP_Identity -AccessRights Owner
        Write-Host "----------------------------------------------------------------------------------------------------------------"
    }
    else{
        Write-Host "----------------------------------------------------------------------------------------------------------------"
        Write-Host "Escribiendo permiso para " $RCP_Identity "en el usuario" $EMS_Identity "| Inglés "
        Write-Host "----------------------------------------------------------------------------------------------------------------"
        $EMS_Identity = $GBL_IMP_USR.Emisor + ':\Calendar'
        $RCP_Identity = $GBL_IMP_USR.Receptor + '@s10peru.com'
        Set-MailboxFolderPermission -Identity $EMS_Identity -User $RCP_Identity -AccessRights Owner
        Write-Host "----------------------------------------------------------------------------------------------------------------"

    }
}

<#

#Add-MailboxFolderPermission -Identity ayla@contoso.com:\Marketing -User ed@contoso.com -AccessRights Owner



#Get-MailboxFolderStatistics -Identity dramirez@s10peru.com | Where-Object {$_.Name -eq 'Calendario'}



Set-MailboxFolderPermission -Identity rsarayasi@s10peru.com:\Calendario -User emendiola@s10peru.com -AccessRights Owner




Get-MailboxFolderPermission -Identity adupont@s10peru.com:\Calendario



$CSV_PermisionAll = Import-Csv -Path C:\Users\Rodolfo\Desktop\Permisos.csv -Delimiter ';'
foreach($PermisionUSR in $CSV_PermisionAll){
        $EMS_Identity = $PermisionUSR.Emisor + '@s10peru.com:\Calendario'
        Get-MailboxFolderPermission -Identity rsarayasi@s10peru.com:\Calendario
        Write-Host "-------" $PermisionUSR.Emisor
}


Set-MailboxFolderPermission 


#>