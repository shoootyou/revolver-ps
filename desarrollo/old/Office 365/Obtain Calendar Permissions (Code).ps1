$GBL_CSV_PATH = 'C:\Users\Rodolfo\Desktop\Permissions.csv' 
Out-File -FilePath $GBL_CSV_PATH -InputObject "Emisor,Lector,Ubicación,Permiso" -Encoding UTF8 

$GBL_USR_ALL = Get-User -ResultSize Unlimited  | Where-Object {($_.RecipientType -eq 'UserMailbox') -and ($_.Name -notlike '*{*')} | Select DisplayName,UserPrincipalName,WindowsEmailAddress,WindowsLiveID,MicrosoftOnlineServicesID

foreach($TMP_USR in $GBL_USR_ALL){
    Write-Host "Processing $($TMP_USR.DisplayName)..." 
    $IF_TMP_01 = Get-MailboxFolderStatistics -Identity $TMP_USR.UserPrincipalName | Where-Object {$_.Name -eq 'Calendario'} | Select Name
    if($IF_TMP_01){
        $TMP_INT_PATH = $TMP_USR.UserPrincipalName + ':\Calendario'
        $TMP_INT_PRM = Get-MailboxFolderPermission -Identity $TMP_INT_PATH | Select FolderName,User,AccessRights
    }
    else{
        $TMP_INT_PATH = $TMP_USR.UserPrincipalName + ':\Calendar'
        $TMP_INT_PRM = Get-MailboxFolderPermission -Identity $TMP_INT_PATH | Select FolderName,User,AccessRights
    }


    foreach($TMP_PER in $TMP_INT_PRM){
        Out-File -FilePath $GBL_CSV_PATH -InputObject "$($TMP_USR.UserPrincipalName),$($TMP_PER.User),$($TMP_PER.FolderName),$($TMP_PER.AccessRights)" -Encoding UTF8 -append 
    }
}


#Get-MailboxFolderPermission -Identity adupont@s10peru.com:\Calendario | Select FolderName,User,AccessRights