$DB = Get-MsolUser -MaxResults 10000 | Where {($_.UserPrincipalName -notlike '*onmicrosoft*') -and ($_.UserPrincipalName -notlike '*breca*')}
$MailboxSizes = @()
foreach($Usu in $DB){
    $DOM = $usu.UserPrincipalName.Substring($usu.UserPrincipalName.IndexOf("@")+1)
    $LIC = $usu.IsLicensed
    if($Usu.ProxyAddresses -like '*mail.onmicrosoft.com'){}
    else{
        if($LIC){
            $ObjProperties = New-Object PSObject

            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DisplayName" -Value $usu.DisplayName
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "UserprincipalName" -Value $Usu.UserPrincipalName
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Dominio" -Value $DOM
            if($usu.LastDirSyncTime){
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Synced?" -Value 'Yes'
            }
            if($Usu.ProxyAddresses -like ''){
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "ProxyAddresses" -Value 'No hay'
            }
            else{
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "ProxyAddresses" -Value $Usu.ProxyAddresses
            }
            $MailboxSizes += $ObjProperties
        }
    }


}

$MailboxSizes | Out-GridView -Title "Mailbox Deleted items and size"


#182 + 17
#$MailboxSizes | Export-Csv -Path E:\Users\Rodolfo\Desktop\Reporte.csv
