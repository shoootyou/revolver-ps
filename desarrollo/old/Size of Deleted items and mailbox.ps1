Write-Host "Gathering Stats, Please Wait.."
$DB = Import-Csv -Path E:\Users\Rodolfo\Desktop\centria.csv

$MailboxSizes = @()

foreach($usuario in $DB){
    $Mailboxes = Get-Mailbox $usuario.mail | Select UserPrincipalName, identity

<#$Mailboxes = Get-Mailbox -ResultSize Unlimited | Where {($_.RecipientTypeDetails -eq 'UserMailbox') -and
($_.UserPrincipalName -like '*americatel.com.pe') -and
($_.UserPrincipalName -notlike 'migracion*')} | Select UserPrincipalName, identity
#>



foreach ($Mailbox in $Mailboxes) {
    $USR_01 = $Mailbox.UserPrincipalName
         
    $ObjProperties = New-Object PSObject
               
    $MailboxStats = Get-MailboxStatistics $Mailbox.UserPrincipalname | Select LastLogonTime, TotalItemSize, ItemCount

    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "UserPrincipalName" -Value $Mailbox.UserPrincipalName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Mailbox Size" -Value $MailboxStats.TotalItemSize
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Mailbox Item Count" -Value $MailboxStats.ItemCount
    
    $MailboxSizes += $ObjProperties
    
 
}   
         
Write-Host "Correct process to" $usuario.mail
}
       
$MailboxSizes | Out-GridView -Title "Mailbox and Archive Sizes"
#$MailboxSizes | Export-Csv -Path C:\Users\Rodolfo\Desktop\Reporte.csv