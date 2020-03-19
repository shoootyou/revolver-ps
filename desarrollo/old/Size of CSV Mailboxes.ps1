Write-Host "Gathering Stats, Please Wait.."
$DB = Import-Csv -Path E:\Users\Rodolfo\Desktop\1.csv

$MailboxSizes = @()
$i = 1
foreach($usuario in $DB){
    $USR_01 = $usuario.Mail
    Write-Progress -Activity “Gathering Information” -status “Working on $USR_01” -percentComplete ($i / $DB.count*100)
    
    $Mailboxes = Get-Mailbox $USR_01 | Select UserPrincipalName, identity

    foreach ($Mailbox in $Mailboxes) {
        $USR_01 = $Mailbox.UserPrincipalName
         
        $ObjProperties = New-Object PSObject
               
        $MailboxStats = Get-MailboxStatistics $Mailbox.UserPrincipalname | Select LastLogonTime, TotalItemSize, ItemCount
    
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "UserPrincipalName" -Value $Mailbox.UserPrincipalName
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Mailbox Size" -Value $MailboxStats.TotalItemSize
    
        $MailboxSizes += $ObjProperties

    }             
$i++
}
       
$MailboxSizes | Out-GridView -Title "Mailbox and Archive Sizes"
$MailboxSizes | Export-Csv -Path E:\Users\Rodolfo\Desktop\Reporte.csv