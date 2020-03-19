Write-Host "Gathering Stats, Please Wait.."
 
$Mailboxes = Get-Mailbox oparedes | Select UserPrincipalName, identity
 
$MailboxSizes = @()
$i = 1
foreach ($Mailbox in $Mailboxes) {
    $USR_01 = $Mailbox.UserPrincipalName
    #Write-Host 'Procesando el usuario: ' $USR_01
    $ObjProperties = New-Object PSObject
               
    $MailboxStats = Get-MailboxFolderStatistics $Mailbox.UserPrincipalname -FolderScope RecoverableItems  | Where {$_.Name -eq 'Deletions'} | Select ItemsInFolder,FolderSize
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "UserPrincipalName" -Value $Mailbox.UserPrincipalName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Deleted Item" -Value $MailboxStats.ItemsInFolder
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Deleted Item Size" -Value $MailboxStats.FolderSize
    $MailboxSizes += $ObjProperties
    #Write-Progress -Activity “Gathering Information” -status “Working on $USR_01” -percentComplete ($i / $Mailboxes.count*100)
    $i++
}             
               
$MailboxSizes | Out-GridView -Title "Mailbox Deleted items and size"
#$MailboxSizes | Export-Csv -Path E:\Users\Rodolfo\Desktop\Reporte.csv


 