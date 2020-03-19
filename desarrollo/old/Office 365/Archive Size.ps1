$ArchiveUsers = Get-Mailbox | Where-Object {($_.ArchiveStatus -eq 'Active')} | Select PrimarySmtpAddress

foreach($User in $ArchiveUsers){
    $UPN = $User.PrimarySmtpAddress
    Get-MailboxFolderStatistics -Archive -Identity $UPN| 
    Select Name,FolderPath,ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,Identity
}