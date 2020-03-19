#$DB_ARC = Get-Mailbox -Archive
foreach($USR in $DB_ARC){
    $UPN = $USR.Alias + '@ausa.com.pe'
    Set-Mailbox $UPN -ArchiveName "PST en la nube"
    Get-Mailbox $UPN | Select Alias,ArchiveName
}