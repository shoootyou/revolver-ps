$DB_FIR = Get-Mailbox -ResultSize 1000

foreach($USR in $DB_FIR){
    Set-MailboxMessageConfiguration -Identity $USR.Alias -DefaultFontName Arial
    Write-Host $USR.Alias     
}