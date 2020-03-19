$mail = Read-Host '¿De qué usuario deseas saber el tamaño de Buzón?'
get-mailboxstatistics -Identity $mail | Select displayname, totalitemsize
Get-MailboxStatistics $mail -Archive | Select TotalItemSize, ItemCount
