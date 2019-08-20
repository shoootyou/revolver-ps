$Date = Get-Date
$USR = 'fmisari'
$DOM = '@americatel.com.pe'

$USR_DOM = $USR + $DOM
$Date
Search-Mailbox $USR_DOM -SearchDumpsterOnly -TargetMailbox "Discovery Search Mailbox" -TargetFolder $USR -LogLevel Full
$Date
