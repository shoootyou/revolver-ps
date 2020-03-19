#New-MailboxSearch -Name 'Testing' -SourceMailboxes proveedor@americatel.com.pe -TargetMailbox proveedor-dsmbx@americatel.com.pe 
#Start-MailboxSearch 'Testing' -Confirm:$false -Force



#Search-Mailbox proveedor@americatel.com.pe -TargetMailbox proveedor-dsmbx -TargetFolder 'Query' -SearchQuery {Received:01/01/1900..10/08/2016} -LogLevel Full -SearchDumpsterOnly 
#Search-Mailbox proveedor@americatel.com.pe -TargetMailbox proveedor-dsmbx -TargetFolder 'Query' -SearchQuery {sent:01/01/1900..10/08/2016} -LogLevel Full -SearchDumpsterOnly 

Get-mailbox Salapiso9-oropendolas | select *
Search-Mailbox proveedor-dsmbx -TargetMailbox proveedor -TargetFolder 'Inbox' -LogLevel Full

Get-MailboxFolderPermission proveedor-dsmbx | Select *

Get-mailbox proveedor-dsmbx | Select identity