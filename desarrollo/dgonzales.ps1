David Gonzalez

Get-Mailbox rcruz@ahr.com.pe | Select guid
Get-Mailbox dgonzales@ahr.com.pe | select EmailAddresses 

	
Set-Mailbox dgonzales@ahr.com.pe -EmailAddresses @{Remove="dgonzales@ahr.com.pe"}
Set-Mailbox dgonzales@ahr.com.pe -
| Out-File E:\Users\Rodolfo\Desktop\dgonzales.txt


Get-Mailbox | Where {$_.Alias -like '*gonzal*'}

Get-Command *mailbox*

New-MailboxRestoreRequest -SourceMailbox f66df0f6-7a0d-4f89-85d8-85abb68206b4


Set-MsolUser -UserPrincipalName dgonzales3@ahr.com.pe -ImmutableId 'qVed0Ecrs0+DX4PJzsxtWA=='
Get-MsolUser -UserPrincipalName dgonzales2@ahr.com.pe | Select ImmutableId 
Set-MsolUser -UserPrincipalName dgonzales3@ahr.com.pe -ImmutableId 'AAAAAAAAAAAAAAAAAAAAAA=='