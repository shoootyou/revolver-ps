$USR = 'salasistemas'
$CAL = '@aris.com.pe:\Calendar'

$USR_CAL = $USR + $CAL

$USR = 'salapantera'
$CAL = '@aris.com.pe:\Calendar'

$USR_CAL = $USR + $CAL

Get-MailboxFolderPermission $USR_CAL
Write-Host '---------------------------------------------------'
Set-MailboxFolderPermission -Identity $USR_CAL -User ktor res@americatel.com.pe -AccessRights Owner
Write-Host '---------------------------------------------------'
Get-MailboxFolderPermission $USR_CAL

Remove-MailboxFolderPermission -Identity $USR_CAL -User nmaidana@americatel.com.pe


