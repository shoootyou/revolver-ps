Get-MsolUser -MaxResults 10000 | Where {$_.UserprincipalName -like '*socorro*'}

$MBX = 'juan.hernandez@dinet.com.pe'
$NEW = 'juan.hernandez@dinet.com.pe'


Get-MSOlUser -UserPrincipalName 'op_sist_unilever@dinet.com.pe'



Get-MsolUser -UserPrincipalName $MBX | Select *time,*ID*
Get-MsolUser -UserPrincipalName $NEW | Select *time,*ID*

Remove-MsolUser -UserPrincipalName $NEW
Get-MsolUser -ReturnDeletedUsers -UserPrincipalName $NEW
Remove-MsolUser -RemoveFromRecycleBin -UserPrincipalName $NEW

Set-MsolUserPrincipalName -UserPrincipalName $MBX -NewUserPrincipalName $NEW
Get-MSOLUser -UserPrincipalName $NEW
Set-MsolUser -UserPrincipalName $NEW -ImmutableId '5mgZTWU5JUiHRs1rtdWzzg=='

Get-MSOLUser -UserPrincipalName $NEW | Select *time*,*id*