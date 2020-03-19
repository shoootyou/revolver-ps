Get-MsolUser -ReturnDeletedUsers | Where {
($_.UserPrincipalName -notlike '*cubica.com.pe') -and 
($_.UserPrincipalName -notlike '*sausac.com.pe') -and 
($_.UserPrincipalName -notlike '*IBM*')} | Select UserprincipalName,ImmutableId,LastDirSyncTime,LastPasswordChangeTimestamp,Licenses,ObjectId,ProxyAddresses,SoftDeletionTimestamp | Sort SoftDeletionTimestamp