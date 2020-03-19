$aadUser = Get-AzureADUser -ObjectId $USR.UserPrincipalName
$aadUser | Select -ExpandProperty ExtensionProperty
$aadUser.ToJson()
$aadUser | Get-Member