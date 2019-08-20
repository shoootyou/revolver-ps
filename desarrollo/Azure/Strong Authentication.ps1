#Create the StrongAuthenticationRequirement object and insert required settings
$mf= New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$mf.RelyingParty = "*"
$mfa = @($mf)
#Enable MFA for a user
Set-MsolUser -UserPrincipalName rodolfo.castelo@tecnofor.pe -StrongAuthenticationRequirements $mfa
 


#Enable MFA for all users (use with CAUTION!)
Get-MsolUser -All | Set-MsolUser -StrongAuthenticationRequirements $mfa
 
#Disable MFA for a user
$mfa = @()
Set-MsolUser -UserPrincipalName aaron.beverly@365lab.net -StrongAuthenticationRequirements $mfa



Get-MsolUser -UserPrincipalName rodolfo.castelo@tecnofor.pe | Select *
