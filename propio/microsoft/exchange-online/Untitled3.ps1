#Set-Mailbox $USR.UserPrincipalName -CustomAttribute1 $USR.UserPrincipalName
Get-Mailbox Alexander.Porles@spsa.pe | Select UserPrincipalName,CustomAttribute1
Get-Mailbox Santiago.Chavarria@spsa.pe | Select UserPrincipalName,CustomAttribute1

Get-MsolUser -UserPrincipalName Santiago.Chavarria@spsa.pe  | Select UserPrincipalName,Immuta*
Get-MsolUser -UserPrincipalName Alexander.Porles@spsa.pe  | Select UserPrincipalName,Immuta*


Santiago.Chavarria@spsa.pe