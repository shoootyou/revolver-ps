##########################################################################
##                     GAL for Tecnofor Perú
##########################################################################

$TCNF = Get-User | Where-Object {
($_.UserPrincipalName -like "*@tecnofor.pe") -or 
($_.UserPrincipalName -like "*@tecnofor.com.pe") -or 
($_.UserPrincipalName -like "*@mvpconsulting.pe") -and 
($_.UserPrincipalName -notlike "ventascorporativas*") -and 
($_.UserPrincipalName -notlike "jose.castillo*")
} | Select *
foreach($UserTCNF in $TCNF){
    Set-Mailbox -Identity $UserTCNF.UserPrincipalName -AddressBookPolicy "ABP_Tecnofor" | Out-Null
    #Set-Mailbox -Identity $UserTCNF.UserPrincipalName -CustomAttribute1 "Tecnofor"  | Out-Null
}

##########################################################################
##                     GAL for TCreatividad
##########################################################################

$TC2=Get-Mailbox | Where-Object {($_.UserPrincipalName -like "*@tcreatividad.com")   -or 
($_.UserPrincipalName -like "*@tcrea.pe") -or
($_.UserPrincipalName -like "*@fasscorp.com.pe")  -or 
($_.UserPrincipalName -like "*@fassil.com.pe")  -and 
($_.Name -notlike "*SMO-*")  -and 
($_.Name -notlike "*Tecnofor*") -and 
($_.Name -notlike "*Aula*")} | Select *
foreach($UserTC2 in $TC2){
    Set-Mailbox -Identity $UserTC2.UserPrincipalName -AddressBookPolicy "ABP_Creatividad"  | Out-Null
    #Set-Mailbox -Identity $UserTC2.UserPrincipalName -CustomAttribute1 "Fasscorp"  | Out-Null
}
##########################################################################
##                     GAL for Consultoria PYM
##########################################################################
$CONS=Get-User | Where-Object {($_.UserPrincipalName -like "*@consultoria*")} | Select *
foreach($UserCON in $CONS){
    Set-Mailbox -Identity $UserCON.UserPrincipalName -CustomAttribute1 "ConsultoriaPyM" -AddressBookPolicy "ABP_Consultoria"   | Out-Null
}

##########################################################################
##                     GAL for Administración
##########################################################################
$ADM= Get-User | Where-Object {
($_.UserPrincipalName -like "*@administracion.com.pe") -or 
($_.Name -like "*jose.castillo*")
} | select *
foreach($UserADM in $ADM){
    Set-Mailbox -Identity $UserADM.UserPrincipalName -AddressBookPolicy "ABP_Administracion"   | Out-Null
}


<##########################################################################
##                     General Commands
##########################################################################
Get-AddressBookPolicy -Identity ABP_Gerencial | fl *
Get-GlobalAddressList -Identity GAL_Consultoria
Remove-AddressBookPolicy -Identity ABP_Gerencial

Get-Mailbox | Where-Object {($_.CustomAttribute1 -eq "")}

#>