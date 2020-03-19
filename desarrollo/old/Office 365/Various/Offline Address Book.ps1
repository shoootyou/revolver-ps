#Set-TransportConfig -AddressBookPolicyRoutingEnabled $true
#Get-TransportConfig |Select AddressBookPolicyRoutingEnabled

#Get-Mailbox * | Select RecipientTypeDetails
#New-AddressList -name "AL_Salas" -RecipientFilter {((RecipientTypeDetails -eq 'RoomMailbox') -and (CustomAttribute3 -eq "Room"))} -DisplayName "Listado de Salas"

#Get-mailbox -resultsize unlimited | where {$_.PrimarySMTPaddress -like  '*@administracion.com.pe'} | Set-Mailbox -CustomAttribute1 'Administracion'

#Get-mailbox -resultsize unlimited | where {$_.PrimarySMTPaddress -like  '*@TecnologiayCreatividad.onmicrosoft.com'} | Select PrimarySmtpAddress, CustomAttribute1

#Get-Mailbox * | Select PrimarySMTPAddress, CustomAttribute1 | Sort-Object PrimarySMTPAddress -Descending

#Set-Mailbox veronica.mon@tecnologiaycreatividad.onmicrosoft.com -CustomAttribute1 'CRM'
#Set-Mailbox rossana.reccio@tecnologiaycreatividad.onmicrosoft.com -CustomAttribute1 'CRM'

#New-AddressList -name "AL_Consultoria"  -RecipientFilter {(RecipientType -eq 'UserMailbox') -and (CustomAttribute1 -eq "ConsultoriaPyM")}  –DisplayName “Consultoria P y M”
#New-AddressList -name "AL_Tecnofor"  -RecipientFilter {(RecipientType -eq 'UserMailbox') -and (CustomAttribute1 -eq "Tecnofor")}  –DisplayName “Tecnofor Perú”
#New-AddressList -name "AL_Creatividad"  -RecipientFilter {(RecipientType -eq 'UserMailbox') -and ((CustomAttribute1 -eq "TCreatividad") -or (CustomAttribute1 -eq "Fasscorp"))}  –DisplayName “Tecnología y Creatividad”
#New-AddressList -name "AL_Administracion"  -RecipientFilter {(RecipientType -eq 'UserMailbox') -and (CustomAttribute1 -eq "Administracion")}  –DisplayName “Administracion”

#Get-AddressList | Sort-Object Name -Descending

#New-GlobalAddressList -name "GAL_Consultoria" -RecipientFilter {(CustomAttribute1 -eq "ConsultoriaPyM")}
#New-GlobalAddressList -name "GAL_Tecnofor" -RecipientFilter {(CustomAttribute1 -eq "Tecnofor")}
#New-GlobalAddressList -name "GAL_Creatividad" -RecipientFilter {((CustomAttribute1 -eq "TCreatividad") -or (CustomAttribute1 -eq "Fasscorp"))}
#New-GlobalAddressList -name "GAL_Gerencial" -RecipientFilter {((CustomAttribute1 -eq "TCreatividad") -or (CustomAttribute1 -eq "Fasscorp") -or (CustomAttribute1 -eq "ConsultoriaPyM"))}
#New-GlobalAddressList -name "GAL_Administracion" -RecipientFilter {(CustomAttribute1 -eq "Administracion")}

#Remove-GlobalAddressList "GAL_TCreatividad"

#Get-GlobalAddressList "GAL_Creatividad" | Select *

#New-OfflineAddressBook -name "OAB_Consultoria" -AddressLists "GAL_Consultoria"
#New-OfflineAddressBook -name "OAB_Tecnofor" -AddressLists "GAL_Tecnofor"
#New-OfflineAddressBook -name "OAB_Creatividad" -AddressLists "GAL_Creatividad"
#New-OfflineAddressBook -name "OAB_Gerencial" -AddressLists "GAL_Gerencial"
#New-OfflineAddressBook -name "OAB_Administracion" -AddressLists "GAL_Administracion"

#Get-OfflineAddressBook | fl  Name, Schedule | Format-List

#Remove-OfflineAddressBook "OAB_Consultoria"

#New-AddressBookPolicy -name "ABP_Consultoria" -AddressLists "AL_Consultoria" -OfflineAddressBook "\OAB_Consultoria" -GlobalAddressList "\GAL_Consultoria" -RoomList "\AL_Salas"
#New-AddressBookPolicy -name "ABP_Tecnofor" -AddressLists "AL_Tecnofor" -OfflineAddressBook "\OAB_Tecnofor" -GlobalAddressList "\GAL_Tecnofor" -RoomList "\AL_Salas"
#New-AddressBookPolicy -name "ABP_Creatividad" -AddressLists "AL_Creatividad" -OfflineAddressBook "\OAB_Creatividad" -GlobalAddressList "\GAL_Creatividad" -RoomList "\AL_Salas"
#New-AddressBookPolicy -name "ABP_Gerencial" -AddressLists "AL_Creatividad","AL_Tecnofor","AL_Consultoria" -OfflineAddressBook "\OAB_Gerencial" -GlobalAddressList "\GAL_Gerencial" -RoomList "\AL_Salas"
#New-AddressBookPolicy -name "ABP_Administracion" -AddressLists "AL_Administracion" -OfflineAddressBook "\OAB_Administracion" -GlobalAddressList "\GAL_Administracion" -RoomList "\AL_Salas"

#Remove-AddressBookPolicy "ABP_Creatividad"

#Get-AddressBookPolicy

#Get-Mailbox -resultsize unlimited | where {$_.CustomAttribute1 -eq "ConsultoriaPyM"} | Set-Mailbox -AddressBookPolicy "ABP_Consultoria"
#Get-Mailbox -resultsize unlimited | where {$_.CustomAttribute1 -eq "Tecnofor"} | Set-Mailbox -AddressBookPolicy "ABP_Tecnofor"
#Get-Mailbox -resultsize unlimited | where {$_.CustomAttribute1 -eq "TCreatividad"} | Set-Mailbox -AddressBookPolicy "ABP_Creatividad"
#Get-Mailbox -resultsize unlimited | where {$_.CustomAttribute1 -eq "Fasscorp"} | Set-Mailbox -AddressBookPolicy "ABP_Creatividad"
#Get-Mailbox -resultsize unlimited | where {$_.CustomAttribute1 -eq "Administracion"} | Set-Mailbox -AddressBookPolicy "ABP_Administracion"

#Get-Mailbox * | Select Name, PrimarySMTPAddress,AddressBookPolicy

#Set-Mailbox jlizarraga@tcreatividad.com -AddressBookPolicy "ABP_Creatividad"