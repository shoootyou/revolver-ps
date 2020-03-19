Get-MessageTrace -RecipientAddress cuenta_olvido@tecnologiaycreatividad.onmicrosoft.com -StartDate "10/22/2015" -EndDate "10/28/2015" | Select * | Export-CSV C:\Users\Rodolfo\Desktop\1.csv

Get-Command *Rule* | Where-Object {$_.Source -eq "tmp_ifllf4r0.j0j"}

Get-TransportRule | Where-Object {$_.Name -like "*Rebote*"} | Select  * -ExpandProperty SenderDomainIs

New-TransportRule "Bloqueo a nivel de Usuario" -From "centrum@pucp.edu.pe" -RejectMessageReasonText "El mensaje no superó con éxito los análisis de spam y malware a los que fue sometido."


#This is what I do (But I do it in a script and import the allowed domains in bulk from a txt file)
#First, attach the transport rule to a variable called $Rule

$Rule = get-transportrule "Bloqueo a nivel de Usuario"

#Get the SenderDomainIs property and attach it to a variable called $SenderDomain

$SenderDomains2 = $Rule.FROM

#Piping that to the command shows you what domains are currently allowed.

$SenderDomains2 = @()

#Use the += operator to add another domain to the $SenderDomains2 Variable.

$Dominios = Import-Csv C:\Users\Rodolfo\Desktop\Usuarios.csv

foreach($Dominio in $Dominios){

$SenderDomains2 += $Dominio.From

}
#Here you can pipe it again to see it's been added to the list.


$SenderDomains2


Set-TransportRule "Bloqueo a nivel de Usuario" -RejectMessageEnhancedStatusCode "5.1.10"


##############################################################################################################################

Get-Mailbox cuenta_olvido@tecnologiaycreatividad.onmicrosoft.com | Select * -ExpandProperty EmailAddresses




Set-Mailbox cuenta_olvido@tecnologiaycreatividad.onmicrosoft.com -EmailAddresses @{add="smtp:inegocios@consultoriapym.com","smtp:roncoperu@administracion.com.pe"}