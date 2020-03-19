$mycredentials = Get-Credential

Send-MailMessage -To "rodolfo.castelo@cmscloud.pe" -SmtpServer "smtp.office365.com" -Credential $mycredentials -UseSsl "Actualizacion de Informacion" -Port "587" -Body "Como parte de la actualizacion del Sistema se le solicita por favor, responda este mensaje con la siguiente informacion: <br><br>Numero de RPM:<br>Numero movil personal (Opcional):<br>Anexo asignado: <br><br><b>Muchas gracias.<br><br> Mensaje autogenerado de Office 365</b>" -From "rodolfo.castelo@tcreatividad.com" -BodyAsHtml

#-Cc "simon@otherdomain.com" -Attachments "d:\logs\log1.txt","d:\logs\log2.txt","d:\logs\log3.log"