$DB_FIR = Import-Csv .\ausa-firma-update.csv
$BA_FIR = Get-Content .\signature.html -Encoding UTF8

foreach($USR in $DB_FIR){
    $BA_FIR.ToString().Replace('[Nombres y Apellidos]',$USR.Name).Replace("[Cargo]",$USR.Position).Replace("[Telefono_]",$USR.Phone).Replace("[Ext_]",$USR.Ext).Replace("[Celular_]",$USR.Mobile) | Set-Content (".\firmas\" + $USR.UserName +  ".html") -Encoding UTF8
    Set-MailboxMessageConfiguration -Identity $USR.UserName -SignatureHtml (Get-Content (".\firmas\" + $USR.UserName +  ".html") -Encoding UTF8) -AutoAddSignature $true
    Write-Host $USR.UserName     
}