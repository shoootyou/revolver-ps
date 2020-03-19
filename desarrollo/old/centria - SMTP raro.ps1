$DB = Get-Mailbox -ResultSize 10000 |where {$_.RecipientTypeDetails -eq 'UserMailbox'} | Select UserPrincipalName,Alias
$i = 1
foreach($User in $DB){
    $INT_USR = $User.UserPrincipalName

    $Int_SIP = Get-Mailbox $INT_USR | Select EmailAddresses

    Write-Progress -Activity “Verificando usuarios” -status “Procesando $INT_USR” -percentComplete ($i / $DB.count*100)

    if(($Int_SIP.EmailAddresses -contains ('SMTP:' + $User.Alias + '@estrategicaperu.onmicrosoft.com')) -and
       ($Int_SIP.EmailAddresses -notcontains ('SMTP:' + $User.UserPrincipalName))){
        Write-Host 'El usuario: ' $User.UserPrincipalName 'posee la siguiente dirección: ' ($Int_SIP.EmailAddresses | where {$_ -like 'SMTP:*'})
        Write-Host '------------------------------------------------------------------------------------------------------------------------------'

    }
    $i++
}
