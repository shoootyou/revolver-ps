$DB = Get-Mailbox -ResultSize 10000 |where {$_.RecipientTypeDetails -eq 'UserMailbox'} | Select UserPrincipalName 
$i = 1
foreach($User in $DB){
    $INT_USR = $User.UserPrincipalName

    $Int_SIP = Get-Mailbox $INT_USR | Select EmailAddresses

    Write-Progress -Activity “Verificando usuarios” -status “Procesando $INT_USR” -percentComplete ($i / $DB.count*100)

    if($Int_SIP.EmailAddresses -notcontains ('SIP:' +$User.UserPrincipalName)){
        Write-Host 'El usuario: ' $User.UserPrincipalName 'posee la siguiente dirección: ' ($Int_SIP.EmailAddresses | where {$_ -like 'SIP:*'})
        Write-Host '------------------------------------------------------------------------------------------------------------------------------'

    }
    $i++
}