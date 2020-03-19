$BCK_DB = Import-CSV $ENV:USERProfile\Desktop\centria.csv
foreach($USR in $BCK_DB){
    Set-MsolUserPassword -userPrincipalName $USR.mail –NewPassword 'Centria2016' -ForceChangePassword $False
    Set-MsolUser -UserPrincipalName $USR.mail -BlockCredential $false
    Write-Host $USR.mail
}
