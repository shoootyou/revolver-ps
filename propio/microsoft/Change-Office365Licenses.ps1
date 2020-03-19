$I = 0
$DB_USR = Import-Csv -Path "C:\Users\operador\Downloads\Scripts\UsersLicenses.csv" -Encoding UTF8
foreach($USR in $DB_USR){

    $userUPN = $USR.UserPrincipalName
	Write-Progress -activity “Ejecutando Script” -status "Removiendo licencia: $userUPN” -PercentComplete (($i / $DB_USR.count)*100)
    Set-MsolUserLicense -UserPrincipalName $userUPN -RemoveLicenses "comercioperu:ENTERPRISEPACK"
    Sleep 10
	Write-Progress -activity “Ejecutando Script” -status "Adicionando licencia: $userUPN”
    Set-MsolUserLicense -UserPrincipalName $userUPN -AddLicenses "comercioperu:STANDARDPACK"
    Set-MsolUserLicense -UserPrincipalName $userUPN -AddLicenses "comercioperu:OFFICESUBSCRIPTION"
    sleep 10
$i++
}