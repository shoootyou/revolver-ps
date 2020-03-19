cls
$I = 1
$OUT_PAT = "C:\Users\Rodolfo\OneDrive\Work\Sapia\Clientes\AUSA\ReporteUPN.csv"
"Estado, UPN, PrimarySMTPADDR" | Out-File $OUT_PAT
$DB_USR = Get-MsolUser -All | ? { ($_.UserPrincipalName -like '*@ausa.com.pe') -and ($_.isLicensed -eq $true)} | Sort-Object UserPrincipalName
foreach($USR in $DB_USR){
    $UPN = $USR.UserPrincipalName
    Write-Progress -Activity “Cargando reporte de usuarios” -status “Procesando usuario $UPN” -percentComplete ($i / $DB_USR.count*100)
    $VAL_UPN = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue | Select UserPrincipalName
    $VAL_ADD = Get-Mailbox -Identity $UPN  -ErrorAction SilentlyContinue | Select PrimarySmtpAddress
    if($VAL_UPN.UserPrincipalName -ne $VAL_ADD.PrimarySmtpAddress){
        Write-Host "ERROR, " $VAL_UPN.UserPrincipalName ", " $VAL_ADD.PrimarySmtpAddress -ForegroundColor Yellow
        "ERROR, " + $VAL_UPN.UserPrincipalName + ", " + $VAL_ADD.PrimarySmtpAddress| Out-File $OUT_PAT -Append
        #Set-MsolUserPrincipalName -UserPrincipalName $VAL_UPN.UserPrincipalName -NewUserPrincipalName $VAL_ADD.PrimarySmtpAddress
    }
    else{
        Write-Host "CORRECTO, " $VAL_UPN.UserPrincipalName ", " $VAL_ADD.PrimarySmtpAddress -ForegroundColor Green
        "CORRECTO, " + $VAL_UPN.UserPrincipalName + ", " + $VAL_ADD.PrimarySmtpAddress | Out-File $OUT_PAT -Append
    }
    $I++
}