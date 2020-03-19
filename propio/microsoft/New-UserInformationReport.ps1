$DB_USR = Get-ADUser -Filter * -ResultSetSize $null  -Properties Department,Title,DisplayName,Mail,Enabled,Manager,ipPhone,LastLogonDate,Mobile,Name,facsimileTelephoneNumber,PhysicalDeliveryOfficeName,telephoneNumber,PostalCode,sAMAccountName | Select *
$CN_GPR = 1
$TM_LCL = (Get-Date).ToShortDateString().Replace("/","-")
$OT_PAT = "C:\Script\ReportesUsuario\Reporte-" + $TM_LCL + ".csv"
$OB_OUT = @()
foreach($USR in $DB_USR){
    $OB_TMP = New-Object PsObject
    $PR_BAR = $USR.sAMAccountName
    Write-Progress -Activity “Revisando Información de usuarios" -status “Revisando el usuario $PR_BAR” -percentComplete ($CN_GPR / $DB_USR.count*100) -Id 500
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Nombre para mostrar" -Value $USR.DisplayName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Usuario de Red" -Value $USR.sAMAccountName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Nombre" -Value $USR.Name
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name Departamento -Value $USR.Department
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name Cargo -Value $USR.Title
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name Correo -Value $USR.Mail
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name Habilitado -Value $USR.Enabled
    if($USR.Manager){
        $US_MGR = Get-ADUser $USR.Manager -ErrorAction SilentlyContinue | Select sAMAccountName
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Jefe inmediato (Manager)" -Value $US_MGR.sAMAccountName
    }
    else{
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Jefe inmediato (Manager)" -Value ""
    }
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Jefe inmediato (Fax)" -Value $USR.facsimileTelephoneNumber
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name ipPhone -Value $USR.ipPhone
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Última fecha de inicio de sesión" -Value $USR.LastLogonDate
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Móvil" -Value $USR.Mobile
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name Office -Value $USR.PhysicalDeliveryOfficeName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name Anexo -Value $USR.telephoneNumber
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Código Postal" -Value $USR.PostalCode
    
    $OB_OUT += $OB_TMP
    $CN_GPR++
}
$OB_OUT | Out-GridView -Title "Relación de Usuarios"
$OB_OUT | Export-Csv  $OT_PAT -Encoding UTF8