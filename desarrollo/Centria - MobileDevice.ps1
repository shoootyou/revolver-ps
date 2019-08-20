$OUT_001  = @()
$DB_MOB = Get-MobileDevice | Select Identity,FriendlyName,DeviceImei,DeviceMobileOperator,DeviceOS,DeviceOSLanguage
foreach($MOB in $DB_MOB){
    $ObjProperties = New-Object PSObject
    $MOB_IDE = $MOB.Identity
    $ARR_POS = $MOB_IDE.IndexOf('\')
    $COM_001 = $MOB_IDE.Substring(0,$ARR_POS)
    
    $TMP_01 = Get-Mailbox -Identity $COM_001 | Select DisplayName,PrimarySmtpAddress
    $TIM_01 = Get-MobileDeviceStatistics -Identity $MOB_IDE

    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DisplayName" -Value $TMP_01.DisplayName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $TMP_01.PrimarySmtpAddress
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "FriendlyName" -Value $MOB.FriendlyName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DeviceImei" -Value $MOB.DeviceImei
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DeviceMobileOperator" -Value $MOB.DeviceMobileOperator
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DeviceOS" -Value $MOB.DeviceOS
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DeviceOSLanguage" -Value $MOB.DeviceOSLanguage
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "FirstSyncTime" -Value $TIM_01.FirstSyncTime
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "LastSuccessSync" -Value $TIM_01.LastSuccessSync

    $OUT_001 += $ObjProperties

    Write-Host 'Función completada exitosamente para' $TMP_01.DisplayName

}

$OUT_001 | Out-GridView

$OUT_001 | Export-Csv $ENV:USERPROFILE\Desktop\mobile2.csv -Encoding UTF8