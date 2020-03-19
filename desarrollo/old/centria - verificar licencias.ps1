$DB = Get-MsolUser -MaxResults 10000 | Where {$_.UserprincipalName -like '*@centria.net'} | Select UserPrincipalName
Write-Host 'DB obtained'
$LicenseExpo = @()
$i = 1
foreach($User in $DB){
    $LIC_DET = Get-MsolUser -UserPrincipalName $User.UserPrincipalName | Select DisplayName,UserPrincipalName -ExpandProperty Licenses
    $LIC_DET_UPN = $LIC_DET.UserPrincipalName
    $ObjProperties = New-Object PSObject
               
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Display Name" -Value $LIC_DET[0].DisplayName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "User Principal Name" -Value $LIC_DET[0].UserPrincipalName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "License" -Value $LIC_DET[0].AccountSkuId
    $LicenseExpo += $ObjProperties
    Write-Progress -Activity “Gathering Information” -status “Working on $LIC_DET_UPN” -percentComplete ($i / $DB.count*100)
    $i++

}

$LicenseExpo | Out-GridView -Title "License Details"