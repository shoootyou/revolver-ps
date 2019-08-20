Write-Host ''
Write-Host '--------------------------------------------------------------------' -ForegroundColor Green
Write-Host '                 Obteniendo data, por favor espere.                 ' 
Write-Host '--------------------------------------------------------------------' -ForegroundColor Green

Import-Module ActiveDirectory
$OUT_INF = @()
$AD_DB = Get-ADUser -Filter * 

foreach($OUT_USR in $AD_DB){
    
    $TMP_OBJ = New-Object PSObject
               
    $EXP_OBJ = Get-ADUser -identity $OUT_USR.SamAccountName -Properties SamAccountName,Description,Office,TelephoneNumber,wWWHomePage,City,State,PostalCode,Country,Title,Department,Company,msDS-cloudExtensionAttribute1,ProxyAddresses,employeeID,dNI,dntCeCo,dntDesCeco,dntDesOrdenInterna,dntFechaIngreso,dntOrdenInterna,dntPuesto

    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "SamAccountName" -Value $EXP_OBJ.SamAccountName
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "UserPrincipalName" -Value $EXP_OBJ.UserPrincipalName
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "Description" -Value $EXP_OBJ.Description
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "Office" -Value $EXP_OBJ.Office
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "TelephoneNumber" -Value $EXP_OBJ.TelephoneNumber
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "WebPage" -Value $EXP_OBJ.wWWHomePage
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "City" -Value $EXP_OBJ.City
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "State" -Value $EXP_OBJ.State
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "PostalCode" -Value $EXP_OBJ.PostalCode
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "Country" -Value $EXP_OBJ.Country
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "JobTitle" -Value $EXP_OBJ.JobTitle.Value
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "Department" -Value $EXP_OBJ.Department
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "Company" -Value $EXP_OBJ.Company
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "dNI" -Value $EXP_OBJ.dNI
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "dntCeCo" -Value $EXP_OBJ.dntCeCo
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "dntDesCeco" -Value $EXP_OBJ.dntDesCeco
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "dntFechaIngreso" -Value $EXP_OBJ.dntFechaIngreso
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "dntOrdenInterna" -Value $EXP_OBJ.dntOrdenInterna
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "dntPuesto" -Value $EXP_OBJ.dntPuesto
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "employeeID" -Value $EXP_OBJ.employeeID
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "msDScloudExtensionAttribute1" -Value $EXP_OBJ.'msDS-cloudExtensionAttribute1'
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses1" -Value $EXP_OBJ.ProxyAddresses[0]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses2" -Value $EXP_OBJ.ProxyAddresses[1]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses3" -Value $EXP_OBJ.ProxyAddresses[2]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses4" -Value $EXP_OBJ.ProxyAddresses[3]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses5" -Value $EXP_OBJ.ProxyAddresses[4]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses6" -Value $EXP_OBJ.ProxyAddresses[5]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses7" -Value $EXP_OBJ.ProxyAddresses[6]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses8" -Value $EXP_OBJ.ProxyAddresses[7]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses9" -Value $EXP_OBJ.ProxyAddresses[8]
    Add-Member -InputObject $TMP_OBJ -MemberType NoteProperty -Name "ProxyAddresses10" -Value $EXP_OBJ.ProxyAddresses[9]

    
    $OUT_INF += $TMP_OBJ
}

$OUT_INF | Out-GridView -Title "Mailbox and Archive Sizes"
$OUT_INF | Export-Csv -Path $ENV:USERPROFILE\Desktop\Reporte.csv