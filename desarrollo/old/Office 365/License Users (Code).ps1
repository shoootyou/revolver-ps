$OBJ_OUT = @()
$GBL_License = Get-MsolUser| Where-Object {($_.UserPrincipalName -notlike '*onmicrosoft*')} | Sort-Object isLicensed | Select *
foreach($LIC_USR in $GBL_License){
        $OBJ_OUT_PRO = New-Object PSObject

             

        Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Nombre' -Value $LIC_USR.UserPrincipalName
        Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Licencia' -Value $LIC_USR.Licenses.AccountSkuId
            
        $OBJ_OUT += $OBJ_OUT_PRO
}



$OBJ_OUT | Out-GridView -Title "Información Personal"
$OBJ_OUT | Export-Csv C:\Users\Rodolfo\Desktop\Exportación.csv 
