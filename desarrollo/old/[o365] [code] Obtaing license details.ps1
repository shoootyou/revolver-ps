$GRID_OUT = @()
$DB_ACCO = Get-MSOLUser | Select DisplayName,UserPrincipalName,Licenses 

foreach($INT_ACC in $DB_ACCO){
        foreach($INT_LIC in $INT_ACC.Licenses){
            $GRID_TMP = New-Object PsObject 

            Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name DisplayName -Value $INT_ACC.DisplayName
            Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name UserPrincipalName -Value $INT_ACC.UserPrincipalName

            if(!$INT_LIC.AccountSKUid){
                Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name Licencia -Value "Sin Licencia Asignada"
            }
            ElseIf($INT_LIC.AccountSKUid -eq "TecnologiayCreatividad:ENTERPRISEPACK"){
                Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name Licencia -Value "Office 365 Enterprise E3"
            }
            Elseif($INT_LIC.AccountSKUid -eq "TecnologiayCreatividad:EMS"){
                Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name Licencia -Value "Enterprise Mobility Suite"
            }
            Elseif($INT_LIC.AccountSKUid -eq "TecnologiayCreatividad:CRMIUR"){
                Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name Licencia -Value "Microsoft Dynamics CRM"
            }
            Elseif($INT_LIC.AccountSKUid -eq "TecnologiayCreatividad:POWER_BI_PRO"){
                Add-Member -InputObject $GRID_TMP -MemberType NoteProperty -Name Licencia -Value "Power BI Pro"
            }
            $GRID_OUT += $GRID_TMP
        }
}

$GRID_OUT | Out-GridView -Title "Información Personal"
$GRID_OUT | Export-Csv $env:USERPROFILE\desktop\Licencias.csv