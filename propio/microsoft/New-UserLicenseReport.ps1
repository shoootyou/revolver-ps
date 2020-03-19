#$DB_USR = Get-AzureADUser -All $true | Select ObjectId,DisplayName,AssignedLicenses,UserPrincipalName
#$DB_LIC = Get-AzureADSubscribedSku | Select SkuPartNumber,SkuID
$GR_OUT = @()
$I = 1
foreach ($USR in $DB_USR){
    Write-Progress -Id 1 -Activity “Obteniendo información” -status (“Trabajando en " + $USR.UserPrincipalName) -percentComplete ($i / $DB_USR.count*100)
    $DB_ULI = $null
    $DB_ULI = $USR.AssignedLicenses
    $GR_TMP = New-Object PsObject
    Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name ObjectId -Value $USR.ObjectId
    Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name DisplayName -Value $USR.DisplayName
    Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name UserPrincipalName -Value $USR.UserPrincipalName
    Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name Domain -Value (($USR.UserPrincipalName).Substring(($USR.UserPrincipalName).IndexOf("@")+1)).ToLower()
    if(!$DB_ULI){
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name PrimaryLicense -Value ""
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name OtherLicenses -Value ""
    }
    else{
        $PRO_LIC = ""
        $OTH_LIC = ""
        $DB_ULI = $DB_ULI | Sort-Object SkuId
        $DB_LIC = $DB_LIC | Sort-Object SkuPartNumber
        foreach($LIC in $DB_LIC){
            foreach($ULI in $DB_ULI){
                if($ULI.SkuId -eq $LIC.SkuId){
                    $PRO_PWB = $null
                    if($LIC.SkuPartNumber -eq 'DESKLESSPACK'){
                        $PRO_LIC += "OFFICE 365 F1,"
                    }
                    elseif($LIC.SkuPartNumber -eq 'ENTERPRISEPACK'){
                        $PRO_LIC += "OFFICE 365 ENTERPRISE E3,"
                    }
                    elseif($LIC.SkuPartNumber -eq 'EXCHANGESTANDARD'){
                        $PRO_LIC += "EXCHANGE ONLINE (PLAN 1),"
                    }
                    elseif($LIC.SkuPartNumber -eq 'OFFICESUBSCRIPTION'){
                        $PRO_LIC += "OFFICE 365 PROPLUS,"
                    }
                    elseif($LIC.SkuPartNumber -eq 'STANDARDPACK'){
                        $PRO_LIC += "OFFICE 365 ENTERPRISE E1,"
                    }
                    else{
                        $OTH_LIC += $LIC.SkuPartNumber + ","
                    }
                }
            }
        }
        if($PRO_LIC){
            $PRO_LIC = $PRO_LIC.Substring(0,$PRO_LIC.Length -1)
        }
        if($OTH_LIC){
            $OTH_LIC = $OTH_LIC.Substring(0,$OTH_LIC.Length -1)
        }
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name PrimaryLicense -Value $PRO_LIC
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name OtherLicenses -Value $OTH_LIC
    }
    $GR_OUT += $GR_TMP
    $I++
}
$GR_OUT | Out-GridView -Title "Reporte de licencias de Office 365"
$GR_OUT | Export-Csv C:\Scripts\Laive-ReporteLicencia.csv