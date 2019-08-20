<#$DB_LI_USR = Get-MsolUser -All | Where {$_.isLicensed -eq $true} | sort UserprincipalName
$DB_LI_USR.Count

Get-MsolAccountSku | Where-Object {$_.AccountSkuId -like '*ENTERPRISEPREMIUM*'} | select -ExpandProperty ServiceStatus
#>
#Planner Sway Yammer

#$DB_LI_USR = Get-MsolUser -All | Where {$_.isLicensed -eq $true} | sort UserprincipalName

$DET_E3 = "adpmail:ENTERPRISEPACK"
$DET_E5 = 'adpmail:ENTERPRISEPREMIUM'
$DET_BP = "adpmail:O365_BUSINESS_PREMIUM"
$DET_BE = "adpmail:O365_BUSINESS_ESSENTIALS"
$LIC_E3 = New-MsolLicenseOptions -AccountSkuId $DET_E3 -DisabledPlans "FLOW_O365_P2","POWERAPPS_O365_P2","Deskless","SWAY","YAMMER_ENTERPRISE"
$LIC_E5 = New-MsolLicenseOptions -AccountSkuId $DET_E5 -DisabledPlans "FLOW_O365_P3","POWERAPPS_O365_P3","Deskless","SWAY","YAMMER_ENTERPRISE"
$LIC_BP = New-MsolLicenseOptions -AccountSkuId $DET_BP -DisabledPlans "FLOW_O365_P1","POWERAPPS_O365_P1", "SWAY","YAMMER_ENTERPRISE"
$LIC_BE = New-MsolLicenseOptions -AccountSkuId $DET_BE -DisabledPlans "FLOW_O365_P1","POWERAPPS_O365_P1", "SWAY","YAMMER_ENTERPRISE"
$CON_E3 = 0
$CON_E5 = 0
$CON_BP = 0
$CON_BE = 0

foreach($USR in $DB_LI_USR){
        $LIC = $USR.Licenses[0].AccountSkuId
        $UPN = $USR.UserPrincipalName
        if($LIC -eq $DET_E3){
                #Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LIC_E3
                Set-Mailbox -Identity $UPN -MaxSendSize 25MB -MaxReceiveSize 25MB 
                Write-Host 'Cambio realizado en ' $UPN 'exitoso de licencia E3'
                $CON_E3++    
        }
        elseif($LIC -eq $DET_E5){
                #Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LIC_E5
                Set-Mailbox -Identity $UPN -MaxSendSize 25MB -MaxReceiveSize 25MB
                Write-Host 'Cambio realizado en ' $UPN 'exitoso de licencia E5'
                $CON_E5++    
        }
        elseif($LIC -eq $DET_BP){
                #Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LIC_BP
                Set-Mailbox -Identity $UPN -MaxSendSize 25MB -MaxReceiveSize 25MB
                Write-Host 'Cambio realizado en ' $UPN 'exitoso de licencia BP'
                $CON_BP++    
        }
        elseif($LIC -eq $DET_BE){
                #Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LIC_BE
                Set-Mailbox -Identity $UPN -MaxSendSize 15MB -MaxReceiveSize 25MB
                Write-Host 'Cambio realizado en ' $UPN 'exitoso de licencia BE'
                $CON_BE++    
        }
}
Write-Host '----------------------------------------------------------------------------'
Write-Host 'Se realizadon ' $CON_E3 ' cambios de licencia en productos E3'
Write-Host '----------------------------------------------------------------------------'
Write-Host 'Se realizadon ' $CON_E5 ' cambios de licencia en productos E5'
Write-Host '----------------------------------------------------------------------------'
Write-Host 'Se realizadon ' $CON_BP ' cambios de licencia en productos Business Premium'
Write-Host '----------------------------------------------------------------------------'
Write-Host 'Se realizadon ' $CON_BE ' cambios de licencia en productos Business Esentials'
Write-Host '----------------------------------------------------------------------------'