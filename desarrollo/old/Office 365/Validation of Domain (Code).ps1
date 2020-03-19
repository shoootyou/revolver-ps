##################################################################################################################################
#######################################                 Connecting Services                 ######################################
#---------------------------------------------------------------------------------------------------------------------------------
#                    If you use any method to verify the domain, please, change that variables by $GBL_TMP_11
#                                       Remove the comment to services do you need
#---------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------                 Global Credentials                  --------------------------------------

$GBL_USR_CDR = Get-Credential

#--------------------------------------                   Exchange Online                   --------------------------------------

$GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $GBL_USR_CDR -Authentication Basic -AllowRedirection
Import-PSSession $GBL_USR_SSN

#--------------------------------------                    MSOL Services                    --------------------------------------

Connect-MsolService -Credential $GBL_USR_CDR

Add-AzureAccount -Credential $GBL_USR_CDR

##############################################################################################################################
Write-Host "What is the domain that will you use?"
$GBL_TMP_01 = Read-Host 
Write-Host
Write-Host "=========================================================================================================="
Write-Host "                              Obtaining and verifying information, please wait"
Write-Host "=========================================================================================================="
Write-Host
$GBL_TMP_10 = 0
$GBL_TMP_11 = "*" + $GBL_TMP_01 + "*"
$GBL_TMP_12 = Get-AcceptedDomain | Select DomainName
foreach ($GBL_FEC_12 in $GBL_TMP_12) {
        if($GBL_FEC_12 -like $GBL_TMP_11){
            if($GBL_FEC_12 -notlike "*microsoft.com*"){
                $GBL_TMP_10++
                IF($GBL_TMP_10 -eq 1){
#--------------------------------------                    Your code STARTS here                    --------------------------------------






           
                
#--------------------------------------                    Your code ENDS here                    --------------------------------------
            pause
##############################################################################################################################
                Break
                }
            }
        }
}
        if($GBL_TMP_10 -ne 1){
                Write-Host "You can't modify that domain, sorry."

##############################################################################################################################
            Write-Host
            pause
##############################################################################################################################
        }
