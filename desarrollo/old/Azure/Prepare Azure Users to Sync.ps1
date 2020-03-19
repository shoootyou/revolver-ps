##############################################################################################################################
################################                    Variables                    #############################################
##############################################################################################################################
#   Variables details:
#            RG_TMP_YX = Temporal registry. "Y" defines the level and "X" defines the correlative number
#            UN_X = Define the universe to work
#            SWX_Y = "X" defines the number of switch and "Y" defines the kind of use
#            FEC_YX = "Foreach" variable. "Y" defines the level and "X" defines the correlative number
##############################################################################################################################
        Write-Host "What is the domain that will you use?"
        $RG_TMP_11 = Read-Host 
        $RG_TMP_12 = "*" + $RG_TMP_11 + "*"
        $UN_CLOUD = Get-MsolUser | Where-Object {($_.UserPrincipalName -like "$RG_TMP_12")} | Select * | Sort-Object UserPrincipalName
##############################################################################################################################
################################                     Switches                    #############################################
##############################################################################################################################
#--------------------------------------                  Select language               ---------------------------------------
        $SW1_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&English", ` "English"
        $SW1_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&Spanish", ` "Spanish"
        $SW1_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW1_1ST, $SW1_2ND)
        $SW1_ASW = $host.ui.PromptForChoice("Aditional information", "Which language do you use?", $SW1_OPT, 0) 
        switch ($SW1_ASW){
        0{$RG_TMP_13="en-US"}
        1{$RG_TMP_13="es-PE"}
        }
#---------------------------------------------------------------------------------------------------------------------------------
        $SW2_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Prepare users to sync", ` "Prepare users to sync"
        $SW2_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&Verify synced users", ` "Verify synced users"
        $SW2_3RD = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
        $SW2_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW2_1ST, $SW2_2ND,$SW2_3RD)
        $SW2_ASW = $host.ui.PromptForChoice("Aditional information", "Do you like check the information since a specific location?", $SW2_OPT, 0) 
        switch ($SW2_ASW){
            0{
#---------------------------------------------------------------------------------------------------------------------------------
            $SW2_YES = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
            $SW2_NO = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
            $SW2_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW2_YES, $SW2_NO)
            $SW2_ASW = $host.ui.PromptForChoice("Aditional information", "Do you like check the information since a specific location?", $SW2_OPT, 0) 
            switch ($SW2_ASW){
                0{
                    Write-Host "Where will start verification?"
                    Write-Host "Don't forget the format: OU=PrincipalOU,DC=DOM,DC=DOM"
                    $SW1_TMP_11 = Read-Host 
                    $UN_LOCAL = Get-ADUser -Filter * -SearchBase $SW1_TMP_11 | Where-Object {($_.UserPrincipalName -like "$RG_TMP_12")} | Select * | Sort-Object UserPrincipalName
                }
                1{
                    $UN_LOCAL = Get-ADUser -Filter * | Where-Object {($_.UserPrincipalName -like "$RG_TMP_12")} | Select * | Sort-Object UserPrincipalName
                }
            }
##############################################################################################################################
##############################                    Only see information                   #####################################
##############################################################################################################################
                foreach ($FEC_11 in $UN_LOCAL) {
                    $RG_TMP_21 = $FEC_11.UserPrincipalName
                    $RG_TMP_22 = $FEC_11.ObjectGuid
                    $RG_TMP_23 = [System.Convert]::ToBase64String($RG_TMP_22.tobytearray())
                    $RG_TMP_24 = $RG_TMP_23
#---------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------                Search in Cloud Users                --------------------------------------
#---------------------------------------------------------------------------------------------------------------------------------
                    foreach ($FEC_21 in $UN_CLOUD){
                        $RG_TMP_31 = $FEC_21.UserPrincipalName
                        $RG_TMP_32 = $FEC_21.ImmutableId
#---------------------------------------------------------------------------------------------------------------------------------
#----------------------------              Validate equality between local and cloud users             ---------------------------
#---------------------------------------------------------------------------------------------------------------------------------
                        If($RG_TMP_21 -eq $RG_TMP_31){
#---------------------------------------------------------------------------------------------------------------------------------
#------------------------------                   Validate if it's already prepared                  -----------------------------
#---------------------------------------------------------------------------------------------------------------------------------
                            If(!$RG_TMP_32){
                                $RG_TMP_51 = Get-MsolUser -UserPrincipalName $RG_TMP_31 | Select -ExpandProperty ProxyAddresses
#---------------------------------------------------------------------------------------------------------------------------------
#-----------------------                   Exporting proxyAddresses properties to local users                 --------------------
#---------------------------------------------------------------------------------------------------------------------------------
                                foreach ($FEC_31 in $RG_TMP_51) {
                                    $FEC_41 = [String]$FEC_31
                                    Set-ADUser -Identity $FEC_11.SamAccountName -Add @{Proxyaddresses=$FEC_41}
                                }
#---------------------------------------------------------------------------------------------------------------------------------
                            Set-MSOLuser -UserPrincipalName $RG_TMP_31 -ImmutableId $RG_TMP_23
                            }
                        get-aduser -Identity $FEC_11.SamAccountName -Properties preferredLanguage | Set-ADUser -Replace @{preferredLanguage=$RG_TMP_13}
#---------------------------------------------------------------------------------------------------------------------------------
                        }
#---------------------------------------------------------------------------------------------------------------------------------
                    }
#---------------------------------------------------------------------------------------------------------------------------------
                }
             }
            1{
##############################################################################################################################
##############################                    Verifying synced users                 #####################################
##############################################################################################################################
            $OBJ_OUT_2 = @()
                foreach ($FEC_31 in $UN_CLOUD){
                    $OBJ_TMP_2 = New-Object PSObject
                    $RG_TMP_61 = $FEC_31.ImmutableID
                        If(!$RG_TMP_61){
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name "Synced?" -Value "No"
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name Name -Value $FEC_31.DisplayName
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name Email -Value $FEC_31.UserPrincipalName
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name ImmutableID -Value $FEC_31.ImmutableID
                        }
                        else{
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name "Synced?" -Value "Yes"
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name Name -Value $FEC_31.DisplayName
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name Email -Value $FEC_31.UserPrincipalName
                            Add-Member -InputObject $OBJ_TMP_2 -MemberType NoteProperty -Name ImmutableID -Value $FEC_31.ImmutableID
                        }
                        $OBJ_OUT_2 += $OBJ_TMP_2
                }
                $OBJ_OUT_2 | Out-GridView -Title "Information"
            }
            2{

        }

        }