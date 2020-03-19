##################################################################################################################################
#######################################                        V 3.5                       #######################################
##################################################################################################################################
#---------------------------------------------------------------------------------------------------------------------------------
#                                                     Functions
#---------------------------------------------------------------------------------------------------------------------------------

function Connect-MSCloudServices{
    #---------------------------------------------------------------------------------------------------------------------------------
    #                                                      Values to Connect
    #---------------------------------------------------------------------------------------------------------------------------------
    $GBL_Username = "az_sync_acc@sistema10985.onmicrosoft.com"
    $GBL_Password = ConvertTo-SecureString –String 'Pat$$CW0N*/-' –AsPlainText -Force
    $GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password
    #---------------------------------------------------------------------------------------------------------------------------------
    #                                            Connect to Exchange Online and Azure AD
    #---------------------------------------------------------------------------------------------------------------------------------
    $GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $GBL_Credential -Authentication Basic -AllowRedirection
    Import-PSSession $GBL_USR_SSN -Verbose | Out-Null
    Connect-MsolService -Credential $GBL_Credential                                    
    #---------------------------------------------------------------------------------------------------------------------------------
}

##################################################################################################################################
#                                                      Connecting
##################################################################################################################################
$SW0_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
$SW0_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
$SW0_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW0_1ST, $SW0_2ND)
$SW0_ASW = $host.ui.PromptForChoice("Aditional information", "Do you need sign in to the Cloud Services?", $SW0_OPT, 0) 
switch ($SW0_ASW){
    0{
        Connect-MSCloudServices
    }
    1{
    }
}
##################################################################################################################################
#                                                      Clean Up Old migrations
##################################################################################################################################
$SW1_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
$SW1_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
$SW1_3RD = New-Object System.Management.Automation.Host.ChoiceDescription "&Error", ` "Error"
$SW1_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW1_1ST, $SW1_2ND,$SW1_3RD)
$SW1_ASW = $host.ui.PromptForChoice("Aditional information", "Do you want to remove old synchronizations?", $SW1_OPT, 0) 
switch ($SW1_ASW){
    0{

        #---------------------------------------------------------------------------------------------------------------------------------
        #                                                     Removing Old Migrations
        #---------------------------------------------------------------------------------------------------------------------------------
        $PS_MG_OLD = Get-MigrationBatch | Select * 
        foreach($IN_MG_OLD in $PS_MG_OLD){
            if($IN_MG_OLD.Status -like '*Synced*'){
                $IN2_MG_ID = $IN_MG_OLD.Identity.Name
                Remove-MigrationBatch -Identity $IN2_MG_ID -Confirm:$false
            }
        }
        $PS_MG_USR_OLD = Get-MigrationUser | Select *  
        foreach($IN_MG_USR_OLD in $PS_MG_USR_OLD){
            if($IN_MG_USR_OLD.Status -like '*Synced*'){
                $IN2_USR_ID = $IN_MG_USR_OLD.Identity
                Remove-MigrationUser -Identity $IN2_USR_ID -Confirm:$false
            }
        }

    }
    1{   
    }
    2{

        #---------------------------------------------------------------------------------------------------------------------------------
        #                                                  Removing  Migrations with Errors
        #---------------------------------------------------------------------------------------------------------------------------------
        $PS_MG_OLD = Get-MigrationBatch | Select * 
        foreach($IN_MG_OLD in $PS_MG_OLD){
            if($IN_MG_OLD.Status -like '*Errors*'){
                $IN2_MG_ID = $IN_MG_OLD.Identity.Name
                Remove-MigrationBatch -Identity $IN2_MG_ID -Force -Confirm:$false
            }
        }
        $PS_MG_USR_OLD = Get-MigrationUser | Select * 
        foreach($IN_MG_USR_OLD in $PS_MG_USR_OLD){
            if($IN_MG_USR_OLD.Status -like '*Failed*'){
                $IN2_USR_ID = $IN_MG_USR_OLD.Identity
                Remove-MigrationUser -Identity $IN2_USR_ID -Force -Confirm:$false
            }
        }

    }
}
##################################################################################################################################
#                                                      Start new migrations
##################################################################################################################################
$SW2_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
$SW2_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
$SW2_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW2_1ST, $SW2_2ND)
$SW2_ASW = $host.ui.PromptForChoice("Aditional information", "Do you want to start the synchronizations?", $SW2_OPT, 0) 
switch ($SW2_ASW){
    0{
        #---------------------------------------------------------------------------------------------------------------------------------
        #                                                      Preparing Migration
        #---------------------------------------------------------------------------------------------------------------------------------
        $PS_MG_END = Get-MigrationEndPoint | Where-Object {$_.EndpointType -eq 'IMAP'}
        $PS_MG_END_ID = $PS_MG_END.Identity
        $PS_CSV_FL = Get-ChildItem -Path 'C:\Users\Rodolfo\SharePoint\MVP Consulting - Clientes\Sistemas10\Ingenieria\Migraciones\Cuarto\*' –Include *.csv
        $i = 0
        foreach($IN_CSV_FL in $PS_CSV_FL){
            $i++
            if($i -le 20){
            $PS_MG_CSV_PATH = $IN_CSV_FL.FullName
            $PS_MG_CSV_DATA = ([System.IO.File]::ReadAllBytes($PS_MG_CSV_PATH))
            $PS_MG_CSV_IMP  = Import-Csv $PS_MG_CSV_PATH -Delimiter ','

                #---------------------------------------------------------------------------------------------------------------------------------
                #                                                      Start Migration
                #---------------------------------------------------------------------------------------------------------------------------------

                foreach($IN_USER in $PS_MG_CSV_IMP){
                    New-MigrationBatch -Name $IN_USER.EmailAddress -CSVData $PS_MG_CSV_DATA -SourceEndpoint $PS_MG_END_ID -NotificationEmails 'rcastelo@mvpconsulting.pe'
                    Start-MigrationBatch -Identity $IN_USER.EmailAddress
                }

            }
            if($i -eq 20){
                Write-Host $IN_USER.EmailAddress
                Write-Host $i "--------------------------"
            }
        }
    }
    1{
        Write-Host "See you later"
    }

}
