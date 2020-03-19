        $SW1_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
        $SW1_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
        $SW1_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW1_1ST, $SW1_2ND)
        $SW1_ASW = $host.ui.PromptForChoice("Aditional information", "Do you want to remove old synchronizations?", $SW1_OPT, 0) 
        switch ($SW1_ASW){
            0{
            Write-Host "First"
            }
            1{
            Write-Host "Second"   
            }
        }