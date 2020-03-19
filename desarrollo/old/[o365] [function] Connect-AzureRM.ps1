function Connect-AzureRM{
    #---------------------------------------------------------------------------------------------------------------------------------
    #                                                      Values to Connect
    #---------------------------------------------------------------------------------------------------------------------------------
    $GBL_Username = "rodolfo.castelo@tecnofor.pe"
    $GBL_Password = ConvertTo-SecureString –String 'Feelsync240@' –AsPlainText -Force
    $GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password

    Login-AzureRmAccount -Credential $GBL_Credential
}

Connect-AzureRM