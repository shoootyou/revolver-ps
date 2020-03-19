function Connect-MSCloudServices{
    #---------------------------------------------------------------------------------------------------------------------------------
    #                                                      Values to Connect
    #---------------------------------------------------------------------------------------------------------------------------------
    $GBL_Username = "rodolfo.castelo@tecnofor.pe"
    $GBL_Password = ConvertTo-SecureString –String 'Feelsync240@' –AsPlainText -Force
    $GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password
    #---------------------------------------------------------------------------------------------------------------------------------
    #                                            Connect to Exchange Online and Azure AD
    #---------------------------------------------------------------------------------------------------------------------------------
    $GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $GBL_Credential -Authentication Basic -AllowRedirection
    Import-PSSession $GBL_USR_SSN -Verbose | Out-Null
    Connect-MsolService -Credential $GBL_Credential                                    
    #---------------------------------------------------------------------------------------------------------------------------------
}

Connect-MSCloudServices