##################################################################################################################################
#---------------------------------------------------------------------------------------------------------------------------------
#                                                      Values to Connect
#---------------------------------------------------------------------------------------------------------------------------------
$GBL_Username = "rodolfo@m365x164202.onmicrosoft.com"
$GBL_Password = ConvertTo-SecureString –String 'Clave123456' –AsPlainText -Force
$GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password
#---------------------------------------------------------------------------------------------------------------------------------
#                                            Connect to Exchange Online and Azure AD
#---------------------------------------------------------------------------------------------------------------------------------
$GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $GBL_Credential -Authentication Basic -AllowRedirection
Import-PSSession $GBL_USR_SSN
Connect-MsolService -Credential $GBL_Credential                                    
#---------------------------------------------------------------------------------------------------------------------------------