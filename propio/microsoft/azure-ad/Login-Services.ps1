$orgName="ausasolucioneslogisticas"
Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Connect-MsolService -Credential $UserCredential
#Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $UserCredential