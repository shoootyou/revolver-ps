<#
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

#>

Get-MigrationBatch | Where {$_.Identity -like '*rmateo*'} | Select Identity

#Remove-MigrationBatch -Identity $IDEN.Identity

Get-MigrationUser | Where {$_.identity -like '*hernan*'}


#Get-MsolUser -UserPrincipalName riesgoseguro@estrategicaperu.onmicrosoft.com | Remove-MsolUser -Force
#Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -UserPrincipalName riesgoseguro@estrategicaperu.onmicrosoft.com -RemoveFromRecycleBin -Force 

Connect-MsolService -Credential $UserCredential

#Get-MsolUser -MaxResults 10000 | where {($_.userprincipalName -like '*torioux*') }-or ($_.userprincipalName -like '*rhernandez*') } | Remove-MsolUser -Force

#Get-MsolUser -ReturnDeletedUsers | where {($_.userprincipalName -like '*torioux*') } | Remove-MsolUser -RemoveFromRecycleBin -Force

Get-User -ResultSize 100000 | Where-Object {$_.Name -like '*torioux*'}

