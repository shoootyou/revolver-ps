$GBL_Username = 'admintorioux@estrategicaperu.onmicrosoft.com'
$GBL_Password = ConvertTo-SecureString –String 'Aqzsd159_' –AsPlainText -Force

$DB_SSN = Get-PSSession
$CN_SSN = 0
foreach($SSN in $DB_SSN){
    if(($SSN.ComputerName -like '*office365*') -and ($SSN.State -eq 'Opened')){
        Remove-PSSession $SSN
    }
    $CN_SSN++
}

$GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password
$GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $GBL_Credential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
Import-PSSession $GBL_USR_SSN 
Connect-MSOLService -Credential $GBL_Credential


Set-MsolUser -UserPrincipalName mlindley@estrategica.com.pe -ImmutableId $null
Get-MsolUser -UserPrincipalName mlindley@estrategica.com.pe | Select ImmutableId

#Restore-MsolUser -UserPrincipalName mlindley@estrategica.com.pe -AutoReconcileProxyConflicts


<#

mlindley ID Correct = c0c52cca-5794-485d-a45d-2e4955671ae7

#>

Get-MSOlUser -ReturnDeletedUsers | where {$_.UserPrincipalName -like '*mlindley*'} | Select ObjectID
#Get-MsolUser -UserPrincipalName mlindley@estrategica.com.pe | Remove-MsolUser

#Remove-MsolUser -RemoveFromRecycleBin -ObjectId b30c9f24-f528-4b2a-a39a-7d362b79e4c3 -Force