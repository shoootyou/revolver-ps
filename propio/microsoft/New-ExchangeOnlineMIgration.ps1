Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
$DB_LOC = "D:\Scripts\MoveRequest\"
$DB_DAT = "20191030"
$DB_RML = "marathonperu.mail.onmicrosoft.com"
$DB_USR = Import-Csv ($DB_LOC + $DB_DAT + '.csv')
Remove-Item ($DB_LOC + $DB_DAT + '-Info.csv') -Force -ErrorAction SilentlyContinue
foreach($USR in $DB_USR){
    $USR_MBX = Get-Mailbox $USR.UserPrincipalName
    $USR_MBX  | Get-MailboxStatistics | Select DisplayName,ItemCount,DeletedItemCount,TotalItemSize,TotalDeletedItemSize | Export-Csv ($DB_LOC + $DB_DAT + '-Stats.csv') -Append
    $USR_RML = $USR_MBX.Alias + '@' + $DB_RML
    If([string]$USR_MBX.EmailAddresses -notlike ('*' + $DB_RML + '*')){
        Set-Mailbox $USR.UserPrincipalName -EmailAddresses @{Add="$USR_RML"}
    }
}
Remove-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

Invoke-Command -ComputerName PESQADCONNECT -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
Start-Sleep 300

Set-ExecutionPolicy RemoteSigned -Force
$password = ConvertTo-SecureString “@~r.&_#G6rZurq” -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential (“svc_adconnect@marathonperu.onmicrosoft.com”, $password)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

$OnPrempassword = ConvertTo-SecureString “+CBm=AgXhUcHZX” -AsPlainText -Force
$OnPremUserCredential = New-Object System.Management.Automation.PSCredential (“ASEYCO\svc_excmigration”, $password)
New-MoveRequest -Identity $USR.UserPrincipalName -Remote -RemoteHostName nstores.pe -TargetDeliveryDomain $DB_RML -RemoteCredential $OnPremUserCredential -BadItemLimit 20 -BatchName $USR.UserPrincipalName

Get-MigrationBatch -Identity ($USR.UserPrincipalName.Replace("@","_"))
New-MigrationBatch -Name ($USR.UserPrincipalName.Replace("@","_")) -Users $USR.UserPrincipalName -DisableOnCopy -AutoStart -NotificationEmails rcastelo@canvia.com -ReportInterval 150 -LargeItemLimit 1000 -BadItemLimit 20 -AutoComplete
