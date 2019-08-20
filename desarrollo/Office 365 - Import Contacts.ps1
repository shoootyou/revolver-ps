$ADM_USR = 'torioux@exsanet.onmicrosoft.com'
$ADM_PAS = 'Tg159357'
$DES_USR = 'mprueba@exsa.net'
$DB = Get-ChildItem -Path 'C:\Users\Rodolfo\OneDrive - TORIOUX GROUP S.A.C\Clientes\EXSA\Export kmaslo' | Where {$_.Name -like 'Export*.csv'}
Get-Date
foreach($CSV in $DB){
    .\Import-MailboxContacts.ps1 -Username $ADM_USR -Password $ADM_PAS -CSVFileName $CSV.FullName -Impersonate:$true -EmailAddress $DES_USR -EwsUrl 'https://outlook.office365.com/EWS/Exchange.asmx'
}
Get-Date