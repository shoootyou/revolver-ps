$GBL_Username = 'admin365cloud@dinetcorp.onmicrosoft.com'
$GBL_Password = ConvertTo-SecureString –String 'T2UhaT_&wA7e' –AsPlainText -Force

$DB_SSN = Get-PSSession
$CN_SSN = 0
foreach($SSN in $DB_SSN){
    if(($SSN.ComputerName -like '*office365*') -and ($SSN.State -eq 'Opened')){
        Remove-PSSession $SSN
    }
    $CN_SSN++
}

$GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password
$GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $GBL_Credential -Authentication Basic –AllowRedirection -ErrorAction SilentlyContinue
Import-PSSession $GBL_USR_SSN -Verbose -ErrorAction SilentlyContinue
Connect-MsolService -Credential $GBL_Credential