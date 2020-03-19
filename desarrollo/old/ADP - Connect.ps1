$USER = 'torioux@adp.com.pe'
$PASS = ConvertTo-SecureString –String 'Aqzsd159_' –AsPlainText -Force

$CRED = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $USER, $PASS
$SESS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CRED -Authentication Basic -AllowRedirection

Import-PSSession $SESS
Connect-MsolService -Credential $CRED