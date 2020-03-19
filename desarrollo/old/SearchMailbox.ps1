$GBL_Username = 'proveedor@americatel.com.pe'
$GBL_Password = ConvertTo-SecureString –String 'Pr0v33d0r_190816' –AsPlainText -Force

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
Import-PSSession $GBL_USR_SSN -Verbose -ErrorAction SilentlyContinue

$GBL_DB = Get-Mailbox -ResultSize Unlimited -Filter {(RecipientTypeDetails -ne "DiscoveryMailbox") -and 
(RecipientTypeDetails -ne "RoomMailbox") -and
(Alias -notlike '*migraci*')} | Select Alias | sort Alias

$GBL_DB | Out-File $ENV:USERPROFILE\Desktop\AllUsers.txt

foreach($RCV_USR in $GBL_DB){
    $MAIL = $RCV_USR.Alias

    $GBL_Username = 'proveedor@americatel.com.pe'
    $GBL_Password = ConvertTo-SecureString –String 'Pr0v33d0r_190816' –AsPlainText -Force

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
    Import-PSSession $GBL_USR_SSN -Verbose -ErrorAction SilentlyContinue

    $USR = $MAIL
    $DOM = '@americatel.com.pe'
    $DIS = "-dsmbx"
    $ADM = 'proveedor@americatel.com.pe'
    $RCV = 'RecoveryReceived'
    $SNT = 'RecoverySent'
    $REC = 'Recovery'

    $NON_EX = $false
    $TRY_01 = $null

    $USR_DOM = $USR + $DOM
    $USR_DIS = $USR + $DIS

    $TRY_01 = Get-Mailbox $USR_DIS -ErrorVariable $null
    if($TRY_01 -eq $null){$NON_EX = $true}

    if($NON_EX){
        '----------------------------------------------------------------------------------------------------------' | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        '|    Working on' + $USR + ' user.' | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        New-Mailbox -Name $USR_DIS -Discovery -WarningAction SilentlyContinue | Out-Null
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host '|    Discovery mailbox for' $USR_DOM 'created sucefully' -ForegroundColor Green
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Green
        Add-MailboxPermission $USR_DIS -User $ADM -AccessRights FullAccess -InheritanceType all -WarningAction SilentlyContinue | Out-Null
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host '|    Permissions for ' $ADM ' added on ' $USR_DOM ' discovery mailbox' -ForegroundColor Green
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
        Write-Host '|    Sleeping 180 segundos' -ForegroundColor Yellow 
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
        sleep 180
        $DAT_START = Get-Date
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '|    Starting proccess at:' $DAT_START -ForegroundColor Cyan
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '|    Searching emails, please wait.' -ForegroundColor Cyan
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        '|    Starting proccess at:' + $DAT_START  | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        '----------------------------------------------------------------------------------------------------------' | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        try{
            Search-Mailbox $USR_DOM -TargetMailbox $USR_DIS -TargetFolder $REC -LogLevel Full -SearchDumpsterOnly -Confirm:$false -WarningAction SilentlyContinue | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
            #Search-Mailbox $USR_DOM -TargetMailbox $USR_DIS -TargetFolder $RCV -LogLevel Full -SearchQuery {Received:01/01/1900..10/08/2016} -SearchDumpsterOnly -Confirm:$false -WarningAction SilentlyContinue
            #Search-Mailbox $USR_DOM -TargetMailbox $USR_DIS -TargetFolder $RCV -LogLevel Full -SearchQuery {sent:01/01/1900..10/08/2016} -SearchDumpsterOnly -Confirm:$false -WarningAction SilentlyContinue    
        }
        catch{$USR_DOM | Out-File $ENV:USERPROFILE\Desktop\DB_Failed.csv -Append}
        $DAT_FINA = Get-Date
        '----------------------------------------------------------------------------------------------------------' | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        '|    Proccess finished: ' + $DAT_FINA  | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        '----------------------------------------------------------------------------------------------------------' | Out-File $ENV:USERPROFILE\Desktop\DB_Completed.txt -Append
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '|    Proccess finished: ' $DAT_FINA -ForegroundColor Cyan
        Write-Host '----------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
}
    else{
        $USR + ', worked' | Out-File $ENV:USERPROFILE\Desktop\DB_Antiguos.txt -Append

    }
}
pause
#Get-Mailbox -Resultsize unlimited -Filter {RecipientTypeDetails -eq "DiscoveryMailbox"}
#remove-Mailbox evargas-dsmbx -Force -Confirm:$false