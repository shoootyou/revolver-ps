<##########################################################################################################################################################################
#                                                                            Global Variables
##########################################################################################################################################################################>

$GBL_SC_PATH =  Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue

function Send-Bulkemail{
    param(
            [bool]$Attach = $false,
            [string]$Subject = ''
    )

if(!$GBL_SC_PATH){
    $GBL_SC_PATH = $env:USERPROFILE+'\Desktop'
}

<##########################################################################################################################################################################
#                                                                          Importing Files
##########################################################################################################################################################################>

$DB_ACC_CSV_PT = Get-FilePath -Title 'Selecciona el archivo con las cuentas de envío' -Path $GBL_SC_PATH
$DB_CLI_CSV_PT = Get-FilePath -Title 'Selecciona el archivo con los correos de los destinatarios' -Path $GBL_SC_PATH
$DB_HTML_PT = Get-FilePath -Title 'Encuentra el correo en el formato HTML' -Filter 'Archivo HTHML (*.html)|*.html' -Path $GBL_SC_PATH
if($Attach){
    $DB_ATTC = Get-FilePath -Title 'Encuentra el PDF a atachar' -Filter 'Archivos PDF (*.pdf)|*.pdf' -Path $GBL_SC_PATH
}

<##########################################################################################################################################################################
#                                                                            Loading Files
##########################################################################################################################################################################>

$DB_HTML = ConvertFrom-HTMLtoMail -Path $DB_HTML_PT
$DB_ACC_CSV  = Import-Csv -Path $DB_ACC_CSV_PT -Delimiter ","
$DB_CLI_CSV  = Import-Csv -Path $DB_CLI_CSV_PT -Delimiter ","

<##########################################################################################################################################################################
#                                                                            Global Variables
##########################################################################################################################################################################>

$GBL_FIN = 0
$GBL_CLI = 0
$GBL_SEND = 25
$GBL_ERROR = 0
if(!$Subject){
    Write-Host "¿Qué asunto tendrá el correo?"
    $SUB_MSG = Read-Host 
}

<##########################################################################################################################################################################
#                                                                    Verifying old files and creating new
##########################################################################################################################################################################>

try{Rename-Item -Path $GBL_SC_PATH\error.csv ErrorOld.csv -ErrorAction SilentlyContinue}
catch{"Time,Sender,Mail" | Out-File $GBL_SC_PATH\error.csv -Append}

try{Rename-Item -Path $GBL_SC_PATH\log.csv  LogOld.csv -ErrorAction SilentlyContinue}
catch{"Time,Sender,Mail" | Out-File $GBL_SC_PATH\log.csv -Append}


<##########################################################################################################################################################################
#                                                                           Loading files
##########################################################################################################################################################################>

$Inline_ATT = @{

    image001 = "$GBL_SC_PATH\image001.png"
    image002 = "$GBL_SC_PATH\image002.png"
    image003 = "$GBL_SC_PATH\image003.png"
    image004 = "$GBL_SC_PATH\image004.png"
    image005 = "$GBL_SC_PATH\image005.png"
    image006 = "$GBL_SC_PATH\image006.png"
    image007 = "$GBL_SC_PATH\image007.png"
    image008 = "$GBL_SC_PATH\image008.png"
    image009 = "$GBL_SC_PATH\image009.png"
    image010 = "$GBL_SC_PATH\image010.png"
    image011 = "$GBL_SC_PATH\image011.png"
   

}

<##########################################################################################################################################################################
#                                                                             Process
##########################################################################################################################################################################>

do{
    $GBL_ACC = 0
    
    foreach($ACC in $DB_ACC_CSV){
        $GBL_BCC = 0
        $GBL_BCC_TX = ""
        $ACC_USR  = $ACC.Account
        $ACC_PASS = ConvertTo-SecureString -String $ACC.Password -AsPlainText -Force
        $ACC_CRE  = New-Object System.Management.Automation.PSCredential $ACC_USR, $ACC_PASS
        
        
        do{
            $TIME = Get-Date
            try{
                if(!$Attach){
                    Send-EmailMessage -Credential $ACC_CRE -From $ACC_USR -To $DB_CLI_CSV[$GBL_ERROR].Mail -Subject $SUB_MSG -Body $DB_HTML -BodyAsHtml -smtpserver outlook.office365.com -usessl -InlineAttachments $Inline_ATT -ErrorAction Continue
                }
                else{
                    Send-EmailMessage -Credential $ACC_CRE -From $ACC_USR -To $DB_CLI_CSV[$GBL_ERROR].Mail -Subject $SUB_MSG -Body $DB_HTML -BodyAsHtml -smtpserver outlook.office365.com -usessl -InlineAttachments $Inline_ATT -ErrorAction Continue -Attachments $DB_ATTC
                }
            }
            catch{
                if($DB_CLI_CSV[$GBL_ERROR].Mail){
                    $TIME.ToString() + "," + $ACC_USR + "," + $DB_CLI_CSV[$GBL_ERROR].Mail | Out-File $GBL_SC_PATH\error.csv -Append
                }
            }
            if($DB_CLI_CSV[$GBL_ERROR].Mail){
                $TIME.ToString() + "," + $ACC_USR + "," + $DB_CLI_CSV[$GBL_ERROR].Mail | Out-File $GBL_SC_PATH\log.csv -Append
                Write-Host $TIME.ToString() "," $ACC_USR "," $DB_CLI_CSV[$GBL_ERROR].Mail
            }

            $GBL_BCC++ # Breaker of the loop for change sender account #
            $GBL_ERROR++ # Global line of DB of clients  #
        }until($GBL_BCC -ge $GBL_SEND)
        
        $GBL_CLI += $GBL_SEND
        $GBL_ACC++ # Global counter for cound all account in Senders' account #

        if($GBL_ACC -eq $DB_ACC_CSV.Count){
            Write-Host "Sleeping 120 seconds"
            "-------------------,--- Sleep 120 seconds ----,-----------------------" | Out-File $GBL_SC_PATH\log.csv -Append
            sleep 120
        }
        if($GBL_CLI -ge $DB_CLI_CSV.Count-1){$GBL_FIN++; break}
    }
}
until($GBL_FIN -eq 1)

if($GBL_FIN -eq 1){
    $CMP_CLI = Get-Date
    $CMP_CLI_OUT = $CMP_CLI.Day.ToString() +"-" + $CMP_CLI.Month.ToString() +"-"  + $CMP_CLI.Year.ToString() +"-" + $CMP_CLI.Hour.ToString() +"-" + $CMP_CLI.Minute.ToString() + "-ClientsCompleted.csv"
    Rename-Item -Path $DB_CLI_CSV_PT $CMP_CLI_OUT 
}

}

Send-Bulkemail -Attach $false

pause