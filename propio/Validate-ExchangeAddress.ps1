$DB_USR = Import-Csv -path "C:\Temp\Validate-Users.csv"
foreach($USR in $DB_USR){
    $DB_SMTP = (Get-Recipient $USR.UserPrincipalName | Select-Object EmailAddresses).EmailAddresses.SmtpAddress
    $VAL_01 = $DB_SMTP -like '*@correo.mibodega.cl'
    $VAL_02 = $DB_SMTP -like '*@correo.nucleoschile.cl'
    $VAL_03 = $DB_SMTP -like '*@RLC.rlc.cl'
    $VAL_04 = $DB_SMTP -like '*@rentaslacastellana.cl'
    $VAL_05 = $DB_SMTP -like '*@redmegacentrocl.mail.onmicrosoft.com'
    if($VAL_01){
        Set-Mailbox $USR.UserPrincipalName -EmailAddresses @{remove=$VAL_01}
        $VAL_01
    }
    if($VAL_02){
        Set-Mailbox $USR.UserPrincipalName -EmailAddresses @{remove=$VAL_02}
        $VAL_02
    }
    if($VAL_03){
        Set-Mailbox $USR.UserPrincipalName -EmailAddresses @{remove=$VAL_03}
        $VAL_03
    }
    if($VAL_04){
        Set-Mailbox $USR.UserPrincipalName -EmailAddresses @{remove=$VAL_04}
        $VAL_04
    }
    if(!$VAL_05){
        Set-Mailbox $USR.UserPrincipalName -EmailAddresses @{add=$VAL_05}
        $VAL_05
    }
}