Function Get-MigrationProcess{
    param(
        [Parameter(mandatory)]
            [int]$Seconds,
        [Parameter(mandatory)]
            [int]$Times
    )

$elapsed = [System.Diagnostics.Stopwatch]::StartNew()

for ($t=1; $t -le $Times; $t++) {
 
    $MigrationOut = Get-MigrationUser -Identity lsarayasi@s10peru.com | select SyncedItemCount,Status
    if($MigrationOut.Status -eq 'Syncing'){
        $MigrationMail = 'Se han migrado ' + $MigrationOut.SyncedItemCount + ' hasta ahora'
    }
    else{
        $MigrationMail = 'Se ha terminado la migración'
    }
    ##############################################################################
    $From = "rcastelo.mendez@gmail.com"
    $FromPass = ConvertTo-SecureString –String 'zpdjfmejgydtuxtq' –AsPlainText -Force
    $GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $From, $FromPass
    $To = "rodolfo.castelo@mvpconsulting.pe"
    $Subject = "Email Subject"
    $Body = $MigrationMail
    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"
    Send-MailMessage -From $From -to $To -Subject $Subject `
    -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    -Credential $GBL_Credential # -Attachments $Attachment
    ##############################################################################

    Write-Host "============================================="
    Write-Host "         Se ha ejecutado" $t "veces."
    Write-Host "============================================="
    sleep $Seconds
    }

}

Get-MigrationProcess -Seconds 3600 -Times 17
