$ErrorActionPreference = "SilentlyContinue"

$CN_GPR = 1
$DB_USR = Get-Mailbox -ResultSize Unlimited | Select *
$OB_OUT = @()

foreach($USR in $DB_USR){
    Write-Progress -Activity “Revisando Información de usuarios" -status “Revisando el usuario $PR_BAR” -percentComplete ($CN_GPR / $DB_USR.count*100) -Id 500
    $USR_STA = Get-MailboxStatistics -Identity $USR.Identity | Select *
    $OB_TMP = New-Object PsObject
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "DisplayNamer" -Value $USR.DisplayName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Alias" -Value $USR.Alias
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "SamAccountName" -Value $USR.SamAccountName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "RecipientType" -Value $USR.RecipientType
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "OrganizationalUnit" -Value $USR.OrganizationalUnit
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $USR.PrimarySmtpAddress
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "Database" -Value $USR.Database
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ServerNamer" -Value $USR.ServerName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveDatabase" -Value $USR.ArchiveDatabase
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveGuid" -Value $USR.ArchiveGuid
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveName" -Value $USR.ArchiveName
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveQuota" -Value $USR.ArchiveQuota
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveWarningQuota" -Value $USR.ArchiveWarningQuota
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveDomain" -Value $USR.ArchiveDomain
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ArchiveStatus" -Value $USR.ArchiveStatus
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $USR.UseDatabaseQuotaDefaults
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ProhibitSendReceiveQuota" -Value $USR.ProhibitSendReceiveQuota
    $USR_SIZ = ([math]::Round(($USR_STA.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2))
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "TotalItemSize" -Value $USR_SIZ

    If($USR_SIZ -gt 125){
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "MailboxClass" -Value "CLASS A"
    }
    elseIf($USR_SIZ -gt 500){
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "MailboxClass" -Value "CLASS B"
    }
    elseIf($USR_SIZ -gt 1024){
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "MailboxClass" -Value "CLASS C"
    }
    elseIf($USR_SIZ -gt 3072){
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "MailboxClass" -Value "CLASS D"
    }
    else{
        Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "MailboxClass" -Value "CLASS E"
    }
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "ItemCount" -Value $USR_STA.ItemCount
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "DeletedItemCount" -Value $USR_STA.DeletedItemCount
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "TotalDeletedItemSize" -Value ([math]::Round(($USR_STA.TotalDeletedItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2))
    Add-Member -InputObject $OB_TMP -MemberType NoteProperty -Name "LastLogonTime" -Value $USR_STA.LastLogonTime

    $OB_OUT += $OB_TMP
    $CN_GPR++

}

$OB_OUT | Out-GridView -Title "Relación de Usuarios"