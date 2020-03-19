$ARC_USR_PT = Get-FilePath
$ARC_USR_DB = Import-CSV -Path $ARC_USR_PT -Delimiter ';'
#$ARC_USR_DB = Get-Mailbox -ResultSize 10000 | Select UserPrincipalName

foreach($LIN in $ARC_USR_DB){
    #[int]$MBX_GB = $LIN.mbxGB
    [string]$MBX_USR = $LIN.UserPrincipalName
    if($MBX_GB -gt 24.99){
        #Set-Mailbox $MBX_USR -RetentionPolicy "Default MRM Policy"
        Write-Host $MBX_USR
    }
}