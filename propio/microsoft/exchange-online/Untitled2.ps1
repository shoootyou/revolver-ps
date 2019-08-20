#$DB_USR = Get-MSOLUser -All 
#$DB_MBX = Get-Mailbox -ResultSize 100000

#$DB_USR_SOR = $DB_USR | Sort-Object UserPrincipalName
#$DB_USR_SOR = $DB_USR | Sort-Object UserPrincipalName

foreach($USR in $DB_USR_SOR){
    If(
        ($USR.UserType -eq  'Member') -and 
        ($USR.UserPrincipalName -like  '*@spsa.pe') -and
        ($USR.ImmutableId -eq $null)
    ){
        
        Write-Host $USR.UserPrincipalName


    }

}