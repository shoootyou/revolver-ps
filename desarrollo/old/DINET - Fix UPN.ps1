#$GBL_DB = Get-MSOlUSer  | Where {$_.UserPrincipalName -like '*onmicrosoft*'}
#$CSV_DB = Import-CSV 'C:\Users\Rodolfo\OneDrive - TORIOUX GROUP S.A.C\Clientes\Dinet\Comprobacion.csv'
foreach($usuario in $GBL_DB){
    $USR_UPN = $usuario.UserPrincipalName
    $ARR_POS = $USR_UPN.IndexOf('@')
    $USR_LEN = $USR_UPN.Length
    $COM_001 = $USR_UPN.Substring(0,$ARR_POS)

    foreach($USR_CSV in $CSV_DB){
        
        $CSV_UPN = $USR_CSV.UserPrincipalName
        $CSV_POS = $CSV_UPN.IndexOf('@')
        $CSV_LEN = $CSV_UPN.Length
        $COM_002 = $CSV_UPN.Substring(0,$CSV_POS)
        
        if($COM_001 -eq $COM_002){
            $OLD_UPN = $USR_UPN
            $NEW_UPN = $CSV_UPN
            #Set-MsolUserPrincipalName -UserPrincipalName $OLD_UPN -NewUserPrincipalName $NEW_UPN
            #sleep 5
            #Get-MsolUser -UserPrincipalName $NEW_UPN
            #
            Write-Host 'Se modificó la información para el usuario' $NEW_UPN
        }
    }
}