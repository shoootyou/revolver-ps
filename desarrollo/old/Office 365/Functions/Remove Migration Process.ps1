function Remove-MigrationProcess{
    param(
        [Parameter(mandatory)]
        [string]$UserPrincipalName
    )
    Write-Host "=================================================================================="
    Write-Host "        Se iniciará la remoción del usuario " $UserPrincipalName
    Write-Host "=================================================================================="
 
        Remove-MigrationBatch -Identity $UserPrincipalName  -Confirm:$false
        Remove-MigrationUser -Identity $UserPrincipalName -Confirm:$false


}

Remove-MigrationProcess -UserPrincipalName lsarayasi@s10peru.com