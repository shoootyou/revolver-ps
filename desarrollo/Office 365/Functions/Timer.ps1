Function Get-MigrationProcess{
    param(
        [Parameter(mandatory)]
            [int]$Seconds,
        [Parameter(mandatory)]
            [int]$Times
    )

$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "=================================================================================="
    Write-Host "        Su tiempo de intervalo es de" $Seconds "segundos, entre ejecución."
    Write-Host "----------------------------------------------------------------------------------"
    Write-Host "                     El bucle tiene" $Times "repeticiones."
    Write-Host "=================================================================================="
for ($t=1; $t -le $Times; $t++) {
 
    Get-MigrationBatch | ft -AutoSize
    Get-MigrationUser  | ft -AutoSize

    Write-Host "----------------------------------------------------------------------------------"
    sleep $Seconds
    }

}

Get-MigrationProcess -Seconds 1 -Times 1
