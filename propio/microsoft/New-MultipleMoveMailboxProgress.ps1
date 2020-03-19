$MV_LMT = 25
do{
    $MV_SUS = Get-MoveRequest -MoveStatus Suspended
    if($MV_SUS){
        $MV_PRO = Get-MoveRequest -MoveStatus InProgress
        if($MV_PRO){
            If($MV_PRO.Length -lt ($MV_LMT+1)){
                $TMP = $MV_SUS | Select -First ($MV_LMT - $MV_PRO.Length)
                if($TMP){
                    $TMP | % { Resume-MoveRequest -Identity $_.Identity }
                }
            }
        }
        else{
            $MV_SUS | Select -First $MV_LMT | % { Resume-MoveRequest -Identity $_.Identity }
        }
    }
    else{
        break
    }
    Clear-Host
    Write-Host "============================================================================================================================" -ForegroundColor Gray
    Write-Host "                                             " (Get-Date) -ForegroundColor Gray
    Write-Host "============================================================================================================================" -ForegroundColor Gray
    Get-MoveRequest -MoveStatus InProgress | Get-MoveRequestStatistics
    Start-Sleep -Seconds 360
    Clear-Host
}
until(!$MV_SUS)