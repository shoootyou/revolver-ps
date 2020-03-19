do{
    
    Write-Host '------------------------------------------------------------------------' -ForegroundColor Yellow
    Write-Host "Checking....." -ForegroundColor Yellow
    

    $COUN_INPR = 0
    $COUN_COMP = 0
    $MAX_MOVE  = 20
    $SLEEP_TIME = 60

    $DB_MOVE = Get-MoveRequest

    foreach($MOVE in $DB_MOVE){
    if(($MOVE.Status -ne 'Completed') -and ($MOVE.Status -ne 'Suspended')){
        $COUN_INPR++;
    }
    elseif($MOVE.Status -eq 'Completed'){
        Remove-MoveRequest -Identity $MOVE.Identity -Confirm:$false
        Write-Host '------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host 'We removed the process for ' $Move.Alias ". It's all completed." -ForegroundColor Green
    }
    }

    foreach($MOVE in $DB_MOVE){
    if($MOVE.Status -eq 'Suspended'){
        if($COUN_INPR -lt $MAX_MOVE){
            Resume-MoveRequest -Identity $MOVE.Identity -Confirm:$false
            Write-Host '------------------------------------------------------------------------' -ForegroundColor Green
            Write-Host 'We started the process for ' $Move.Alias ".Run run run!" -ForegroundColor Green
            $COUN_INPR++
        }

    }
    }
    Write-Host '------------------------------------------------------------------------' -ForegroundColor Cyan
    Write-Host "We have 20 move request, so, I go to bed for " $SLEEP_TIME "seconds...." -ForegroundColor Cyan
    sleep $SLEEP_TIME


}
while(1 -lt 2)