$CM_SLE = 3600
$CM_LIM = 10
$CM_CLE = 100
$CM_BAS = 0
$CM_INS = 0
$DB_CLI = Get-CMCollectionMember -CollectionId 'FOH00018' | ? { $_.ResourceID -like '2*'} | Select Name,ResourceID

$CM_RET = 'C:\Users\svc_cm_admin\Downloads\ccmretry\'
$CM_QUE = 'C:\Users\svc_cm_admin\Downloads\ccmqueue\'
$CM_SYS = 'D:\Program Files\Microsoft Configuration Manager\inboxes\ccrretry.box\'
$CM_EXT = '.ccr'
$SC_PAT = Split-Path $MyInvocation.MyCommand.Path
$SC_CLE = $SC_PAT + '\Move-CMRetryQueue.ps1'

do{
    $CLI_NAM = $DB_CLI[$CM_BAS].Name
    $CLI_RID = $DB_CLI[$CM_BAS].ResourceID

    $CLE_MAR = $CM_BAS/$CM_CLE
    $CLE_LIM = $CM_INS/$CM_LIM
     
    If(
        $CLE_MAR.GetType().Name -ne 'Double'
    ){
        Invoke-Expression $SC_CLE
    }
    
    If(
        $CM_INS -eq 10
    ){
        $CM_INS = 0
    }
    
    Write-Host 'Revisando equipo: ' $CLI_NAM  -ForegroundColor Cyan

    if(  
        !(Get-Item ($CM_RET + $CLI_RID + $CM_EXT) -ErrorAction SilentlyContinue) -and 
        !(Get-Item ($CM_QUE + $CLI_NAM + $CM_EXT) -ErrorAction SilentlyContinue) -and 
        !(Get-Item ($CM_SYS + $CLI_RID + $CM_EXT) -ErrorAction SilentlyContinue) -and
        $CLE_LIM.GetType().Name -eq  'Double'
    ){
        Write-Host 'Iniciando instalación en equipo: ' $CLI_NAM  -ForegroundColor Green
        Install-CMClient -DeviceId $MBM.ResourceID -SiteCode 'FOH' | Out-Null
        New-Item ($CM_QUE + $CLI_NAM + $CM_EXT) | Out-Null
    }
    if($CLE_LIM.GetType().Name -ne 'Double'){
        if($CLE_LIM -ne 0){
            Write-Host '-----------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
            Write-Host '                               Script durmiendo, se reanudará el proceso en ' $CM_SLE ' segundos                                    ' -ForegroundColor Yellow
            Write-Host '-----------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
            Start-Sleep -Seconds $CM_SLE
        }
    }

    if($DB_CLI.Count-1 -eq $CM_BAS){
        $CM_BAS = -1
        $DB_CLI = Get-CMCollectionMember -CollectionId 'FOH00018' | ? { $_.ResourceID -like '2*'} | Select Name,ResourceID
        Remove-Item ($CM_RET + '*') -Force -ErrorAction SilentlyContinue -Confirm:$false
        Remove-Item ($CM_QUE + '*') -Force -ErrorAction SilentlyContinue -Confirm:$false
    }
    $CM_INS++
    $CM_BAS++
}
until($DB_CLI.Count -eq $CM_BAS)