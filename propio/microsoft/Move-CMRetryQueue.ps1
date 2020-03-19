$CM_LIM = 10
$CM_BAS = 0

$CM_RET = 'C:\Users\svc_cm_admin\Downloads\ccmretry\'
$CM_QUE = 'C:\Users\svc_cm_admin\Downloads\ccmqueue\'
$CM_SYS = 'D:\Program Files\Microsoft Configuration Manager\inboxes\ccrretry.box\'
$CM_EXT = '.ccr'

$DB_RET = Get-ChildItem $CM_SYS
if($DB_RET.Length -gt 0){
    do{

        $FIL_NAM = $DB_RET[$CM_BAS].Name
        $FIL_FNM = $DB_RET[$CM_BAS].FullName
        $RES_RID = $FIL_NAM.Substring(0,$FIL_NAM.Length - 4)
    
        Move-Item -Path $FIL_FNM -Destination $CM_RET -Force

        $CM_BAS++
    }
    until($DB_RET.Count -eq $CM_BAS)
}
else{
    Write-Host '-----------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
    Write-Host '                                      No se encuentra cola de reintentos que limpiar                                         ' -ForegroundColor Yellow
    Write-Host '-----------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
}