
#foreach($ArchiveUser in $ArchiveUsers){

for ($t=1; $t -le 5; $t++) {

    $AUPN = 'atlopez@americatel.com.pe'
    Start-ManagedFolderAssistant -identity $AUPN
 
    Write-Host '------------------------------------------------------'
    sleep 300
    }



#}