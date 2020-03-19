$LOG_NAM = 'ccm.log'
$LOG_DAT = (Get-Date).ToShortDateString().Replace('/','-')
$LOG_TIM = (Get-Date).ToShortTimeString().Replace(' ','-').Replace(':','-')
$DB_LOG = Get-Content ('D:\Program Files\Microsoft Configuration Manager\Logs\' + $LOG_NAM) | Select-String -Pattern '---> ERROR: Unable to access target machine for request: '
"Objeto SCCM/Nombre de equipo/ADDS OU/Falla"| Out-File ('D:\Scripts\' + $LOG_NAM + '-' + $LOG_DAT + $LOG_TIM + '.csv') -Append
foreach ($IT_LOG in $DB_LOG){
    $SCM_LOG_LIN = $IT_LOG.Line

    $SCM_OBJ_STR = $SCM_LOG_LIN.IndexOf('"')+1
    $SCM_OBJ_END = $SCM_LOG_LIN.Substring($SCM_LOG_LIN.IndexOf('"')+1).IndexOf('"')
    $SCM_OBJ_NAM = $SCM_LOG_LIN.Substring($SCM_OBJ_STR,$SCM_OBJ_END)

    $COM_OBJ_LIN = $SCM_LOG_LIN.Substring($SCM_OBJ_STR+$SCM_OBJ_END+1)
    $COM_OBJ_STR = $COM_OBJ_LIN.IndexOf('"')+1
    $COM_OBJ_END = $COM_OBJ_LIN.Substring($COM_OBJ_LIN.IndexOf('"')+1).IndexOf('"')
    $COM_OBJ_NAM = $COM_OBJ_LIN.Substring($COM_OBJ_STR,$COM_OBJ_END)

    $TXT_OBJ_LIN = $COM_OBJ_LIN.Substring($COM_OBJ_STR+$COM_OBJ_END +4)
    $TXT_OBJ_STR = 0
    $TXT_OBJ_END = $TXT_OBJ_LIN.IndexOf('$')
    $TXT_OBJ_NAM = $TXT_OBJ_LIN.Substring($TXT_OBJ_STR,$TXT_OBJ_END)

    $ADS_OBJ_LIN = (Get-ADComputer $COM_OBJ_NAM | Select DistinguishedName).DistinguishedName
    $ADS_OBJ_STR = $ADS_OBJ_LIN.IndexOf(",")+1
    $ADS_OBJ_NAM = $ADS_OBJ_LIN.Substring($ADS_OBJ_STR)

    $SCM_OBJ_NAM + "/" + $COM_OBJ_NAM + "/" + $ADS_OBJ_NAM + '/' + $TXT_OBJ_NAM | Out-File ('D:\Scripts\' + $LOG_NAM + '-' + $LOG_DAT + $LOG_TIM + '.csv')  -Append
    
    #Move-Item "D:\Program Files\Microsoft Configuration Manager\inboxes\ccrretry.box\$SCM_OBJ_NAM.ccr" -Destination C:\Users\svc_cm_admin\Downloads\ccmretry
}
Remove-Item ('D:\Program Files\Microsoft Configuration Manager\Logs\' + $LOG_NAM) -Force
