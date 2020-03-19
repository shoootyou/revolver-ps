$Date_1 = Get-Date
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '                         Proceso iniciado: ' $Date_1 -ForegroundColor Green
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host
##############################################################################################################
$GBL_PAT_MS = 'C:\Users\d3pl0y\Downloads\'
$GBL_PAT_EX = 'C:\Users\d3pl0y\Downloads\PST\'
$GBL_USR_DB = Get-ChildItem -Path $GBL_PAT_MS | Where {($_.Attributes -eq 'Directory') -and ($_.FullName -like '*@*')} | Select FullName,Name
foreach($USR_FOR_01 in $GBL_USR_DB){
        $USR_NAME = $USR_FOR_01.Name
        $USR_MSGS = $USR_FOR_01.FullName + '\mail\'
        $USR_ALIA = $USR_NAME.Substring(0,$USR_NAME.IndexOf('@'))

        ########################################################################
        #                             Creating PST                             #
        ########################################################################
        $USR_PST = $GBL_PAT_EX + $USR_NAME + '.pst'

        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
        $OUT_APP_COM = new-object -comobject outlook.application 
        $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '                    Creating PST for user: ' $USR_NAME -ForegroundColor Cyan
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host
        $OUT_NAME.AddStoreEx($USR_PST,2)
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host '                                    PST created succefully                                   ' -ForegroundColor Green
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host
        ########################################################################
        #                      Creating  Folder Structure                      #
        ########################################################################
        $MSG_FOL_DB = Get-ChildItem -Path $USR_MSGS | Select FullName,Name
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '                      Starting the process: Building PST | Please Wait                       ' -ForegroundColor Cyan
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host
        foreach($FOL_FOR_01 in $MSG_FOL_DB){
            $FOL_NAME = $FOL_FOR_01.Name
            $INT_FOLD = $FOL_FOR_01.FullName
                    
            Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
            $OUT_APP_COM = new-object -comobject outlook.application 
            $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

            $DB_STO = $OUT_NAME.Folders
                foreach($STORE in $DB_STO){
                    if($STORE.Name -like '*Outlook*'){
                        $STORE.Name = $USR_ALIA
                        $STO_ID = $STORE.StoreID
                    }
            }
            $PST_PAT = $OUT_NAME.Session.GetFolderFromID($STO_ID)
            [void]$PST_PAT.Folders.Add($FOL_NAME)
        
            ########################################################################
            #                        Sorting Email messages                        #
            ########################################################################
            $MON_EML_DB = Get-ChildItem -Path $INT_FOLD | Where {($_.Attributes -eq 'Directory')}| Select FullName,Name
            if($MON_EML_DB -ne $null){
                #---------------------------------- Sorting items in root folder ----------------------------------
                foreach($MON_FOL in $MON_EML_DB){
                    $TMP_01 = $MON_FOL.FullName
                    Get-ChildItem -Path "$TMP_01\*.eml" | Move-Item -Destination $INT_FOLD -Force -Confirm:$false
                    Remove-Item -Path $TMP_01 -Force -Confirm:$false
                }
                #---------------------------------------------------------------------------------------------------
            }
            $FOL_EML_DB = Get-ChildItem -Path $INT_FOLD | Where {($_.Name -like '*.eml')} | Select FullName,Name
            if($FOL_EML_DB -ne $null){
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Yellow
                Write-Host "            Transforming EML files to MSG of $FOL_NAME folder " -ForegroundColor Yellow
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Yellow
                foreach($EML in $FOL_EML_DB){
                    $PROGRESS_01 = 1
                    $PAT_EML = $EML.FullName
                    $NAM_EML = $EML.Name
                    $EXP_MSG = $PAT_EML.Substring(0,$PAT_EML.LastIndexOf('.')) + '.msg'
                    #---------------------------------- Transforming EML to MSG ----------------------------------
                    $TRANS_OUT_APP = New-Object -comObject Outlook.Application 
                    $TRANS_REDEM = New-Object -ComObject Redemption.SafePostItem
                    $TRANS_TEMPO = $TRANS_OUT_APP.CreateItem('olMailItem')

                    
                    $TRANS_REDEM.Item = $TRANS_TEMPO
                    $TRANS_REDEM.Item.InternetCodepage = 65001
                    $TRANS_REDEM.Import($PAT_EML,1024)
                    #$TRANS_REDEM.Item.BodyFormat = 2
                    $TRANS_REDEM.Item.MessageClass = "IPM.Note"
                    $TRANS_REDEM.Item.InternetCodepage = 65001
                    
                    $TRANS_REDEM.Item.Save()
                    $TRANS_REDEM.SaveAs($EXP_MSG,3)

                    #---------------------------------------------------------------------------------------------
                    $PROGRESS_01++
                }
                #Get-ChildItem -Path $INT_FOLD | Where {($_.Name -like '*.eml')} | Remove-Item
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
                Write-Host "            Transforming for $FOL_NAME folder, completed                       " -ForegroundColor Green
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
            }
            $IMP_MSG_DB = Get-ChildItem -Path $INT_FOLD | Where {($_.Name -like '*.msg')} | Select FullName,Name
            if($IMP_MSG_DB){
                #----------   Finding PST ----------
                $DB_STO = $OUT_NAME.Folders
                #-----------------------------------
                $FOL_IND = 0
                foreach($STORE in $DB_STO){
                    if($STORE.Name -eq $USR_ALIA){
                        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
                        $OUT_APP_COM = new-object -comobject outlook.application 
                        $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

                        $DB_FOLDERS = $OUT_NAME.Session.Folders.GetLast().Folders
                        $FOL_FIN_01 = 1
                        foreach($Folder in $DB_FOLDERS){
                            $INT_FOLD_01 = $Folder.Name
                            $INT_FOLD_02 = $Folder.FolderPath
                            $INT_FOLD_03 = $Folder.EntryID
                            if($INT_FOLD_01 -eq $FOL_NAME){
                                $CON_UN = 0
                                foreach($MSG in $IMP_MSG_DB){
                                    $PAT_MSG_2 = $MSG.FullName
                                  
                                    $outlook = New-Object -COM Outlook.Application 
                                    $routlook = New-Object -COM Redemption.RDOSession
                                    $routlook.Logon | Out-Null
                                    $routlook.MAPIOBJECT = $outlook.Session.MAPIOBJECT
                                    <#
                                    $msg1 = $routlook.GetFolderFromPath($INT_FOLD_02).Items.Add(6)
                                    #$msg2.BodyFormat = 2
                                    $msg1.InternetCodepage = 65001
                                    $msg1.Import($PAT_MSG_2,1024)#>

                                    $msg2 = $routlook.GetFolderFromPath($INT_FOLD_02).Items.Add(6)
                                    #$msg2.BodyFormat = 2
                                    $msg2.InternetCodepage = 65001
                                    $msg2.Import($PAT_MSG_2,3)
                                    $msg2.InternetCodepage = 65001
                                    #$msg2.Subject = $msg2.Subject.Substring($msg2.Subject.IndexOf(" ")+1,$msg2.Subject.Length-($msg2.Subject.IndexOf(" ")+1))
                                    #$msg2.Subject = $msg1.Subject
                                    $msg2.Subject = $msg2.ConversationTopic
                                    #$msg2.ConversationTopic = $msg2.ConversationTopic.Substring($msg2.ConversationTopic.IndexOf(" ")+1,$msg2.ConversationTopic.Length-($msg2.ConversationTopic.IndexOf(" ")+1))
                                    $msg2.Sent = $true
                                    $msg2.InternetCodepage = 65001
                                    $msg2.save() | Out-Null
                                    $routlook.Logoff | Out-Null
                                    
                                }
                            }
                            $FOL_FIN_01++
                        }
                    }
                    $FOL_IND++
                }
            }
            ### Section to Remove items from drafts
            do{
            $REM_DB_BOR = $routlook.GetDefaultFolder(16).items
                foreach($MS_BOR in $REM_DB_BOR){
                    $MS_BOR.Delete()
                }
            }
            while($REM_DB_BOR.Count -gt 0)

            do{
            $REM_DB_ELI = $routlook.GetDefaultFolder(3).items
                foreach($MS_BOR in $REM_DB_ELI){
                    $MS_BOR.Delete()
                }
            }
            while($REM_DB_ELI.Count -gt 0)
        }
<#
$OUT_PST = new-object -com outlook.application 
$OUT_NAS = $OUT_PST.getNamespace("MAPI")

$OUT_PTD = $USR_PST
$OUT_STO = $OUT_NAS.Stores | ? {$_.FilePath -eq $OUT_PTD}
$OUT_ROT = $OUT_STO.GetRootFolder()

$OUT_FOL = $OUT_NAS.Folders.Item($OUT_ROT.Name)
$OUT_NAS.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$OUT_NAS,($OUT_FOL))

Get-Item -Path $USR_MSGS | Remove-Item -Recurse -Force#>
}
$Date_2 = Get-Date
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '                         Proceso Finalizado: ' $Date_2 -ForegroundColor Green
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host
pause