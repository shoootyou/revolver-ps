##############################################################################################################
$GBL_PAT_MS = 'E:\Users\Rodolfo\Desktop\'
$GBL_PAT_EX = 'E:\Users\Rodolfo\Desktop\PST\'
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
        $OUT_NAME.AddStoreEx($USR_PST,3)
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host '                                    PST created succefully                                   ' -ForegroundColor Green
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
        Write-Host
        ########################################################################
        #                      Creating  Folder Structure                      #
        ########################################################################
        $MSG_FOL_DB = Get-ChildItem -Path $USR_MSGS | Select FullName,Name
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Yellow
        Write-Host '                      Starting the process: Building PST | Please Wait                       ' -ForegroundColor Yellow
        Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Yellow
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
                        $STO_ID = $STORE.StoreID # Se utiliza para poder ubicar el PST creado
                    }
                    if($STORE.Name -eq $USR_ALIA){
                        $STO_ID = $STORE.StoreID # Se utiliza para poder ubicar el PST creado
                    }
            }
            $PST_PAT = $OUT_NAME.Session.GetFolderFromID($STO_ID)
            $FOL_ID = $PST_PAT.Folders.Add($FOL_NAME) #Se utiliza para poder ubicar la carpeta que estamos creando
        
            ########################################################################
            #                        Moving Email messages                         #
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
            ##################################################################################################################################################
            #                                                               Convert EML to MSG                                                               #
            ##################################################################################################################################################
            $FOL_EML_DB = Get-ChildItem -Path $INT_FOLD | Where {($_.Name -like '*.eml')} | Select FullName,Name
            if($FOL_EML_DB -ne $null){
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Yellow
                Write-Host "            Transforming EML files to MSG. Please Wait, don't finish the process             " -ForegroundColor Yellow
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Yellow
                foreach($EML in $FOL_EML_DB){
                    $PROGRESS_01 = 1
                    $PAT_MSG = $EML.FullName
                    $NAM_MSG = $EML.Name
                    $EXP_MSG = $PAT_MSG.Substring(0,$PAT_MSG.LastIndexOf('.')) + '.msg'
                    #---------------------------------- Transforming EML to MSG ----------------------------------
                    $TRANS_OUT_APP = New-Object -comObject Outlook.Application 
                    $TRANS_REDEM = New-Object -ComObject Redemption.SafePostItem
                    $TRANS_TEMPO = $TRANS_OUT_APP.CreateItem('olPostItem')

                    $TRANS_REDEM.Item = $TRANS_TEMPO
                    $TRANS_REDEM.Import($PAT_MSG,1024)
                    $TRANS_REDEM.MessageClass = "IPM.Note"
                    $TRANS_REDEM.SaveAs($EXP_MSG,3)
                    #---------------------------------------------------------------------------------------------
                    $PROGRESS_01++
                }
                Get-ChildItem -Path $INT_FOLD | Where {($_.Name -like '*.eml')} | Remove-Item
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
                Write-Host "            Transforming for ' $FOL_NAME ' folder, completed                       " -ForegroundColor Green
                Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
            }
            ####################################################################################################################################################
            $IMP_MSG_DB = Get-ChildItem -Path $INT_FOLD | Where {($_.Name -like '*.msg')} | Select FullName,Name
            if($IMP_MSG_DB){
                foreach($MSG in $IMP_MSG_DB){
                    $PAT_MSG_2 = $MSG.FullName
                                  
                    $outlook = New-Object -COM Outlook.Application 
                    $routlook = New-Object -COM Redemption.RDOSession
                    $routlook.Logon
                    $routlook.MAPIOBJECT = $outlook.Session.MAPIOBJECT

                    $msg2 = $routlook.GetFolderFromID($FOL_ID.EntryID).Items.Add(6)
                                    
                    $msg2.Import($PAT_MSG_2,3)
                    $msg2.Subject = $msg2.Subject.Substring($msg2.Subject.IndexOf(" ")+1,$msg2.Subject.Length-($msg2.Subject.IndexOf(" ")+1))
                    $msg2.ConversationTopic = $msg2.ConversationTopic.Substring($msg2.ConversationTopic.IndexOf(" ")+1,$msg2.ConversationTopic.Length-($msg2.ConversationTopic.IndexOf(" ")+1))
                    $msg2.BodyFormat = 2
                    $msg2.save() | Out-Null
                    $routlook.Logoff | Out-Null
                }
            }
            
            <#
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
                            if($INT_FOLD_01 -eq $FOL_NAME){
                                $CON_UN = 0
                                foreach($MSG in $IMP_MSG_DB){
                                    $PAT_MSG = $MSG.FullName
                                    #if($CON_UN -gt 1){
                                        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
                                        $OUT_APP_COM = new-object -comobject outlook.application 
                                        $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

                                        
                                        $OUT_NAME.OpenSharedItem($PAT_MSG).Move($OUT_NAME.Folders($FOL_IND).Folders($FOL_FIN_01)) | Out-Null
                                        Write-Host '444'
                                    }
                                    else{
                                        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
                                        $OUT_APP_COM = new-object -comobject outlook.application 
                                        $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

                                        $OUT_NAME.OpenSharedItem($PAT_MSG).Move($OUT_NAME.Folders($FOL_IND-1).Folders($FOL_FIN_01-1)) | Out-Null
                                    }
                                    $CON_UN++
                                }
                            }
                            $FOL_FIN_01++
                        }
                    }
                    $FOL_IND++
                }
            }#>


            
        }
}

