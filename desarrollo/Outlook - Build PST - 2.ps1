$Date_1 = Get-Date
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '                         Proceso iniciado: ' $Date_1 -ForegroundColor Green
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host
##############################################################################################################
$GBL_PAT_MS = 'E:\Users\Rodolfo\Downloads\DOS'
$GBL_PAT_EX = 'E:\Users\Rodolfo\Downloads\DOS\PST\'
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
            $MON_EML_DB = Get-ChildItem -Path $INT_FOLD | Where {($_.Attributes -eq 'Directory') -and ($_.Name -like '201*')}| Select FullName,Name
            if($MON_EML_DB -ne $null){
                #---------------------------------- Sorting items in root folder ----------------------------------
                foreach($MON_FOL in $MON_EML_DB){
                    $TMP_01 = $MON_FOL.FullName
                    Get-ChildItem -Path "$TMP_01\*.eml" | Move-Item -Destination $INT_FOLD -Force -Confirm:$false
                    Remove-Item -Path $TMP_01 -Force -Confirm:$false
                }
                #---------------------------------------------------------------------------------------------------
            }
        }
}
$Date_2 = Get-Date
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '                         Proceso Finalizado: ' $Date_2 -ForegroundColor Green
Write-Host '---------------------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host
pause