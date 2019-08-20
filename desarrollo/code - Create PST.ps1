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

        $OUT_NAME.AddStoreEx($USR_PST,3)

        ########################################################################
        #                      Creating  Folder Structure                      #
        ########################################################################
        $MSG_FOL_DB = Get-ChildItem -Path $USR_MSGS | Select FullName,Name
        foreach($FOL_FOR_01 in $MSG_FOL_DB){
            $FOL_NAME = $FOL_FOR_01.Name
            $INT_FOLD = $FOL_FOR_01.FullName
                    
            Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
            $OUT_APP_COM = new-object -comobject outlook.application 
            $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

            $PST_PAT = $OUT_NAME.Session.Folders.GetLast() 
            [void]$PST_PAT.Folders.Add($FOL_NAME)
        
            ########################################################################
            #                        Sorting Email messages                        #
            ########################################################################
            $TIM_FOL_DB = Get-ChildItem -Path $INT_FOLD | Select FullName,Name
            foreach($FOL_MSG in $TIM_FOL_DB){
                $MSG_DB = Get-ChildItem $FOL_MSG.FullName
                foreach($MSG in $MSG_DB){
                    $PAT_MSG = $MSG.FullName
                    if($CON_UN -gt 1){
                        
                        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
                        $OUT_APP_COM = new-object -comobject outlook.application 
                        $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

                        $OUT_NAME.OpenSharedItem($PAT_MSG).Move($OUT_NAME.Folders(1).Folders(2)) | Out-Null
                    }
                    else{
                    
                        Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
                        $OUT_APP_COM = new-object -comobject outlook.application 
                        $OUT_NAME = $OUT_APP_COM.GetNameSpace("MAPI")

                        $OUT_NAME.OpenSharedItem($PAT_MSG).Move($OUT_NAME.Folders(2).Folders(2)) | Out-Null
                    }

                }
            }

        }

}



#$OUT_NAME.

