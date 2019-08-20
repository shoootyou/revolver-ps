Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")

$TMP = $NameSpace.Folders | select FolderPath

$CNT_GBL = 1
$INT_01 = 1
foreach($ACC in $TMP){
if($ACC.FolderPath -like '*soporte.peru*'){
    $INT_02 = 1
    $INT_PT_1 = $NameSpace.Folders.Item($INT_01).Folders | select FolderPath
    foreach($INT_FP_01 in $INT_PT_1){
        if($INT_FP_01 -like '*ReboteProceso*'){
            Write-Host "Information Found"
            Write-Host "Please wait ...."
            $ML_DB = $NameSpace.Folders.Item($INT_01).Folders.Item($INT_02).items | select Body,Subject
        }
        else{
            $INT_02++
        }

    }
}
else{
    $INT_01++

}
}

$GBL_COU = 1

foreach($ML in $ML_DB){
    $ML_Body = $ML.Body
    $WR_PROG = $ML.Subject

    if(($ML_Body.IndexOf("Original")) -ge 1){
    $SRC_NAME_STRG_POS_1 = $ML_Body.IndexOf("Original")
    $SRC_NAME_STRG_PH_1 = $ML_Body.Substring(1,$SRC_NAME_STRG_POS_1)
    $SRC_NAME_STRG_POS_2 = $SRC_NAME_STRG_PH_1.IndexOf("@")
    $SRC_NAME_STRG_PH_2 = $SRC_NAME_STRG_PH_1.Substring(1,$SRC_NAME_STRG_POS_2)
    $SRC_NAME_STRG_POS_3 = $SRC_NAME_STRG_PH_2.LastIndexOf(" ")
    $SRC_NAME_STRG_PH_3 = $SRC_NAME_STRG_PH_2.Substring($SRC_NAME_STRG_POS_3)
    $SRC_NAME_STRG_POS_4 = $SRC_NAME_STRG_PH_3.Length
    $SRC_NAME_STRG_PH_4 = $SRC_NAME_STRG_PH_3.Substring(1,$SRC_NAME_STRG_POS_4-1)

    

    $SRC_DOM_STRG_POS_1 = $ML_Body.IndexOf("Original")
    $SRC_DOM_STRG_PH_1 = $ML_Body.Substring(1,$SRC_DOM_STRG_POS_1)
    $SRC_DOM_STRG_POS_2 = $SRC_DOM_STRG_PH_1.IndexOf("@")
    $SRC_DOM_STRG_PH_2 = $SRC_DOM_STRG_PH_1.Substring($SRC_DOM_STRG_POS_2)
    $SRC_DOM_STRG_POS_3 = $SRC_DOM_STRG_PH_2.IndexOf("`n")
    $SRC_DOM_STRG_PH_3 = $SRC_DOM_STRG_PH_2.Substring(1,$SRC_DOM_STRG_POS_3)
    $SRC_DOM_STRG_POS_4 = $SRC_DOM_STRG_PH_3.Length
    $SRC_DOM_STRG_PH_4 = $SRC_DOM_STRG_PH_3.Substring(0,$SRC_DOM_STRG_POS_4-1)

    $SRC_NAME_STRG_PH_4 + $SRC_DOM_STRG_PH_4
    }

    Write-Progress -Activity “Updating Information” -status “Updating $WR_PROG” -percentComplete ($GBL_COU / $ML_DB.Count*100)
}