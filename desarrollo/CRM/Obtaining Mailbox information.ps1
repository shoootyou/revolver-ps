Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")
#$NameSpace.Folders.Item(4).Folders | FT FolderPath

 $OBJ_OUT = @()
 $i=0
$NameSpace.Folders.Item(1).Folders.Item('Bandeja de Entrada').Folders.Item('Rebotes').Folders.Item('[RCM] 06 04 16').items `
 | Select Subject,Body,Attachments,SenderName,SenderEmailAddress | ForEach-Object{
         $MSGBody  = $_.Body
         $Sub01 = $MSGBody.Substring(0,$MSGBody.IndexOf('@')+50)
         Write-Progress -Activity “Obtaining Information” -status “Updating $_.Subject” -percentComplete ($i / $accounts.Count*100)
        
        
        if($Sub01.Contains('@')){

        $OBJ_OUT_PRO = New-Object PSObject
        Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Body' -Value  $Sub01 
        $OBJ_OUT += $OBJ_OUT_PRO
        }

        $i++
 }

 $OBJ_OUT | Out-GridView -Title "Información Personal"


 $OBJ_OUT | Export-Csv $env:USERPROFILE\Desktop\Testing.csv





 $NameSpace.Folders.Item(1).Folders.Item('Bandeja de Entrada').Folders.Item('Rebotes').Folders.Item('[RCM] 06 04 16').items `
 | Select Subject,Body,Attachments,SenderName,SenderEmailAddress | ForEach-Object {
 $MSGBody  = $_.Body

        $OBJ_OUT_PRO = New-Object PSObject
        Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name 'Body' -Value $MSGBody 
        $OBJ_OUT += $OBJ_OUT_PRO



 }

  $OBJ_OUT | Out-GridView -Title "Información Personal"

 $OBJ_OUT | Export-Csv $env:USERPROFILE\Desktop\Testing.csv