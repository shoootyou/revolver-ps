ForEach-Object{
    $outlook = New-Object -comobject outlook.application
    $msg = $OUT_APP_COM.CreateItemFromTemplate($MSG)
    $msg | Select senderemailaddress,to,subject,Senton,body|ft -AutoSize
    }

    $OUT_APP_COM



    $adoDbStream = New-Object -ComObject ADODB.Stream
$adoDbStream.Open()
$adoDbStream.LoadFromFile($MSG)

RDOSession Session = new RDOSession();
RDOMail Msg = Session.CreateMessageFromMsgFile(@"c:\temp\YourMsgFile.msg");
Msg.Import(@"c:\temp\YourEmlFile.eml", rdoSaveAsType.olRFC822);
Msg.Save();


Function Load-EmlFile
{
Param
(
$EmlFileName
)
$AdoDbStream = New-Object -ComObject ADODB.Stream
$AdoDbStream.Open()
$AdoDbStream.LoadFromFile($PAT_EML)
$CdoMessage = New-Object -ComObject CDO.Message
$CdoMessage.DataSource.OpenObject($AdoDbStream,"_Stream")

return $CdoMessage
}

$TRANS_OUT_APP = New-Object -comObject Outlook.Application 
$TRANS_REDEM = New-Object -ComObject Redemption.SafePostItem
$TRANS_TEMPO = $TRANS_OUT_APP.CreateItem('olPostItem')

$TRANS_REDEM.Item = $TRANS_TEMPO
$TRANS_REDEM.Import($PAT_EML,1024)
$TRANS_REDEM.MessageClass = "IPM.Note"
$TRANS_REDEM.SaveAs($EXP_MSG,3)

$outlook = New-Object -COM Outlook.Application 
$routlook = New-Object -COM Redemption.RDOSession
$routlook.Logon
$routlook.MAPIOBJECT = $outlook.Session.MAPIOBJECT


$routlook.CreateMessageFromMsgFile(

$msg2 = $routlook.GetFolderFromPath($INT_FOLD_02).Items.Add(0)

$msg2.Import($PAT_MSG_2)
$msg2.save() | Out-Null
$routlook.Logoff | Out-Null

set Session = CreateObject("Redemption.RDOSession") 

set Msg = Session.CreateMessageFromMsgFile("C:\Temp\test.msg") 

Msg.Import "C:\Temp\test.eml", 1024 

Msg.Save 


$outlook = New-Object -COM Outlook.Application 
$routlook = New-Object -COM Redemption.RDOSession
$routlook.Logon
$routlook.MAPIOBJECT = $outlook.Session.MAPIOBJECT


$MHD = $routlook.CreateMessageFromMsgFile('E:\Users\Rodolfo\Desktop\testing@americatel.com.pe\mail\DOS\(Serrano Misayauri, Roberto) Clase 5.msg')
$MHD.Import('E:\Users\Rodolfo\Desktop\testing@americatel.com.pe\mail\DOS\(Serrano Misayauri, Roberto) Clase 5.eml',1024)
$MHD.Save()
$MHD.

$msg2 = $routlook.GetFolderFromPath($INT_FOLD_02).Items.Add(0)

$msg2.Import($PAT_MSG_2)
$msg2.save() | Out-Null
$routlook.Logoff | Out-Null