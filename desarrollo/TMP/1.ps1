foreach($i in Get-User cecilia.torres@tecnofor.pe | Select * ) {

　 $i.EmailAddresses |
　　　 ?{$_.AddressString -like '*@tecnologiaycreatividad.onmicrosoft.com'} | %{
　　　　　 Set-Mailbox $i -EmailAddresses @{remove=$_}
　　　 }
}


Get-ADUser rodolfo.castelo


Get-Mailbox rodolfo.castelo@tecnofor.pe | Select -ExpandProperty EmailAddresses | Set-Mailbox -EmailAddresses @{remove="cecilia.torres@tecnologiaycreatividad.onmicrosoft.com"}

$UniversoTMP = Get-Mailbox | Where-Object {$_.WindowsLiveID -like "*@tecnofor.pe"}
foreach($personita in $UniversoTMP){
        



}

Set-Mailbox 