$Output = @()
foreach($FirstLine in (Get-MSOLUser -MaxResults 10000))
{
    $UserInfo = Get-MSOLUser -UserPrincipalName $FirstLine.UserPrincipalName
    foreach($license in $FirstLine.Licenses)
    {
        $ConvUser = Get-User $FirstLine.UserPrincipalName | Select Company,WindowsEmailAddress
        $ConvMX = Get-Mailbox -Identity $FirstLine.UserPrincipalName | Select CustomAttribute1
        $OutTMP = New-Object PsObject 
        Add-Member -InputObject $OutTMP -MemberType NoteProperty -Name Usuario -Value $UserInfo.DisplayName
        Add-Member -InputObject $OutTMP -MemberType NoteProperty -Name Correo -Value $ConvUser.WindowsEmailAddress
        Add-Member -InputObject $OutTMP -MemberType NoteProperty -Name Organizacion -Value $ConvMX.CustomAttribute1
           If($license.AccountSKUid -eq "TecnologiayCreatividad:ENTERPRISEPACK"){
            Add-Member -InputObject $OutTMP -MemberType NoteProperty -Name Licencia -Value "Office 365 Enterprise E3"
           }
           Elseif($license.AccountSKUid -eq "TecnologiayCreatividad:EMS"){
            Add-Member -InputObject $OutTMP -MemberType NoteProperty -Name Licencia -Value "Enterprise Mobility Suite"
           }
           Elseif($license.AccountSKUid -eq "TecnologiayCreatividad:CRMIUR"){
            Add-Member -InputObject $OutTMP -MemberType NoteProperty -Name Licencia -Value "Microsoft Dynamics CRM"
           }
        
        $Output += $OutTMP
    }
}
$Output | Out-GridView -Title "Información Personal"
#$Output | Export-Csv$env:USERPROFILE\desktop\Licencias.csv

