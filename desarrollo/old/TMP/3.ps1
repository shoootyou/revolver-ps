Import-Module ActiveDirectory

$newproxy = "domainname.com"
$userou = 'ou=Users,ou=Directory Sync,ou=Org,dc=domain,dc=domain,dc=com'
$users = Get-ADUser -Filter * -SearchBase $userou -Properties SamAccountName, ProxyAddresses, givenName, Surname

Foreach ($user in $users) {
    Set-ADUser -Identity $_.SamAccountName -Add @{Proxyaddresses="SMTP:"+$_.givenName+"."+$_.Surname+$proxydomain} -whatif
    } 

    Import-Module ActiveDirectory
    Set-ADUser -Identity rodolfo.castelo -Add @{Proxyaddresses="SMTP:rodolfo.castelo@tecnofor.pe,smtp:rodolfo.castelo@tecnofor.com.pe,smtp:rcastelo@tecnofor.pe,smtp:rcastelo@tecnofor.com.pe"}


    Get-ADUser -Identity rodolfo.castelo | ft UserPrincipalname,ProxyAddresses


    Import-Module Microsoft.ActiveDirectory.Management.ADPropertyValueCollection


    Get-ADUser -Identity rodolfo.castelo -Properties proxyaddresses | select name, proxyaddresses | Export-CSV -Path C:\SharedFolders\AzureAD\proxyaddresses.csv –NoTypeInformation



    Get-ADUser -Identity rodolfo.castelo -Properties proxyaddresses | Select-Object Name, @{L = "ProxyAddresses"; E = { $_.ProxyAddresses -join ";"}} |Export-Csv -Path C:\SharedFolders\AzureAD\proxyaddresses.csv -NoTypeInformation

    Get-Mailbox rodolfo.castelo@tecnofor.pe | Select-Object Name, @{L = "EmailAddresses"; E = { $_.EmailAddresses -join ";"}} |Export-Csv -Path C:\SharedFolders\AzureAD\proxyaddresses2.csv -NoTypeInformation


    $UniversoCloud = Get-Mailbox -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -like "*tecnofor.pe*")} | Select *
    foreach($nube in $UniversoCloud){
            $TMP1 = $nube.Alias + "@tecnofor.pe"
            Write-Host $TMP1
            Get-MsolUser -UserPrincipalName $TMP1 | Select -ExpandProperty ProxyAddresses  |Export-Csv -Path C:\SharedFolders\AzureAD\proxyaddresses2.csv -NoTypeInformation

    }

    Get-MsolUser -UserPrincipalName rodolfo.castelo@tecnofor.pe | Select * -ExpandProperty ProxyAddresses
     
    Get-Mailbox | Select Alias

    Get-MsolUser | Where-Object {$_.UserPrincipalName -like "*@tecnofor.pe*"}| Select  DisplayName,UserPrincipalName,ObjectId |Export-Csv -Path C:\SharedFolders\AzureAD\proxyaddresses3.csv




        
