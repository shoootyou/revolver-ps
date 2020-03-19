Connect-MsolService
$TMP11 = Get-MSOLUser | Where-Object {$_.UserPrincipalName -like "*@tecnofor.pe"} | Select DisplayName,UserPrincipalName
$OUserOut = @()
foreach ($RGTMP11 in $TMP11){
        $TMP21 = Get-MsolUser -UserPrincipalName $RGTMP11.UserPrincipalName | Select -ExpandProperty ProxyAddresses
        
        foreach ($RGTMP22 in $TMP21) {
                            $ObjProperties = New-Object PSObject
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Nombre -Value $RGTMP11.DisplayName 
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name ProxyAddresses -Value $RGTMP22
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Domain -Value $RGTMP22.Split("@")[1]

                            if($RGTMP22 -like "*@TecnologiayCreatividad.onmicrosoft.com"){

                            #Set-Mailbox $RGTMP11.UserPrincipalName  -EmailAddresses @{remove="$RGTMP22"}

                            }

                            $OUserOut += $ObjProperties
        }
        #Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Nombre -Value $RGTMP22
                      
        }
        
        $OUserOut | Out-GridView -Title "Información Personal"
        #$OUserOut | Export-Csv C:\SharedFolders\AzureAD\proxyaddresses5.csv
        $ObjProperties =@()
        pause

         