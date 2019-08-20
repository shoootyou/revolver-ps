$Empresa = Read-Host "Cual es el dominio de Office 365 que deseas consultar?"
$VariableConsulta = "*" + $Empresa + "*"
Write-Host
Write-Host "================================================================================================"
Write-Host "                    Validando y obteniendo Informaci�n, por favor espere..."
Write-Host "================================================================================================"
Write-Host
$SIExisteDom = 0
$VariableConsulta = "*" + $Empresa + "*"
$ValidaDominio = Get-AcceptedDomain | Select DomainName
foreach ($IDominio in $ValidaDominio) {
        if($IDominio -like $VariableConsulta){
                $SIExisteDom++
                IF($SIExisteDom -eq 1){

#############################################################################################################################

                ##################################################
                #### Informaci�n de ubicaci�n variables
                #################################################
                $SiUbica = New-Object System.Management.Automation.Host.ChoiceDescription "&S�", ` "S�"
                $NoUbica = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
                $OpcionesUbica = [System.Management.Automation.Host.ChoiceDescription[]]($SiUbica, $NoUbica)
                $ResultadosUbica = $host.ui.PromptForChoice("Confirmaci�n", "�Desea verificar la informaci�n de ubicaci�n?", $OpcionesUbica, 0) 
                #################################################
                $UsuarioPrimario = Get-User -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -like "$VariableConsulta")} | Select *
                $OUserOut = @()
                foreach ($OUser in $UsuarioPrimario) {

                    $ObjProperties = New-Object PSObject
                             
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Identity -Value $OUser.Identity
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name DisplayName -Value $OUser.DisplayName
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name FirstName -Value $OUser.FirstName
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name LastName -Value $OUser.LastName
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Office -Value $OUser.office
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Phone -Value $OUser.Phone
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name WindowsEmailAddress -Value $OUser.WindowsEmailAddress
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Department -Value $OUser.Department
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Title -Value $OUser.Title
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Manager -Value $OUser.Manager
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Company -Value $OUser.Company

                    switch ($ResultadosUbica){
                            0 {
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name PostalCode -Value $OUser.PostalCode
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name StreetAddress -Value $OUser.StreetAddress                        
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name City -Value $OUser.City
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name StateOrProvince -Value $OUser.StateOrProvince
                            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name CountryOrRegion -Value $OUser.CountryOrRegion
                            }
                            1 {
                        
                            }
                        }
                    $OUserOut += $ObjProperties
}
                $OUserOut | Out-GridView -Title "Informaci�n Personal"
                $OUserOut | Export-Csv C:\Users\Rodolfo\Desktop\Exportaci�n.csv         
                
##############################################################################################################################
                Break
                }
        }
}
        if($SIExisteDom -ne 1){
                Write-Host "No posees dicho dominio"
        }
