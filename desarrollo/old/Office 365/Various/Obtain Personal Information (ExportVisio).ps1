$Empresa = Read-Host "Cual es el dominio de Office 365 que deseas consultar?"
$VariableConsulta = "*" + $Empresa + "*"
Write-Host "======================================================="
Write-Host "Validando y obteniendo Información, por favor espere..."
Write-Host "======================================================="
$VariableConsulta = "*" + $Empresa + "*"
$ValidaDominio = Get-AcceptedDomain | Select DomainName
foreach ($IDominio in $ValidaDominio) {
                $UsuarioPrimario = Get-User -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -like "$VariableConsulta") } | Select *
                $OUserOut = @()
                foreach ($OUser in $UsuarioPrimario) {

                    $ObjProperties = New-Object PSObject
                  
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name DisplayName -Value $OUser.DisplayName
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name FirstName -Value $OUser.FirstName
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name LastName -Value $OUser.LastName
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Manager -Value $OUser.Manager
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Title -Value $OUser.Title
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Department -Value $OUser.Department
                    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Office -Value $OUser.office
                    $OUserOut += $ObjProperties
}
                $OUserOut | Out-GridView -Title "Información Personal"
                $OUserOut |  Export-Csv C:\Users\Rodolfo\Desktop\Exportación.csv
                break
           
}
