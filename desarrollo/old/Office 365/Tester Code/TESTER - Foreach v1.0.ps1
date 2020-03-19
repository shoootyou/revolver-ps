$Empresa = Read-Host "Cual es el dominio de Office 365 que deseas consultar?"
$VariableConsulta = "*" + $Empresa + "*"
Write-Host "======================================================="
Write-Host "Validando y obteniendo Información, por favor espere..."
Write-Host "======================================================="
$ValidaDominio = Get-AcceptedDomain | Select DomainName
foreach ($IDominio in $ValidaDominio) {
            If($IDominio -like $VariableConsulta){
            Write-Host "Exito"
            $DominioExistente++
            break
            }
            If($DominioExistente -ne 1){
            Write-Host "Fallo"
            break}
            }