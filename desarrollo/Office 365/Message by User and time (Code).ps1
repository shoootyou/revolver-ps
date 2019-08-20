#########################
#### Valores Básicos ####
#########################

$TotalDelivered = 0
$TotalFilteredAsSpam = 0
$TotalFailed = 0

#########################
####  Tipo de Buzón  ####
#########################

$BuzonPersonal = New-Object System.Management.Automation.Host.ChoiceDescription "&Buzones del Personal", ` "Sí"
$BuzonCompartido = New-Object System.Management.Automation.Host.ChoiceDescription "&Buzones compartidos", ` "No"
$BuzonCompartido = New-Object System.Management.Automation.Host.ChoiceDescription "&Buzones compartidos", ` "No"
$OpcionesBuzon = [System.Management.Automation.Host.ChoiceDescription[]]($BuzonPersonal, $BuzonCompartido)
$ResultadoBuzon = $host.ui.PromptForChoice("Validación", "¿Qué tipos de buzones deseas verificar?", $OpcionesBuzon, 0) 

If($ResultadoBuzon -eq 0){
    $BuzonPersonalIN = "UserMailbox"
}
If($ResultadoBuzon -eq 1){
    $BuzonPersonalIN = "SharedMailbox"
}

#########################

$DireccionesUsuarios = Get-Mailbox -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -like "$VariableConsulta") -and ($_.RecipientTypeDetails -eq "$BuzonPersonalIN")} | Select DisplayName,UserPrincipalName
foreach ($DireccionU in $DireccionesUsuarios){
         $ValidaTipo =  Get-MessageTrace -RecipientAddress $DireccionU.UserPrincipalName -StartDate "05/25/2015 11:00 PM" -EndDate "05/26/2015 2:00 PM" | Select Status
            foreach ($CorreoSobre in $ValidaTipo){
                If($CorreoSobre.Status -eq "Delivered"){
                    $Delivered++
                    $TotalDelivered++
                    }
                elseIf($CorreoSobre.Status -eq "FilteredAsSpam"){
                    $FilteredAsSpam++
                    $TotalFilteredAsSpam++
                    }
                elseIf($CorreoSobre.Status -eq "Failed"){
                    $Failed++
                    $TotalFailed++
                    }
            }
            Write-Host "El usuario" $DireccionU.DisplayName 
            Write-Host "Que tiene de correo" $DireccionU.UserPrincipalName ":"
            If($Delivered -gt 0){
                Write-Host "Tiene" $Delivered "correo(s) marcado(s) como entregado(s)."
                Clear-Variable Delivered 
            }
            If($FilteredAsSpam -gt 0){
                Write-Host "Tiene" $FilteredAsSpam "correo(s) marcado(s) como SPAM."
                Clear-Variable FilteredAsSpam 
            }
            If($Failed -gt 0){
                Write-Host "Tiene" $Failed "correo(s) que no se pudieron entregar."
                Clear-Variable Failed 
            }
            elseif(($Delivered -eq 0) -or ($FilteredAsSpam -eq 0) -or ($Failed -eq 0) -or (!$ValidaTipo)){
                Write-Host "Tiene 0 correo(s) de acuerdo a lo solicitado"
            }
            Write-Host
            Write-Host "================================================================================================"
            Write-Host
}
            Write-Host "De acuerdo al rango de tiempo solicitado:"
            Write-Host
            Write-Host "La empresa tiene " $TotalDelivered "correo(s) marcado(s) como entregado(s) en total." 
            Write-Host "La empresa tiene " $TotalFilteredAsSpam "correo(s) marcado(s) como SPAM en total."
            Write-Host "La empresa tiene " $TotalFailed "correo(s) que no se pudieron entregar en total."
            Write-Host
            Write-Host "================================================================================================"
            
            Clear-Variable TotalDelivered
            Clear-Variable TotalFilteredAsSpam
            Clear-Variable TotalFailed    