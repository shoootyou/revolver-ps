$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
#############################################################################################################################
Write-Host "================================================================================================"
Write-Host "                                         Enjoy your work                                 "
Write-Host "                                  .                          .                           "
Write-Host "                                  ,c.                       :;                           "
Write-Host "                                   'o:.                  .;o,                            "
Write-Host "                                      .cc,.             'cl'                              "
Write-Host "                            '           'cl'        'cc'           ..                    "
Write-Host "                           .xo            .lkccoollOd.            ;0'                    "
Write-Host "                            ,0.          :OXWMMMMMMMWKl          .0d                     "
Write-Host "                             ld         cNMMMMMMMMMMMMWo         dx.                     "
Write-Host "                              c:       .OMMMMMMMMMMMMMMK.       'c.                      "
Write-Host "                               .c;.     oWMMMMMMMMMMMMWO.    .cl.                        "
Write-Host "                                 'ld;.  .xWMMMMMMMMMMWk.  .cOd'                          "
Write-Host "                                    .ckOc. .'l0WMMMKc,. .:O0o.                            "
Write-Host "                                      .dKO; ,NMMMMNc .o0x,                               "
Write-Host "                                        .c0OoNMMMMMKx0l.                                 "
Write-Host "                                          .lXMMMMMMNo.                                   "
Write-Host "                                           'OWMMMMMK;.                                   "
Write-Host "                                    ..;loxxo0NKNWKNNddxol:..                             "
Write-Host "                                .lk0K0d:. .ox,.KX.'xx. .:oOKKOo.                         "
Write-Host "                               .xd.      :k:  ;NWd  :Oc      .ck.                        "
Write-Host "                              ;kc      ,kx.  .xWWO.  .oO;      ;Oc                       "
Write-Host "                            ,dd.     .dKl..cxONMMWNKx,.;0k.     .lx;.                    "
Write-Host "                         .lxo.      .kx. lNMMMMMMMMMMWx..d0.      .cxo'                  "
Write-Host "                      .,dxx,         .x. ,NMMMMMMMMMMMMN: .d.         ,dxd,.              "
Write-Host "                   .:ddl'             o; lNMMMMMMMMMMMMWd .O.            .cdoc'           "
Write-Host "                ,l:'.                co.:NMMMMMMMMMMMMNl :O.                ';l:         "
Write-Host "                                     lx .KMMMMMMMMMMMMX' oO.                             "
Write-Host "                                    .kx  cNMMMMMMMMMMWd  o0.                             "
Write-Host "                                    .0o   lNMMMMMMMMWk.  ;K.                             "
Write-Host "                                    '0,    ;KMMMMMMNl    .K;                             "
Write-Host "                                    c0.     .:x00kl.     .kx                             "
Write-Host "                                   .Oc                    ,0.                            "
Write-Host "                                   'x.                     o;                            "
Write-Host "                                   ;.                      .:                            "
Write-Host "                                                                                         "
Write-Host "================================================================================================"
Write-Host 
#############################################################################################################################
Write-Host "Cual es el dominio de Office 365 que desea consultar?"
$Empresa = Read-Host 
Write-Host
Write-Host "================================================================================================"
Write-Host "                    Validando y obteniendo Información, por favor espere..."
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

##############################################################################################################################
            Write-Host
            pause
##############################################################################################################################
                Break
                }
        }
}
        if($SIExisteDom -ne 1){
                Write-Host "No posees dicho dominio"

##############################################################################################################################
            Write-Host
            pause
##############################################################################################################################
        }