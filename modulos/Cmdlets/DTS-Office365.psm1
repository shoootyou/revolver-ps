#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Copyright (c) 2016 Rodolfo Castelo Méndez. Dos Tercios de Shell
#
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#
#       Versión 1.1.0
#    16 de Junio del 2016
#

function Connect-O365Services{
    <#
        .SYNOPSIS
        Permite la conexión a los servicios de Exchange Online.
        
        .DESCRIPTION
        Permite la conexión a los servicios de Exchange Online a través de una sesión de PowerShell
        
        .LINK
        Para mayor información por favor verificar 'Connect-O365Services' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Connect-O365Services -Username ejemplo@dominio.onmicrosoft.com -Password ClaveAqui

        Establece una sesión con los servicios de Office 365 con el usuario ejemplo@dominio.onmicrosoft.com
        que tiene de clave "ClaveAqui"
        
        .EXAMPLE
        Connect-O365Services -Username ejemplo@dominio.onmicrosoft.com -UseHidePassword $true

        Solicita la clave como texto seguro para prevenir la visualización de la misma, posterior a ello
        establece la sesión don dichas credenciales.        

        .PARAMETER Username
        Parametro de tipo string orientado al establecimiento del usuario.

        .PARAMETER Password
        Parametro de tipo String que permite la inserción de la clave de forma explícita.

        .PARAMETER UseHidePassword
        Parametro booleano que habilita el uso de una clave como texto seguro.
    #>
    [cmdletbinding(
        DefaultParameterSetName='Hidden'
    )]
    param(
        
        [Parameter(ParameterSetName='Showing',Mandatory=$true,Position=0)]
        [Parameter(ParameterSetName='Hidden',Mandatory=$true,Position=0)] 
        [ValidateNotNullOrEmpty()]
        [string]$Username,
        [Parameter(ParameterSetName='Showing',Mandatory=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$Password,
        [Parameter(ParameterSetName='Hidden',Position=1,
        HelpMessage='Introduce los Valores $True o $False')]
        [ValidateNotNullOrEmpty()]
        [bool]$UseHidePassword = $true
    )
    begin{
        if($Username -notlike '*@*'){
            Write-Warning 'El usuario no cumple el formato requerido de usuario de Office 365'
            break
        }
        else{
            $GBL_Username = $Username
        }
        if($Password){
            $GBL_Password = ConvertTo-SecureString –String $Password –AsPlainText -Force
        }
        elseif($UseHidePassword -eq $false){
            Write-Warning 'Por favor, proporcione la clave mediante el parametro -Password <clave> o -UserHidePassword $true'
            break
        }
        else{
            Write-Host 'Proporciona tu clave por favor'
            $GBL_Password = Read-Host -AsSecureString
        }
        if(!$GBL_Password){
            Write-Warning 'No se ha proporcionado una contraseña'
        }
    }
    Process{
        ''
        Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
        Write-Host '               Verificando y cerrando sesiones previamente establecidas               ' -foregroundcolor Cyan
        Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
        $DB_SSN = Get-PSSession
        $CN_SSN = 0
        foreach($SSN in $DB_SSN){
            if(($SSN.ComputerName -like '*office365*') -and ($SSN.State -eq 'Opened')){
                Remove-PSSession $SSN
            }
            $CN_SSN++
        }
        if($CN_SSN -eq 0){
            Write-Host '                                           #'
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            Write-Host '         No se han encontrado sesiones de Office365 establecidas previamente          ' -foregroundcolor Green
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
        }
        else{
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            Write-Host '                        Se cerraron ' $CN_SSN ' sesion(es) existente(s                         '  -foregroundcolor Green
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
        }
        Write-Host '                                           #'
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
            Write-Host '             Verificando credenciales e iniciando sesión en Office 365                ' -foregroundcolor Cyan
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
        $GBL_Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $GBL_Username, $GBL_Password
        $GBL_USR_SSN = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $GBL_Credential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
        if($GBL_USR_SSN){
            return Import-PSSession $GBL_USR_SSN -Verbose -DisableNameChecking -ErrorAction SilentlyContinue
        }
        else{
            Write-Host '                                           #'
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
            Write-Host '        Datos proporcionados erróneos. No se pudo iniciar sesión en Office 365        ' -ForegroundColor Yellow 
            Write-Host '                         Revisa la información proporcionada.                         ' -ForegroundColor Yellow
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
        }
           
   }
}