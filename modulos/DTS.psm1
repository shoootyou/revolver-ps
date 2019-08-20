#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Copyright (c) 2016 Rodolfo Castelo Méndez. Dos Tercios de Shell
#
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#
#       Versión 1.0.0
#    10 de Junio del 2016
#

function Add-PSCustomLocation{
    <#
        .SYNOPSIS
        Añade la ubicación "C:\Users\%username%\.PowerShell_Functions" a las rutas de carga de PowerShell
        
        .DESCRIPTION
        Crea y añade la ubicación "C:\Users\%username%\.PowerShell_Functions" a las rutas de carga de 
        PowerShell brindando así, una carga automática de las funciones y modulos personalizados existentes que 
        se hayan copiado a dicha ruta.
        
        .LINK
        Para mayor información por favor verificar 'Add-PSCustomLocation' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Add-PSCustomLocation

        Añadirá la ruta "C:\Users\%username%\.PowerShell_Functions" a la carga automática de PowerShell

        .EXAMPLE
        Add-PSCustomLocation -PersonalPath $true

        Preguntará por un valor personalizado de ruta.

        .PARAMETER PersonalPath
        Establece una ruta personalizada para la carga automatica de Powershell

    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$false)]
        [bool]$PersonalPath = $false
    )
    Process{
    $WID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $WPRI=new-object System.Security.Principal.WindowsPrincipal($WID)
    $R_ADM=[System.Security.Principal.WindowsBuiltInRole]::Administrator
    if ($WPRI.IsInRole($R_ADM)){
        $PS_PATH = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        if($PersonalPath -eq $false){
            $PerPath = $env:USERPROFILE + '\.PowerShell_Functions'
            if(!(([Environment]::GetEnvironmentVariable("PSModulePath", "Machine")) -like ('*'+$PerPath+'*'))){
                if(!(Get-Item $PerPath -ErrorAction SilentlyContinue)){
                    New-Item -Path $PerPath -ItemType Directory | Out-Null
                }
                $CurrentValue = (Get-ItemProperty -Path $PS_PATH -Name PSModulePath).PSModulePath
                [string]$New_Path = $CurrentValue +';' + $PerPath
                Set-ItemProperty -Path $PS_PATH -Name PSModulePath -Value $New_Path
                Write-Host
                Write-Host 'Se ha añadido la ruta "'$PerPath '" a tu carga automática de Powershell'
                Write-Host 'Copia tus archivos,módulos y/o funciones de powershell a dicha ruta'
                Write-Host 'Y PowerShell detectará automáticamente las funciones que tengas'
            }
            else{
                Write-Host
                Write-Host 'Ya tienes añadida la ruta "'$PerPath '" a tu carga automática de Powershell'
            }
        }
        elseif($PersonalPath -eq $true){
            [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
            $PE_PATH = New-Object System.Windows.Forms.FolderBrowserDialog 
            $PE_PATH.Description = '¿Qué ruta será la que usarás de repositorio de PowerShell?'
            $SHO = $PE_PATH.ShowDialog()
            if($SHO -eq "OK")
            {
	            $PerPath = $PE_PATH.SelectedPath
                if(!(([Environment]::GetEnvironmentVariable("PSModulePath", "Machine")) -like ('*'+$PerPath+'*'))){
                    if(!(Get-Item $PerPath -ErrorAction SilentlyContinue)){
                        New-Item -Path $PerPath -ItemType Directory | Out-Null
                    }
                    $CurrentValue = (Get-ItemProperty -Path $PS_PATH -Name PSModulePath).PSModulePath
                    [string]$New_Path = $CurrentValue +';' + $PerPath
                    Set-ItemProperty -Path $PS_PATH -Name PSModulePath -Value $New_Path
                    Write-Host
                    Write-Host 'Se ha añadido la ruta "'$PerPath '" a tu carga automática de Powershell'
                    Write-Host 'Copia tus archivos,módulos y/o funciones de powershell a dicha ruta'
                    Write-Host 'Y PowerShell detectará automáticamente las funciones que tengas'
                }
                else{
                    Write-Host
                    Write-Host 'Ya tienes añadida la ruta "'$PerPath '" a tu carga automática de Powershell'
                }
            }
            elseif($SHO -eq "Cancel"){
                Write-Warning "No estableciste una ruta personalizada"
            }
        }
    }
    else{
        Write-Warning "Es Necesario la ejecución como administrador para el uso de ésta función"
    }
    }
}

function Update-DTSModules{
    <#
        .SYNOPSIS
        Actualiza los módulos de PowerShell de DTS
        
        .DESCRIPTION
        Descarga y actualiza los módulos necesarios de DTS en tu computador
        
        .LINK
        Para mayor información por favor verificar 'Update-DTSModule' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Update-DTSModule

        Actualizará los módulos de PS en "C:\Users\%username%\.PowerShell_Functions"
    #>
    [CmdletBinding()]
    param(

    )
    Process{
        $PS_PATH = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        $PerPath = $env:USERPROFILE + '\.PowerShell_Functions'
        $DTSPath = $PerPath + '\DTS'
        if(!(([Environment]::GetEnvironmentVariable("PSModulePath", "Machine")) -like ('*'+$PerPath+'*'))){
            $WID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
            $WPRI=new-object System.Security.Principal.WindowsPrincipal($WID)
            $R_ADM=[System.Security.Principal.WindowsBuiltInRole]::Administrator
            if ($WPRI.IsInRole($R_ADM)){
                if(!(Get-Item $PerPath -ErrorAction SilentlyContinue)){
                    New-Item -Path $PerPath -ItemType Directory | Out-Null
                }
            
                $CurrentValue = (Get-ItemProperty -Path $PS_PATH -Name PSModulePath).PSModulePath
                [string]$New_Path = $CurrentValue +';' + $PerPath
                Set-ItemProperty -Path $PS_PATH -Name PSModulePath -Value $New_Path
            
                Write-Host 'Se está actualizando los módulos de DTS, por favor espera'

                $GIT_DW = "https://github.com/dostshell/dosterciosdeshell/archive/master.zip"
                $TMP_OUT = "$env:USERPROFILE\Desktop\master.zip"
                Invoke-WebRequest -Uri $GIT_DW -OutFile $TMP_OUT

                $TMP_DIR = ($TMP_OUT.Substring(0,$TMP_OUT.LastIndexOf("\"))+'\download')
                if(!(Test-Path $TMP_DIR)){
                    mkdir $TMP_DIR | Out-Null
                }
                $SHELLAPP = new-object -com shell.application
                $COM = $SHELLAPP.NameSpace($TMP_OUT)
                foreach($SUB in $COM.items()){
                    $SHELLAPP.Namespace($TMP_DIR).copyhere($SUB)
                }

                Remove-Item $DTSPath -Recurse -Force -Confirm:$false
                if(!(Test-Path $DTSPath -ErrorAction SilentlyContinue)){
                    Copy-Item $TMP_DIR\dosterciosdeshell-master\DTS\  $DTSPath -Recurse -Force -Confirm:$false
                }
                Remove-Item $TMP_OUT -Recurse -Force -Confirm:$false
                Remove-Item $TMP_DIR -Recurse -Force -Confirm:$false
				Import-Module DTS
            }
            else{
                Write-Warning "Debido a que es la primera actualización se requiere ejecutar como Administrador"
            }
        }
        else{
            if(!(Get-Item $PerPath -ErrorAction SilentlyContinue)){
                New-Item -Path $PerPath -ItemType Directory | Out-Null
            }
            $GIT_DW = "https://github.com/dostshell/dosterciosdeshell/archive/master.zip"
            $TMP_OUT = "$env:USERPROFILE\Desktop\master.zip"
            Invoke-WebRequest -Uri $GIT_DW -OutFile $TMP_OUT

            $TMP_DIR = ($TMP_OUT.Substring(0,$TMP_OUT.LastIndexOf("\"))+'\download')
            if(!(Test-Path $TMP_DIR)){
                mkdir $TMP_DIR | Out-Null
            }
            $SHELLAPP = new-object -com shell.application
            $COM = $SHELLAPP.NameSpace($TMP_OUT)
            foreach($SUB in $COM.items()){
                $SHELLAPP.Namespace($TMP_DIR).copyhere($SUB)
            }

            Remove-Item $DTSPath -Recurse -Force -Confirm:$false
            if(!(Test-Path $DTSPath -ErrorAction SilentlyContinue)){
                Copy-Item $TMP_DIR\dosterciosdeshell-master\DTS\  $DTSPath -Recurse -Force -Confirm:$false
            }
            Remove-Item $TMP_OUT -Recurse -Force -Confirm:$false
            Remove-Item $TMP_DIR -Recurse -Force -Confirm:$false
			Import-Module DTS
        }
    Write-Host "Todos los módulos DTS actualizados, gracias por utilizar."
    Write-Host
    pause
    }
}


