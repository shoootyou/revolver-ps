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

function Get-PROArchitecture{
    <#
        .SYNOPSIS
        Define la arquitectura de tu procesador y retorna un valor de 64, 32 o 0
        
        .DESCRIPTION
        Define la arquitectura de tu procesador y retorna un valor de 64 para procesadores
        de 64-bits, 32 para procesadores de 32-bits o 0 en caso de no poder determinar
        la arquitectura de tu procesador.
        
        .LINK
        Para mayor información por favor verificar 'Get-PROArchitecture' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-PROArchitecture

        Retorna la arquitectura del procesador en valores de tipo Int64 ya sea 64, 32 o 0
        
        .EXAMPLE
        Get-PROArchitecture -ReturnInString $true

        Retorna la arquitectura del procesador en valores de tipo String ya sea 64, 32 o 0
        
        .EXAMPLE
        Get-PROArchitecture -ReturnInString $true -LongDescription $true

        Retorna la arquitectura del procesador en valores de tipo String de la siguiente forma:
            - Para procesadores de 64 bits retornará el valor: 64-bits
            - Para procesadores de 32 bits retornará el valor: 32-bits
            - Si en caso no se pudo determinar el procesador : No es posible determinar

        .PARAMETER ReturnInString
        Parametro booleano que permite el retorno en tipo String.

        .PARAMETER LongDescription
        Parametro booleano que permite el retorno en tipo String de forma extendida.

    #>
    param(
        [parameter(Mandatory=$false)]
        [alias("InString","EnTexto")]
        [bool]$ReturnInString = $false,
        [parameter(Mandatory=$false)]
        [bool]$LongDescription = $false
    )
    Process{
        if(($ReturnInString -eq $false) -and ($LongDescription -eq $false)){
            if($env:PROCESSOR_ARCHITECTURE -like '*64*'){
                [int64]$Proc_Ver = 64
                return $Proc_Ver
            }
            elseif(($env:PROCESSOR_ARCHITECTURE -like '*86*') -or ($env:PROCESSOR_ARCHITECTURE -like '*32*')){
                [int64]$Proc_Ver = 32
                return $Proc_Ver
            }
            else{
                [int64]$Proc_Ver = 0
                return $Proc_Ver
            }
        }
        elseif(($ReturnInString -eq $true) -and ($LongDescription -eq $false)){
            if($env:PROCESSOR_ARCHITECTURE -like '*64*'){
                [string]$Proc_Ver = '64'
                return $Proc_Ver
            }
            elseif(($env:PROCESSOR_ARCHITECTURE -like '*86*') -or ($env:PROCESSOR_ARCHITECTURE -like '*32*')){
                [string]$Proc_Ver = '32'
                return $Proc_Ver
            }
            else{
                [string]$Proc_Ver = '0'
                return $Proc_Ver
            }
        }
        elseif(($ReturnInString -eq $true) -and ($LongDescription -eq $true)){
            if($env:PROCESSOR_ARCHITECTURE -like '*64*'){
                [string]$Proc_Ver = '64-bits'
                return $Proc_Ver
            }
            elseif(($env:PROCESSOR_ARCHITECTURE -like '*86*') -or ($env:PROCESSOR_ARCHITECTURE -like '*32*')){
                [string]$Proc_Ver = '32-bits'
                return $Proc_Ver
            }
            else{
                [string]$Proc_Ver = 'No es posible determinar'
                return $Proc_Ver
            }
        }
        elseif(($ReturnInString -eq $false) -and ($LongDescription -eq $true)){
                Write-Warning 'No existe descripción larga para tipos Int64, por favor pruebe sin el parametro LongDescription o pruebe de la siguiente forma: 
                
                Get-PROArchitecture -ReturnInString $false -LongDescription $false
                '
        }
    }
}