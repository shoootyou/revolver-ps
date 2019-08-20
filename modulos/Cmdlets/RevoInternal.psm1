﻿<#MIT License

Copyright (c) 2017 Rodolfo Castelo Méndez

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#
# Generated by: Rodolfo Castelo Méndez
#
# Generated on: 10/31/2017
#>

function Write-RevoMessageConsole{
    <#
        .SYNOPSIS
        Escribirá en la consola de acuerdo al tipo de mensaje que se desea traansmitir.
                
        .DESCRIPTION
        Escribirá en la consola de acuerdo al tipo de mensaje que se desea traansmitir.
        De uso interno de RevoModules
                
        .LINK
        Consultar con el creador para más información.

    #>
    [cmdletbinding()]
    param(
            [Parameter(Mandatory=$True,Position=0)]
            [string]
            $Message,
            [Parameter(Mandatory=$True,Position=1)]
            [ValidateSet('Verbose','Warning','Error','Confirmation')]
            [string]
            $Type
    )
    process{
        if($Type -eq 'Warning'){
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
            Write-Host $Message -ForegroundColor Yellow
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Yellow
        }
        elseif($Type -eq 'Error'){
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Red
            Write-Host $Message -ForegroundColor Red
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Red
        }
        elseif($Type -eq 'Verbose'){
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
            Write-Host $Message -ForegroundColor Cyan
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Cyan
        }
        elseif($Type -eq 'Confirmation'){
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Green
            Write-Host $Message -ForegroundColor Green
            Write-Host '-------------------------------------------------------------------------------------------------------------------------------' -ForegroundColor Green
        }
    }
}

Function Test-RevoSubnetInformation{
    <#
        .SYNOPSIS
        Permite validar si una subnet específica forma parte de una red virtual o no.
        
        .DESCRIPTION
        Permite validar si una subnet específica forma parte de una red virtual o no.
                
        .PARAMETER SubnetInformation
        Tipo: Opcional
        Valor que debe poseer la información de la Subnet a validar.

        .PARAMETER VirtualNetworkInformation      
        Tipo: Requerido
        Valor que debe poseer la información de la VirtualNetwork para validar una subnet.

    #>
    [CmdletBinding(DefaultParametersetName='None')] 
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [ValidateNotNullOrEmpty()]
        $VirtualNetworkInformation,
        [ValidateNotNullOrEmpty()]
        [string]
        $SubnetInformation
    )
    process{
        $BEG_VNET = $VirtualNetworkInformation
        $VNETSubnetAddressPrefix = $SubnetInformation

        $BEG_CON_01 = 0
        $BEG_VNET_SUBNET = $null
        do{
            if($BEG_VNET.Subnets[$BEG_CON_01].AddressPrefix -eq $VNETSubnetAddressPrefix){
                $BEG_VNET_SUBNET = $BEG_VNET.Subnets[$BEG_CON_01]
            }
            $BEG_CON_01++
        }
        until($BEG_VNET.Subnets.Count -eq $BEG_CON_01)
    }
    end{
        
        return $BEG_VNET_SUBNET

    }
}

function Confirm-InteractiveEnviroment{
    <#
        .SYNOPSIS
        confirmará que el ámbito en el que se ejecutará PowerShell sea gráfico.
                
        .DESCRIPTION
        confirmará que el ámbito en el que se ejecutará PowerShell sea gráfico.
        De uso interno de RevoModules
                
        .LINK
        Consultar con el creador para más información.

        .EXAMPLE
        Confirm-InteractiveEnviroment

        confirmará que el ámbito en el que se ejecutará PowerShell sea gráfico,
        si en caso no lo es, retornará el valor de $false
    #>
    process{
        $Enviroment = [Environment]::UserInteractive
        return $Enviroment
    }
}