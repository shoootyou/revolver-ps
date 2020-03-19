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
#    05 de Septiembre del 2016
#
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#        Módulo de caracter Interno / Internal Use
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

function Confirm-InteractiveEnviroment{
    <#
        .SYNOPSIS
        confirmará que el ámbito en el que se ejecutará PowerShell sea gráfico.
                
        .DESCRIPTION
        confirmará que el ámbito en el que se ejecutará PowerShell sea gráfico.
        De uso interno de DTS
                
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

function ConvertFrom-HTMLtoMail{
    <#
        .SYNOPSIS
        Convierte un archivo HTML a un HTML sin altos de línea.
                
        .DESCRIPTION
        Convierte un archivo HTML a un HTML sin altos de línea utilizado para el envío de correos masivos.
                
        .LINK
        Para mayor información por favor verificar 'Convert-HTMLtoMail' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        ConvertFrom-HTMLtoMail

        Preguntará por la ruta del archivo.

        .PARAMETER Path
        Parametro de tipo String que nos permite establecer una ruta del archivo a convertir.
    #>
    [cmdletbinding()]
    param(
            [Parameter(Position=0)]
            [string]$Path= ''
    )
    begin{
        Function Conv3rt{
            $HTML_OUT = ""
            $INT_CONT = 0
            $HTML_IO  = [System.IO.File]::OpenText("$Path")
            $HTML_CNE = Get-Content -Path $Path
            do{
                $HTML_OUT += $HTML_IO.ReadLine()
                $INT_CONT++
            }
            until($INT_CONT -eq $HTML_CNE.Count)
            $HTML_IO.Close()
            return $HTML_OUT
        }
    }
    process{
        if($Path){
           Conv3rt 
        }
        else{
            if(Confirm-InteractiveEnviroment){
                $Path = Get-FilePath -Filter 'Archivo HTHML (*.html)|*.html' -WarningAction SilentlyContinue
                Conv3rt
            }
            else{
                Write-Host '¿Cuál es la ubicación del archivo?'
                $Path = Read-Host
                Conv3rt
            }
        }
    }
}

function Test-CSVHeader{
    <#
        .SYNOPSIS
        Verifica las cabeceras de los CSV Importados.
                
        .DESCRIPTION
        Verifica las cabeceras de los CSV Importados respecto a la cabecera que debería tener.
                
        .LINK
        Para mayor información por favor verificar 'Test-CSVHeader' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Test-CSVHeader -ImportedCSV $TEST_CSV -TestValue 'Prueba'

        Verificará si el valor "Prueba" formar parte de las cabeceras del CSV importado en la variable $TEST_CSV

        .PARAMETER ImportedCSV
        Parametro que permite la alojación de otras variables en las que ha sido guardado un CSV Importado.

        .PARAMETER TestValue
        Parametro de tipo String que solicita la cabecera a testear. Valor único.
    #>
    [cmdletbinding()]
    param(
            [Parameter(Position=0,Mandatory=$true)]
            $ImportedCSV,
            [Parameter(Position=1,Mandatory=$true)]
            [string]$TestValue
    )
    begin{
        $Test_ORI = $ImportedCSV
        $Test_REQ = $TestValue
        $RTN_VAL = $null
    }
    process{
        $CHK_INT = $Test_ORI | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name' -ErrorAction SilentlyContinue
        if($CHK_INT.Count -ge 2){
            $DO_INT_PRO = 0
            do{
                if($CHK_INT[$DO_INT_PRO] -like $TestValue){
                    $RTN_VAL = $true
                }
                $DO_INT_PRO++
            }
            until($DO_INT_PRO -eq $CHK_INT.Count)
        }
        else{
            if($CHK_INT -like $TestValue){
                $RTN_VAL = $true
            }
        }
        if($RTN_VAL -eq $null){$RTN_VAL = $false}
    }
    end{
        return $RTN_VAL
    }
}