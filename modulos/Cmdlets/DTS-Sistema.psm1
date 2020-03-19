#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Copyright (c) 2016 Rodolfo Castelo Méndez. Dos Tercios de Shell
#
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#
#       Versión 1.1.1
#    16 de Junio del 2016
#

Function Get-OSBuild{
    <#
        .SYNOPSIS
        Permite ubicar y encontrar la versión de build del SO
        
        .DESCRIPTION
        Permite ubicar y encontrar la versión de build del SO devolviendo el valor como
        un string. Se permite la opción de brindar la build en su formato corto de identificación, 
        o en su formato largo, asimismo brinda la posibilidad de obtener el listado de Builds completa
        para poder saber qué build pertenece a qué Sitema Operativo.
        
        .LINK
        Para mayor información por favor verificar 'Get-OSBuild' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-OSBuild

        Retorna el valor de la build en su versión corta, esta versión son únicamente el primer digito seguido de
        un punto y el detalle diferenciador del segundo dígito de versión
        
        .EXAMPLE
        Get-OSBuild -FullDetails $true

        Retorna el valor de la build en su versión larga, tal y como la brinda el SO, con el detalle de build.
        
        .EXAMPLE
        Get-OSBuild -ListBuilds $true

        Retorna el listado de Builds para todos los sistemas operativos Windows existentes a la fecha y de los cuales
        se tiene un registro de Build.        

        .PARAMETER FullDetails
        Parametro de tipo string orientado al establecimiento del usuario.

        .PARAMETER ListBuilds
        Parametro de tipo String que permite la inserción de la clave de forma explícita.
    #>
    [cmdletbinding()]
    param(
        [bool]$FullDetails = $false,
        [bool]$ListBuilds = $false
    )
    begin{

    }
    process{
        if(($FullDetails -eq $false) -and ($ListBuilds -eq $false)){
            $DB_OS = Get-WmiObject -Class Win32_OperatingSystem | Select Version,ServicePackMajorVersion,Caption
            $OS_VR = $DB_OS.Version
            $OUT_01 =  $OS_VR.Substring(0,$OS_VR.IndexOf(".")+1)
            $TMP_01 = $OS_VR.Substring($OS_VR.IndexOf(".")+1)
            $POS_01 = $TMP_01.IndexOf(".")
            $OUT_02  = $TMP_01.Substring(0,$POS_01)
            $OUT_TMP = $OUT_01 + $OUT_02

            $INT_OBJ = New-Object PSObject      
            Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Windows' -Value $DB_OS.Caption
            Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Build' -Value $OUT_TMP
            return $INT_OBJ
        }
        elseif(($FullDetails -eq $true) -and ($ListBuilds -eq $true)){
            Write-Warning 'Parámetros incorrectos'
        }
        elseif($FullDetails -eq $true){
            $DB_OS = Get-WmiObject -Class Win32_OperatingSystem | Select Version,ServicePackMajorVersion,Caption
            $OUT_01 = $DB_OS.Version
            if($DB_OS.ServicePackMajorVersion -eq 0){
                $INT_OBJ = New-Object PSObject      
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Windows' -Value $DB_OS.Caption
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Build' -Value $OUT_01
                return $INT_OBJ
            }
            elseif($DB_OS.ServicePackMajorVersion -gt 0){
                $OUT_02 = $DB_OS.ServicePackMajorVersion
                $OUT_TMP = $OUT_01 + ' SP ' + $OUT_02

                $INT_OBJ = New-Object PSObject      
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Windows' -Value $DB_OS.Caption
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Build' -Value $OUT_01
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'SP' -Value $OUT_TMP
                return $INT_OBJ
            }
        }
        elseif($ListBuilds -eq $true){
            $OUT_OBJ = @()
            $DB_OS = @(
                ('Windows Server 2016 Technical Preview','10.0*'),("Windows 10",'10.0*'),
                ('Windows Server 2012 R2','6.3*'),('Windows 8.1','6.3*'),
                ('Windows Server 2012','6.2'),('Windows 8','6.2'),
                ('Windows Server 2008 R2','6.1'),('Windows 7','6.1'),
                ('Windows Server 2008','6.0'),('Windows Vista','6.0')
            )
            $BRK_UNT = 0
            do{
                $INT_OBJ = New-Object PSObject      
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Windows' -Value $DB_OS[$BRK_UNT][0]
                Add-Member -InputObject $INT_OBJ -MemberType NoteProperty -Name 'Build' -Value $DB_OS[$BRK_UNT][1]
                $OUT_OBJ += $INT_OBJ
                $BRK_UNT++
            }
            until($DB_OS.Count -eq $BRK_UNT)
            return $OUT_OBJ
        }
    }
}