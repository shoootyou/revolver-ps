Write-Host '¿Cuál es el nombre o IP del equipo a consultar?'
$PC_NAME = Read-Host
$CON_PC = '\\' + $PC_NAME + '\C$\ProgramData\VeraCrypt\Original System Loader'
$BootLoader = Test-Path $CON_PC -ErrorAction SilentlyContinue
if($BootLoader){
    Write-Host 'El Equipo de encuentra encriptado o se inició ' -ForegroundColor Green
    Write-Host 'un proceso de encriptación en la siguiente fecha:' -ForegroundColor Green
    Write-Host
    Write-Host (Get-Item $CON_PC).LastWriteTime -ForegroundColor Green
}
else{
    Write-Host "El equipo aún no se encuentra encriptado o no se pudo contactactar con el mismo." -ForegroundColor Yellow -BackgroundColor Black
}
pause


function Get-FilePath{
    <#
        .SYNOPSIS
        Permite la obtención de la ruta de algun archivo.
                
        .DESCRIPTION
        Permite la obtención de la ruta de algun archivo, adicional a ello permite filtrar
        por ciertos tipos de archivos.
                
        .LINK
        Para mayor información por favor verificar 'Get-Filepath' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-Filepath

        Solicitará la ruta de un archivo de cualquier extensión comenzando en el perfil del usuario 
        que ejecuta el comando
        
        .EXAMPLE
        Get-Filepath -Title "Selecciona el archivo HTML" -Filter "HTML (HTML files *.html)|*.html"

        Solicitará la ruta de un archivo HTML comenzando en el perfil del usuario que ejecuta el comando      

        .PARAMETER Title
        Parametro de tipo string que permite al usuario especificar el título de la ventana que 
        solicitará el archivo.

        .PARAMETER Filter
        Parametro de tipo String que permite determinar un tipo de archivo en particular que desee.

        .PARAMETER Path
        Parametro de tipo String que nos permite establecer una ruta de inicio en la búsqueda
        del archivo.
    #>
    [cmdletbinding()]
    param(
            [string]$Title,
            [string]$Filter =  "Multiple Files (*.*)|*.*",
            [string]$Path = $env:USERPROFILE
    )
    process{
            if($Title -eq $null){$Title = "Ubica el archivo que desees"}
	        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	        $OBJ_IMP_PT = New-Object System.Windows.Forms.OpenFileDialog
	        $OBJ_IMP_PT.InitialDirectory = $Path
	        $OBJ_IMP_PT.Filter = $Filter
	        $OBJ_IMP_PT.Title = $Title
	        $Show = $OBJ_IMP_PT.ShowDialog()
	        If ($Show -eq "OK")
	        {
		        Return $OBJ_IMP_PT.FileName
	        }
            else{
                Write-Warning "No seleccionaste ningún archivo"
            }
        
        }
    }

$Ext_Pat = Get-FilePath

$DB_Computers = Import-Csv -Path $Ext_Pat

$Out_Veracrypt = @()
$i = 1
foreach($Computer in $DB_Computers){
    $ObjProperties = New-Object PSObject

    $PC_NAME = $Computer.ComputerName
    $CON_PC = '\\' + $PC_NAME + '\C$\ProgramData\VeraCrypt\Original System Loader'
    $BootLoader = Test-Path $CON_PC -ErrorAction SilentlyContinue
    if($BootLoader){
        $ENC_DAT = (Get-Item $CON_PC).LastWriteTime

        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "ComputerName" -Value $PC_NAME
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Status" -Value "Encriptado o Encriptando"
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Fecha de inicio" -Value $ENC_DAT
        $Out_Veracrypt += $ObjProperties
    }
    else{
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "ComputerName" -Value $PC_NAME
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Status" -Value "No Encriptada"
        Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Fecha de inicio" -Value "-"
        $Out_Veracrypt += $ObjProperties
    }

    Write-Progress -Activity “Gathering Information” -status “Working on $PC_NAME” -percentComplete ($i / $DB_Computers.count*100)
    $i++

}
$Out_Veracrypt | Out-GridView -Title "Reporte de Veracrypt"
$Out_Veracrypt | Export-Csv 