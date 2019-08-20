function Get-FolderPath{
    <#
        .SYNOPSIS
        Permite la obtención de la ruta de alguna carpeta.
                
        .DESCRIPTION
        Permite la obtención de la ruta de alguna carpeta permitiendo generar o imponer una descripción personalizada
        asi como el poder o no, crear una nueva carpeta en el proceso.
                
        .LINK
        Para mayor información por favor verificar 'Get-FolderPath' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-FolderPath

        Solicitará la ruta de una carpeta permitiendo la creación de alguna nueva y utilizando "Encuentra tu carpeta"
        como descripción.
        
        .EXAMPLE
        Get-FolderPath -NewFolder $false

        Solicitará la ruta de una carpeta evitando la creación de alguna nueva y utilizando "Encuentra tu carpeta"
        como descripción.      
                
        .EXAMPLE
        Get-FolderPath -NewFolder $false -Description 'Encuentra la carpeta para el archivo'

        Solicitará la ruta de una carpeta evitando la creación de alguna nueva y utilizando "Encuentra la carpeta 
        para el archivo" como descripción. 

        .PARAMETER NewFolder
        Parametro de tipo boleano que determinado la posibilidad de crear o no, nuevas carpetas en el proceso
        de ubicación de una existente.

        .PARAMETER Description
        Parametro de tipo String que permite establecer una descripción determinada.

    #>
    [cmdletbinding()]
    param(
            [bool]$NewFolder = $true,
            [string]$Description
    )
    process{
        
            if($Description -eq $null){ $Description =  "Encuentra tu carpeta" }
            [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
            $OBJ_IMP_PT = New-Object System.Windows.Forms.FolderBrowserDialog
            $OBJ_IMP_PT.ShowNewFolderButton = $NewFolder
            $OBJ_IMP_PT.Description = $Description
            $Show = $OBJ_IMP_PT.ShowDialog()
            If ($Show -eq "OK")
            {
	            Return $OBJ_IMP_PT.SelectedPath
            }
            else{
                Write-Warning "No seleccionaste ninguna carpeta"
            }
            
    }
}
$OUT_PATE = Get-FolderPath -Description "Ubica la Carpeta donde se guardarán los Logs"
wevtutil el | ForEach {$INT = "$_"; $Out_F = $OUT_PATE + '\' + "$INT"+'.evtx';wevtutil cl "$INT" /bu:"$Out_F" }