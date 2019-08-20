#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Copyright (c) 2016 Rodolfo Castelo Méndez. Dos Tercios de Shell
#
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#
#       Versión 2.1.0
#    8 de Septiembre del 2016
#

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
        if(Confirm-InteractiveEnviroment){
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
        else{
            if($Title -eq $null){$Title = "Proporciona la ruta completa del archivo que desees"}
            Write-Host $Title
            $INT_PAT = Read-Host
            $TST_PAT = Test-Path -path $INT_PAT -ErrorAction SilentlyContinue
	        if(!$INT_PAT){
				Write-Warning "No seleccionaste ningún archivo"
				break
			}
			if($TST_PAT){
                Return $INT_PAT
	        }
            else{
                Write-Warning "La ruta proporcionada no es correcta"
            }
        }
    }
}

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
        if(Confirm-InteractiveEnviroment){
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
        else{
            if($Description -eq $null){ $Description =  'Proporciona la ruta de la carpeta que necesitas' }
            Write-Host $Description
            $INT_PAT = Read-Host
            $TST_PAT = Test-Path -path $INT_PAT -ErrorAction SilentlyContinue
	        if(!$INT_PAT){
				Write-Warning "No seleccionaste ninguna carpeta"
				break
			}
			if($TST_PAT){
                Return $INT_PAT
	        }
            else{
                Write-Warning "La ruta proporcionada no es correcta"
            }
        }
    


    }
}

function Get-SelectedTime{
    <#
        .SYNOPSIS
        Permite la obtención de una fecha y/u hora determinada mediante una interfaz visual.
                
        .DESCRIPTION
        Permite la obtención de una fecha y/u hora determinada mediante una interfaz visual.
                
        .LINK
        Para mayor información por favor verificar 'Get-SelectedTime' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-SelectedTime

        Solicitará establecer una hora y fecha determinada.
        
        .EXAMPLE
        Get-SelectedTime -Description "Aqui tu texto"

        Solicitará establecer una hora y fecha determinada con el anuncio de "Aqui tu texto" en el cuadro.      
                
        .EXAMPLE
        Get-SelectedTime -OnlyDate $true

        Solicitará establecer una fecha determinada. La hora que se utilizará será la que se tiene actualmente
        en el computador

        .EXAMPLE
        Get-SelectedTime -OnlyTime $true

        Solicitará establecer una hora determinada. La fecha que se utilizará será la que se tiene actualmente
        en el computador

        .PARAMETER FullDetails
        Parametro de tipo boleano que determina si se desea o no la hora y el día en conjunto.

        .PARAMETER OnlyDate
        Parametro de tipo boleano que determina si se desea o no sólo la fecha.

        .PARAMETER OnlyTime
        Parametro de tipo boleano que determina si se desea o no sólo la hora.

        .PARAMETER Description
        Parametro de tipo String que permite establecer una descripción para guiar al con el fin de esta invocación.
    #>
    [cmdletbinding(
        DefaultParameterSetName='Full'
    )]
    param(
        [Parameter(Position=0,ParameterSetName='Full')]
        [bool]$FullDetails = $true,
        [Parameter(Position=0,ParameterSetName='Date')]
        [bool]$OnlyDate,
        [Parameter(Position=0,ParameterSetName='Time')]
        [bool]$OnlyTime,
        [string]$Description = 'Selecciona la fecha y hora que deseas'
        
    )
    begin{
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    process{
        if(($FullDetails -eq $false) -and (!$OnlyDate) -and (!$OnlyTime)){
            Write-Warning 'No se seleccionó un modo de trabajo'
            break
        }
        $GBL_FORM = New-Object Windows.Forms.Form 

        $GBL_FORM.Size = New-Object Drawing.Size @(250,250) 
        $GBL_FORM.StartPosition = "CenterScreen"
        $GBL_FORM.MaximumSize = New-Object Drawing.Size @(250,250) 
        $GBL_FORM.MinimumSize = New-Object Drawing.Size @(250,250) 
        $GBL_FORM.ShowIcon = $false

        $GBL_LBL= New-Object System.Windows.Forms.Label
        $GBL_LBL.Location = New-Object System.Drawing.Point(0, 5)
        $GBL_LBL.Size = New-Object System.Drawing.Point(240, 15)
        $GBL_LBL.TextAlign = "MiddleCenter"
        $GBL_LBL.Margin = 0
        
        $PCK_DT_TM = New-Object System.Windows.Forms.DateTimePicker 
        $PCK_DT_TM.Location = New-Object System.Drawing.Point(0,25)
        $PCK_DT_TM.Size = New-Object System.Drawing.Point(240, 15)
        $PCK_DT_TM.MinDate = '1/1/1900'
        $PCK_DT_TM.MaxDate = '12/31/2099'
        $PCK_DT_TM.ShowCheckBox = $false
        if($Description -ne 'Selecciona la fecha y hora que deseas'){
            $GBL_LBL.Text = $Description
            $GBL_FORM.Controls.Add($GBL_LBL)
        }
        else{
            if($OnlyDate -eq $true){
                $GBL_LBL.Text = 'Selecciona el día que deseas'
                $GBL_FORM.Controls.Add($GBL_LBL)
            }
            elseif($OnlyTime -eq $true){
                $PCK_DT_TM.Format = 'Time'
                $PCK_DT_TM.ShowUpDown = $True
                $GBL_LBL.Text = 'Selecciona la hora que deseas'
                $GBL_FORM.Controls.Add($GBL_LBL)
            }
            else{
                $PCK_DT_TM.Format = 'Time'
                $GBL_LBL.Text = $Description
                $GBL_FORM.Controls.Add($GBL_LBL)
            }
        }
        $PCK_DT_TM.DropDownAlign = 'Left'

        $GBL_FORM.Controls.Add($PCK_DT_TM) 

        $GBL_OK_BTN = New-Object System.Windows.Forms.Button
        $GBL_OK_BTN.Location = New-Object System.Drawing.Point(38,175)
        $GBL_OK_BTN.Size = New-Object System.Drawing.Size(75,23)
        $GBL_OK_BTN.Text = "OK"
        $GBL_OK_BTN.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $GBL_FORM.AcceptButton = $GBL_OK_BTN
        $GBL_FORM.Controls.Add($GBL_OK_BTN)

        $GBL_CN_BTN = New-Object System.Windows.Forms.Button
        $GBL_CN_BTN.Location = New-Object System.Drawing.Point(113,175)
        $GBL_CN_BTN.Size = New-Object System.Drawing.Size(75,23)
        $GBL_CN_BTN.Text = "Cancel"
        $GBL_CN_BTN.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $GBL_FORM.CancelButton = $GBL_CN_BTN
        $GBL_FORM.Controls.Add($GBL_CN_BTN)

        $GBL_FORM.Topmost = $True

        $GBL_RESULT = $GBL_FORM.ShowDialog() 

        if ($GBL_RESULT -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $OUTPUT = $PCK_DT_TM.Value
        }
        else{
            Write-Warning 'No se seleccionó un valor de fecha y/u hora determinado'
            break
        }
        return $OUTPUT
    }
}

function Get-TextBox{
<#
        .SYNOPSIS
        Permite la obtención de un determinado texto mediante una interfaz visual.
                
        .DESCRIPTION
        Permite la obtención de un determinado texto mediante una interfaz visual adicionando una descripción
        para guiar al usuario al momento de uso. Por defecto, soporta la inserción de texto con saltos de
        línea entre él.
                
        .LINK
        Para mayor información por favor verificar 'Get-TextBox' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-TextBox

        Solicitará un texto mediante la ventana gráfica de dimensiones 150 x 150
        
        .EXAMPLE
        Get-TextBox -Height 200

        Solicitará un texto mediante la ventana gráfica de dimensiones 150 x 200
        
        .EXAMPLE
        Get-TextBox -Width 200

        Solicitará un texto mediante la ventana gráfica de dimensiones 200 x 150
        
        .EXAMPLE
        Get-TextBox -Description 'Aquí la descripción'

        Solicitará un texto mediante la ventana gráfica de dimensiones 150 x 150 en la cual aparecerá un texto
         'Aquí la descripción' como etiquera.
        
        .EXAMPLE
        Get-TextBox -BasicText $true

        Solicitará un texto mediante la ventana gráfica de dimensiones 150 x 85 en la cual se
        permitirá sólo el uso de una línea de texto. No habrán saltos de línea.

        .PARAMETER Height
        Parametro de tipo int32 que permite al usuario especificar el tamaño de la altura de la ventana. 
        El valor como mínimo es 150

        .PARAMETER Width
        Parametro de tipo int32 que permite al usuario especificar el tamaño de la anchura de la ventana. 
        El valor como mínimo es 150

        .PARAMETER Description
        Parametro de tipo String que nos permite establecer la descripción que aparecerá en el cuadro.
        
        .PARAMETER BasicText
        Parametro de tipo booleano que permite el activar o no la propiedad de texto básico. 
        Orientado a la recoleción de texto de una sola línea.

        .PARAMETER Title
        Parametro de tipo booleano que permite el activar o no la propiedad de texto básico. 
        Orientado a la recoleción de texto de una sola línea.
    #>
    [cmdletbinding(
        DefaultParameterSetName='Description'
    )]
    param(
            [Parameter(Position=0,ParameterSetName='Description')]
            [int32]$Height,
            [Parameter(Position=1,ParameterSetName='Description')]
            [Parameter(Position=1,ParameterSetName='Basic')]
            [int32]$Width,
            [Parameter(Position=2,ParameterSetName='Description')]
            [String]$Description,
            [Parameter(Position=2,ParameterSetName='Basic')]
            [bool]$BasicText = $false,
            [Parameter(Position=3,ParameterSetName='Description')]
            [Parameter(Position=3,ParameterSetName='Basic')]
            [string]$Title = ''
    )
    begin{
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    process{
        if(!$Height){
            $Height = 150
        }
        elseif($Height -lt 150){
            Write-Warning 'Tamaño no válido. Utilice mínimo 150 de altura'
            break
        }
        if(!$Width){
            $Width = 200
        }
        elseif($Width -lt 150){
            Write-Warning 'Tamaño no válido. Utilice mínimo 150 de anchura'
            break
        }

        $GBL_FORM = New-Object Windows.Forms.Form 
        if($Title){
            $GBL_FORM.Text = $Title
        }
        if($Description){
            $GBL_FORM.Size = New-Object Drawing.Size @($Width,($Height + 50)) 
            $GBL_FORM.MaximumSize = New-Object Drawing.Size @($Width,($Height + 50))
            $GBL_FORM.MinimumSize = New-Object Drawing.Size @($Width,($Height + 50))
        }
        elseif($BasicText){
            $GBL_FORM.Size = New-Object Drawing.Size @($Width,85) 
            $GBL_FORM.MaximumSize = New-Object Drawing.Size @($Width,85)
            $GBL_FORM.MinimumSize = New-Object Drawing.Size @($Width,85)
        }
        else{
            $GBL_FORM.Size = New-Object Drawing.Size @($Width,$Height) 
            $GBL_FORM.MaximumSize = New-Object Drawing.Size @($Width,$Height)
            $GBL_FORM.MinimumSize = New-Object Drawing.Size @($Width,$Height)
        }
        $GBL_FORM.StartPosition = "CenterScreen"
        $GBL_FORM.ShowIcon = $false

        if($BasicText){
            $GBL_TXT_BOX = New-Object System.Windows.Forms.TextBox
            $GBL_TXT_BOX.Size = New-Object System.Drawing.Size ($Width,20)
            $GBL_TXT_BOX.MaximumSize = New-Object Drawing.Size ($Width,20)
        }
        else{
            $GBL_TXT_BOX = New-Object System.Windows.Forms.RichTextBox
            $GBL_TXT_BOX.Size = New-Object System.Drawing.Size ($Width,($Height/2.5))
            $GBL_TXT_BOX.MaximumSize = New-Object Drawing.Size ($Width,($Height/2.5))  
        }
        
        $GBL_TXT_BOX.MinimumSize = New-Object Drawing.Size ($Width,20)
        $GBL_TXT_BOX.AutoSize = $false
        $GBL_TXT_BOX.AcceptsTab = $true

        if($Description){
            $GBL_TXT_BOX.Location = New-Object System.Drawing.Size(0,50)

            $GBL_LBL= New-Object System.Windows.Forms.Label
            $GBL_LBL.Location = New-Object System.Drawing.Point(0, 5)
            $GBL_LBL.Size = New-Object System.Drawing.Point($Width, 45)
            $GBL_LBL.TextAlign = "MiddleCenter"
            $GBL_LBL.Margin = 0
            $GBL_LBL.Text = $Description
            $GBL_FORM.Controls.Add($GBL_LBL)
        }
        elseif($BasicText){
            $GBL_TXT_BOX.Location = New-Object System.Drawing.Size(0,0)
        }
        else{
            $GBL_TXT_BOX.Location = New-Object System.Drawing.Size(0,0)
        }

        $GBL_FORM.Controls.Add($GBL_TXT_BOX) 

        $GBL_OK_BTN = New-Object System.Windows.Forms.Button
        $GBL_OK_BTN.Size = New-Object System.Drawing.Size(50,25)
        $GBL_OK_BTN.Text = "OK"
        $GBL_OK_BTN.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $GBL_FORM.AcceptButton = $GBL_OK_BTN
        
        $GBL_CN_BTN = New-Object System.Windows.Forms.Button
        $GBL_CN_BTN.Size = New-Object System.Drawing.Size(50,25)
        $GBL_CN_BTN.Text = "Cancel"
        $GBL_CN_BTN.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $GBL_FORM.CancelButton = $GBL_CN_BTN

        if($Description){
            $GBL_OK_BTN.Location = New-Object System.Drawing.Point(($Width/2 - 60),($Height/2.5 + 65))
            $GBL_CN_BTN.Location = New-Object System.Drawing.Point(($Width/2 + 10),($Height/2.5 + 65))
        }
        elseif($BasicText){
            $GBL_OK_BTN.Location = New-Object System.Drawing.Point(($Width/2 - 60),20)
            $GBL_CN_BTN.Location = New-Object System.Drawing.Point(($Width/2 + 10),20)
        }
        else{
            $GBL_OK_BTN.Location = New-Object System.Drawing.Point(($Width/2 - 60),($Height/2.5 + 15))
            $GBL_CN_BTN.Location = New-Object System.Drawing.Point(($Width/2 + 10),($Height/2.5 + 15))
        }

        $GBL_FORM.Controls.Add($GBL_OK_BTN)
        $GBL_FORM.Controls.Add($GBL_CN_BTN)

        $GBL_RESULT = $GBL_FORM.ShowDialog() 

        if ($GBL_RESULT -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $GBL_TXT_BOX.Text
        }
    }


}

function Send-BulkMail{
    <#
        .SYNOPSIS
        Permite el envío masivo de correos mediante la carga de archivos con los datos necesarios.
                
        .DESCRIPTION
        Permite el envío masivo de correos mediante la carga de archivos con los datos necesarios,
        actualmente sólo soporta Office 365 para el envío.
                
        .LINK
        Para mayor información por favor verificar 'Send-BulkMail' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Send-BulkMail

        Iniciará el proceso de envío, solicitando lo necesario.
        
        .EXAMPLE
        Send-BulkMail -Subject 'Aquí va el asunto'

        Iniciará el proceso de envío, solicitando lo necesario, con el asunto 'Aquí va el asunto'
        
        .EXAMPLE
        Send-BulkMail -Attach $true

        Iniciará el proceso de envío, solicitando lo necesario y preguntará por un PDF a adjuntar.
        
        .EXAMPLE
        Send-BulkMail -ScriptVersion $true -Attach $true

        Iniciará el proceso de envío asumiento los siguientes valores:
        Cuentas que enviarán:        Ruta proporcionada\senders.csv
        Cuenta a enviar el mail:     Ruta proporcionada\clients.csv
        Mensaje a enviar             Ruta proporcionada\mail.html
        Documento a atachar:         Ruta proporcionada\(preguntará por el nombre).pdf
    #>
    [cmdletbinding(
        DefaultParameterSetName='Visual'
    )]
    param(
        [Parameter(Position=0,Mandatory=$false,ParameterSetName='Script')]
        [bool]$ScriptVersion = $false,
        [Parameter(Position=0,Mandatory=$false,ParameterSetName='Visual')]
        [Parameter(Position=1,Mandatory=$false,ParameterSetName='Script')]
        [bool]$Attach = $false,
        [Parameter(Position=1,Mandatory=$false,ParameterSetName='Visual')]
        [string]$Subject = ''
        
    )
    begin{
        if(!(Confirm-InteractiveEnviroment) -and !($ScriptVersion)){
            Write-Host '----------------------------------------------------------------------------' -ForegroundColor Yellow -BackgroundColor Black
            Write-Host '      Utiliza el parametro -ScriptVersion para éste modo de PowerShell      ' -ForegroundColor Yellow -BackgroundColor Black
            Write-Host '----------------------------------------------------------------------------' -ForegroundColor Yellow -BackgroundColor Black
            Break
        }
        function p4th{
            if($ScriptVersion){$p4th =  Get-FolderPath -Description 'Proporciona la ruta donde se almacenarán los registros y donde se encuentran los archivos necesarios'-WarningAction SilentlyContinue}
            else{$p4th =  Get-FolderPath -Description '¿Cuál es la carpeta donde se encuentran los archivos?'-WarningAction SilentlyContinue}        
            if(!$p4th){
                $p4th = $env:USERPROFILE+'\Desktop'
            }
            return $p4th
        }
        $GBL_SC_PATH = p4th
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Validating path
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        do{
            Write-Host '----------------------------------------------------------' -ForegroundColor Green
            Write-Host '|   Mostrando los archivos de la carpeta proporcionada   |' -ForegroundColor Green
            Write-Host '----------------------------------------------------------' -ForegroundColor Green
            Get-ChildItem $GBL_SC_PATH
            $SW1_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
            $SW1_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
            $SW1_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW1_1ST, $SW1_2ND)
            $SW1_ASW = $host.ui.PromptForChoice("Aditional information", "¿Se encuentra correcta la ruta?", $SW1_OPT, 0) 
            switch ($SW1_ASW){
                0{   $PTH_COR = $true}
                1{   $GBL_SC_PATH = p4th;$PTH_COR = $false}
            }
        }
        until($PTH_COR -eq $true)
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Import required files
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
            if(!$ScriptVersion){
                    $DB_ACC_CSV_PT = Get-FilePath -Title 'Selecciona el archivo con las cuentas de envío' -Path $GBL_SC_PATH
                    $DB_CLI_CSV_PT = Get-FilePath -Title 'Selecciona el archivo con los correos de los destinatarios' -Path $GBL_SC_PATH
                    $DB_HTML_PT = Get-FilePath -Title 'Encuentra el correo en el formato HTML' -Filter 'Archivo HTHML (*.html)|*.html' -Path $GBL_SC_PATH
                    if($Attach){
                        $DB_ATTC = Get-FilePath -Title 'Encuentra el PDF a atachar' -Filter 'Archivos PDF (*.pdf)|*.pdf' -Path $GBL_SC_PATH
                    }
            }
            else{
                $DB_ACC_CSV_PT = $GBL_SC_PATH + '\senders.csv'
                $DB_CLI_CSV_PT = $GBL_SC_PATH + '\clients.csv'
                $DB_HTML_PT = $GBL_SC_PATH + '\mail.html'
                if($Attach){
                    Write-Host '¿Cómo se llama el PDF a adjuntar?'
                    $TMP = Read-Host
                    $DB_ATTC = $GBL_SC_PATH + '\' + $TMP + '.pdf'
                }
            }
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Creating checking variables
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
            $DB_ACC_CSV_CHK = $false
            $DB_CLI_CSV_CHK = $false
            $DB_HTML_CHK = $false
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Loading files in variables and testing it
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
            Write-Host '                                          '
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
            Write-Host '                           Cargando los archivos necesarios                           ' -foregroundcolor Cyan
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
            try{$DB_ACC_CSV  = Import-Csv -Path $DB_ACC_CSV_PT -Delimiter ","}
            catch{$DB_ACC_CSV_CHK = $true}
            try{$DB_CLI_CSV  = Import-Csv -Path $DB_CLI_CSV_PT -Delimiter ","}
            catch{$DB_CLI_CSV_CHK = $true}
            try{$DB_HTML = ConvertFrom-HTMLtoMail -Path $DB_HTML_PT}
            catch{$DB_HTML_CHK = $true}
            if($DB_ACC_CSV_CHK -or $DB_CLI_CSV_CHK -or $DB_HTML_CHK){
                if($DB_ACC_CSV_CHK -and $DB_CLI_CSV_CHK -and $DB_HTML_CHK){
                    Write-Host '                                          '
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '                    Los archivos indicados no pudieron ser cargados                   ' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                }
                elseif($DB_ACC_CSV_CHK){
                    Write-Host '                                          '
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '               El archivo CSV de las cuentas para enviar, no es correcto.             ' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                }
                elseif($DB_CLI_CSV_CHK){
                    Write-Host '                                          '
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '                 El archivo CSV de los destinatarios, no es correcto.                 ' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                }
                elseif($DB_HTML_CHK){
                    Write-Host '                                          '
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '                 La ruta del archivo HTML no es correcta.                 ' -foregroundcolor Yellow -BackgroundColor Black
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                }
            break
            }
            elseif(!$DB_ACC_CSV_CHK -and !$DB_CLI_CSV_CHK -and !$DB_HTML_CHK){
                if(!$DB_ACC_CSV_CHK){
                    if(!(Test-CSVHeader -ImportedCSV $DB_ACC_CSV -TestValue '*ccount') -or
                       !(Test-CSVHeader -ImportedCSV $DB_ACC_CSV -TestValue '*assword')){
                            Write-Host '                                          '
                            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                            Write-Host '     El archivo CSV de las cuentas para enviar debe poseer "Account" y "Password"     ' -foregroundcolor Yellow -BackgroundColor Black
                            Write-Host '                                como cabeceras del CSV                                ' -foregroundcolor Yellow -BackgroundColor Black
                            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                            break
                    }
                    elseif(!(Test-CSVHeader -ImportedCSV $DB_CLI_CSV -TestValue '*ail')){
                            Write-Host '                                          '
                            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                            Write-Host '             El archivo CSV de las cuentas para enviar debe poseer "Mail"             ' -foregroundcolor Yellow -BackgroundColor Black
                            Write-Host '                                como cabeceras del CSV                                ' -foregroundcolor Yellow -BackgroundColor Black
                            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow -BackgroundColor Black
                            break
                    }
                }
                Write-Host '                                          '
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
                Write-Host '                   Todos los archivos fueron cargados correctamente.                  ' -foregroundcolor Green
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            }
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Creating and loading global variables
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
            $GBL_FIN = 0
            $GBL_CLI = 0
            $GBL_SEND = 25
            $GBL_ERROR = 0
            $GBL_NO_LOG = $true
            if(!$Subject){
                Write-Host "¿Qué asunto tendrá el correo?" -ForegroundColor Yellow
                $SUB_MSG = Read-Host 
            }
            $TIME = Get-Date
            $GBL_DAT = $TIME.Day.ToString() + '-' + $TIME.Month.ToString() + '-' + $TIME.Year.ToString() + '_' + $TIME.Hour.ToString() + '-' + $TIME.Minute.ToString() + '-' + $TIME.Second.ToString()
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Creating log and error files
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
            $ErrorLogName = $GBL_SC_PATH + '\BulkMailErrorLog-' + $GBL_DAT + '.csv'
            "Time,Sender,Mail" | Out-File $ErrorLogName -Append

            $SW1_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Yes"
            $SW1_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
            $SW1_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW1_1ST, $SW1_2ND)
            $SW1_ASW = $host.ui.PromptForChoice("Aditional information", "¿Requieres un registro de lo enviado?", $SW1_OPT, 0) 
            switch ($SW1_ASW){
                0{   $CheckLogName = $GBL_SC_PATH + '\BulkMailCheckLog-' + $GBL_DAT + '.csv'
                    "Time,Sender,Mail" | Out-File $CheckLogName -Append}
                1{   $GBL_NO_LOG = $false}
            }
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
        # Loading message internal images
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
            $FLD_IMG = Get-Item $GBL_SC_PATH\Images -ErrorAction SilentlyContinue
            if($FLD_IMG -eq $null){$FLD_IMG = Get-FolderPath}
            $Inline_ATT = @{}
            $IMG_DB = Get-ChildItem $FLD_IMG | Select Name,Fullname | Where {($_.Name -like '*.png') -or ($_.Name -like '*.jpeg') -or ($_.Name -like '*.jpg') -or ($_.Name -like '*.bmp')-or ($_.Name -like '*.ico')}
            foreach($IMG in $IMG_DB){
                $INT_NAM = $IMG.Name.Substring(0,($IMG.Name).IndexOf('.'))
                $Inline_ATT.Add($INT_NAM,$IMG.Fullname)
            }
        #-----------------------------------------------------------------------------------------------------------------------------------------------------------
    }
    process{
        do{
            $GBL_ACC = 0
    
            foreach($ACC in $DB_ACC_CSV){
                $GBL_BCC = 0
                $GBL_BCC_TX = ""
                $ACC_USR  = $ACC.Account
                $ACC_PASS = ConvertTo-SecureString -String $ACC.Password -AsPlainText -Force
                $ACC_CRE  = New-Object System.Management.Automation.PSCredential $ACC_USR, $ACC_PASS
        
                do{
                    try{
                        if(!$Attach){
                            Send-EmailMessage -Credential $ACC_CRE -From $ACC_USR -To $DB_CLI_CSV[$GBL_ERROR].Mail -Subject $SUB_MSG -Body $DB_HTML -BodyAsHtml -smtpserver outlook.office365.com -usessl -InlineAttachments $Inline_ATT -ErrorAction Continue
                        }
                        else{
                            Send-EmailMessage -Credential $ACC_CRE -From $ACC_USR -To $DB_CLI_CSV[$GBL_ERROR].Mail -Subject $SUB_MSG -Body $DB_HTML -BodyAsHtml -smtpserver outlook.office365.com -usessl -InlineAttachments $Inline_ATT -ErrorAction Continue -Attachments $DB_ATTC
                        }
                    }
                    catch{
                        if($DB_CLI_CSV[$GBL_ERROR].Mail){
                            $TIME.ToString() + "," + $ACC_USR + "," + $DB_CLI_CSV[$GBL_ERROR].Mail | Out-File $ErrorLogName -Append
                        }
                    }
                    if($DB_CLI_CSV[$GBL_ERROR].Mail){
                        if($GBL_NO_LOG){$TIME.ToString() + "," + $ACC_USR + "," + $DB_CLI_CSV[$GBL_ERROR].Mail | Out-File $CheckLogName -Append}
                        Write-Host $TIME.ToString() "," $ACC_USR "," $DB_CLI_CSV[$GBL_ERROR].Mail
                    }

                    $GBL_BCC++ # Breaker of the loop for change sender account #
                    $GBL_ERROR++ # Global line of DB of clients  #
                }until($GBL_BCC -ge $GBL_SEND)
        
                $GBL_CLI += $GBL_SEND
                $GBL_ACC++ # Global counter for cound all account in Senders' account #

                if($GBL_ACC -eq $DB_ACC_CSV.Count){
                    Write-Host "Sleeping 120 seconds"
                    if($GBL_NO_LOG){"-------------------,--- Sleep 120 seconds ----,-----------------------" | Out-File $CheckLogName -Append}
                    sleep 120
                }
                if($GBL_CLI -ge $DB_CLI_CSV.Count-1){$GBL_FIN++; break}
            }
        }
        until($GBL_FIN -eq 1)
    }
    end{
        Write-Host '----------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '  Proceso de envío masivo terminado correctamente el:' (Get-Date) -ForegroundColor Cyan
        Write-Host '----------------------------------------------------------------------------' -ForegroundColor Cyan
        if(!(Confirm-InteractiveEnviroment) -and ($ScriptVersion)){
            break
        }
        else{
            pause
        }
    }
}