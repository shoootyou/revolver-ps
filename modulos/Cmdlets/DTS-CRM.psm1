#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Copyright (c) 2016 Rodolfo Castelo Méndez. Dos Tercios de Shell
#
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#
#       Versión 2.0.0
#    19 de Julio del 2016
#

function Import-CRMSDKModule{
    <#
        .SYNOPSIS
        Permite la importación del SDK de CRM
        
        .DESCRIPTION
        Permite la carga de las dlls necesarias para las diferentes tareas en masa de CRM 
        a través de PowerShell.
        
        .LINK
        Para mayor información por favor verificar 'Import-CRMSDKModule' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Import-CRMSDKModule

        Pregunta a través de una ventana la ruta donde se encuentra descomprimido el módulo de CRM
        
        .EXAMPLE
        Import-CRMSDKModule -Path C:\CRMSDK

        Importa los archivos necesarios del SDK de CRM en la ruta especificada.

        .PARAMETER Path
        Parametro de tipo string orientado al establecimiento de la ruta, orientado a actividades en scripts.
    #>
    [cmdletbinding()]
    param(
        
        [Parameter(Position=0)]
        [string]$Path
    )
    begin{
        if(!$Path){
            $VisualPath = $true
        }
        elseif($Path){
            $VisualPath = $false
        }
    }
    process{
        if($VisualPath -eq $true){
            $INT_PTH = Get-FolderPath -Description 'Encuentra la carpeta del SDK de CRM' -NewFolder $false 
            if(!$INT_PTH){
                break
            }
            else{
                $InternalPath =  $INT_PTH + '\SDK\Bin'
            }
        }
        elseif($VisualPath -eq $false){
            $InternalPath = $Path + '\SDK\Bin'
        }
        if(Test-Path $InternalPath){
            [void][System.Reflection.Assembly]::LoadFile("$InternalPath\microsoft.xrm.sdk.dll")
            [void][System.Reflection.Assembly]::LoadFile("$InternalPath\microsoft.crm.sdk.proxy.dll")
            [void][System.Reflection.Assembly]::LoadWithPartialName("system.servicemodel")
            Write-Host '                                          '
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            Write-Host '                         Módulos de CRM cargados exitosamente                         ' -foregroundcolor Green
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
        }
        elseif(!(Test-Path $InternalPath)){
            Write-Warning 'La ruta seleccionada no contiene los binarios necesarios para el trabajo de PowerShell CRM'
        }
    }

}

function Connect-CRMOnlineServices{
    <#
        .SYNOPSIS
        Permite la conexión al servicio de CRM Online a través de su API
        
        .DESCRIPTION
        Permite la conexión al servicio de CRM Online a través de su API para poder generar las 
        diversas actividades en masa a través de PowerShell. La sesión por defecto dura 10 minutos abierta.
        
        .LINK
        Para mayor información por favor verificar 'Connect-CRMOnlineServices' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Connect-CRMOnlineServices -OnlineOrganization organizationame

        Realiza el inicio de sesión en la organización organizationame.crm.dynamics.com mediante el mecanismo
        visual de solicitud de contraseña.
        
        .EXAMPLE
        Connect-CRMOnlineServices -OnlineOrganization organizationame -SessionMinutes 15

        Realiza el inicio de sesión en la organización organizationame.crm.dynamics.com mediante el mecanismo
        visual de solicitud de contraseña y establece una sesión de 15 minutos.

        .EXAMPLE
        Connect-CRMOnlineServices -OnlineOrganization organizationame -Username usuario@onmicrosoft.com -Password ClaveAqui

        Realiza el inicio de sesión en la organización organizationame.crm.dynamics.com con el usuario "usuario@onmicrosoft.com"
        y la clave "ClaveAquí".

        .PARAMETER OnlineOrganization
        Parametro de tipo string que permite establecer la dirección de la organización a conectar, este valor es el que se
        utiliza en el dominio de CRM por ejemplo: midominio.crm.dynamics.com

        .PARAMETER VisualCredential
        Parametro de tipo bool que permite la habilitación o deshabilitación de la interfaz visual de solicitud de contraseña.

        .PARAMETER Username
        Parametro de tipo string que sirve para almacenar el usuario en un ambiente de automatización

        .PARAMETER Password
        Parametro de tipo string que sirve para almacenar la contraseña en un ambiente de automatización
        
        .PARAMETER WarningDisable
        Parametro de tipo bool que permite deshabilitar el recordatorio de almacenaje en la variable $global:CRMOnlineService.

        .PARAMETER SessionMinutes
        Parametro de tipo int32 que permite establecer la cantidad de minutos para la sesión abierta.
    #>
    [cmdletbinding(
        DefaultParameterSetName='Visual'
    )]
    param(
        [Parameter(Position=0,Mandatory=$true,ParameterSetName='Visual',
        HelpMessage="Introduce el nombre de tu organización, la puedes ubicar en la página Web de CRM siguiendo el ejemplo: https://organizacióndeejemplo.crm.dynamics.com donde 'organizaciondeejemplo' es lo que necesitas.")]
        [Parameter(Position=0,Mandatory=$true,ParameterSetName='Script',
        HelpMessage="Introduce el nombre de tu organización, la puedes ubicar en la página Web de CRM siguiendo el ejemplo: https://organizacióndeejemplo.crm.dynamics.com donde 'organizaciondeejemplo' es lo que necesitas.")]
        [string]$OnlineOrganization,
        [Parameter(Position=1,ParameterSetName='Visual')]
        [bool]$VisualCredential=$true,
        [Parameter(Position=1,Mandatory=$true,ParameterSetName='Script')]
        [ValidateNotNullOrEmpty()]
        [string]$Username,
        [Parameter(Position=2,Mandatory=$true,ParameterSetName='Script')]
        [ValidateNotNullOrEmpty()]
        [string]$Password,
        [Parameter(Position=3,Mandatory=$false,ParameterSetName='Visual')]
        [Parameter(Position=3,Mandatory=$false,ParameterSetName='Script')]
        [bool]$WarningDisable = $true,
        [int32]$SessionMinutes = 10
    )
    begin{
        $CRMOnline_URL = 'https://' + $OnlineOrganization + '.api.crm.dynamics.com/XRMServices/2011/Organization.svc'
        $CRMOnline_CRED = new-object System.ServiceModel.Description.ClientCredentials
    }
    process{
        if($Username){
            if($Username -like '*@*'){
                $CRMOnline_CRED.UserName.UserName = $Username
            }
            else{
                Write-Warning 'El usuario proporcionado no cumple el formato requerido para iniciar sesión en CRM Online'
                pause
                break
            }
        }
        if($Password){
            $CRMOnline_CRED.UserName.Password = $Password
        }
        switch ($PsCmdlet.ParameterSetName) {
            "Visual" {
                $VisualCRED = Get-Credential -ErrorAction SilentlyContinue
                if(!$VisualCRED){
                    Write-Warning 'No se proporcionó ningún valor de usuario y contraseña'
                    pause
                    break
                }
                if(($VisualCRED.UserName.ToString()) -like '*@*' ){
                    $CRMOnline_CRED.UserName.UserName = $VisualCRED.UserName.ToString()
                }
                elseif(($VisualCRED.UserName.ToString()) -notlike '*@*' ){
                    Write-Warning 'El usuario proporcionado no cumple el formato requerido para iniciar sesión en CRM Online'
                    pause
                    break
                }
                $TMP_1_PASS = $VisualCRED.Password
                $TMP_2_PASS = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($TMP_1_PASS)
                $CRMOnline_CRED.UserName.Password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($TMP_2_PASS).ToString()
            }
        }
        
        if(($CRMOnline_CRED.UserName.UserName) -and ($CRMOnline_CRED.UserName.Password)){
            Write-Host '                                          #' 
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
            Write-Host '                     Intentando iniciar sesión, por favor espere.                     ' -foregroundcolor Cyan
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
            $CRMOnline_SRV = new-object Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy($CRMOnline_URL, $null, $CRMOnline_CRED, $null)
            $CRMOnline_SRV.Timeout = new-object System.Timespan(0, $SessionMinutes, 0)
            $RQS = new-object Microsoft.Crm.Sdk.Messages.WhoAmIRequest
            try{
                $LGN_OK = $CRMOnline_SRV.Execute($RQS)
                Write-Host '                                          #' 
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
                Write-Host '             Sesión con el servidor de Dynamics CRM iniciada exitosamente             ' -foregroundcolor Green
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            }
            catch{
                Write-Warning 'Las credenciales proporcionadas no pudieron iniciar sesión en CRM Online'
				pause
				break
            }
            if(($LGN_OK) -and ($WarningDisable -eq $false)){
                if(!(Get-Variable | Where-Object {$_.Name -eq 'CRMOnlineService'})){
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
                    Write-Host '            Para el uso de las demás funciones debe almacenar la ejecución            ' -foregroundcolor Yellow
                    Write-Host '                 de Connect-CRMOnlineServices en la variable de nombre                ' -ForegroundColor Yellow
                    Write-Host '                               $global:CRMOnlineService                               ' -ForegroundColor Yellow
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
                }
            }
            Return $CRMOnline_SRV
        }
        elseif(($CRMOnline_CRED.UserName.UserName) -and (!$CRMOnline_CRED.UserName.Password)){
            Write-Warning 'No se registra la contraseña para el inicio de sesión'
            break
        }
        else{
            Write-Warning 'No se registra usuario y contraseña para el inicio de sesión'
            break
        }
    }
}

function Test-CRMOnlineServices{
    <#
        .SYNOPSIS
        Permite la validación de la conexión al servicio de CRM Online a través de su API.
        
        .DESCRIPTION
        Permite la validación de la conexión al servicio de CRM Online a través de su API devolviendo un valor de
        true o False, dependiendo del estado de la conexión.
        
        .LINK
        Para mayor información por favor verificar 'Test-CRMOnlineServices' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Test-CRMOnlineServices

        Retorna el valor de True si en caso está establecida la sesión y False si en caso no se logró iniciar correctamente.
    #>
    [cmdletbinding()]
    param(
    )
    process{
        $RQS = new-object Microsoft.Crm.Sdk.Messages.WhoAmIRequest
        try{
            $LGN_OK = $CRMOnlineService.Execute($RQS)
            return $true
        }
        catch{
            return $false
        }
    }
}

Function Add-CRMActivityLastInformation{
    <#
        .SYNOPSIS
        Permite la adición de una línea de información a la actividad proporcionada.
        
        .DESCRIPTION
        Permite la adición de una línea de información a la(s) actividad(es) proporcionadas a través de un CSV
        asi como su reprogramación de vencimiento.
        
        .LINK
        Para mayor información por favor verificar 'Add-CRMActivityLastInformation' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Add-CRMActivityLastInformation

        Solicita el archivo CSV asi como la información textual que se introducirá a través de una interfaz gráfica.
        .EXAMPLE
        Add-CRMActivityLastInformation -UpdateSchedule $true

        Solicita el archivo CSV , la información textual y la fecha y hora que se introducirá a través de una interfaz gráfica.
        .EXAMPLE
        Add-CRMActivityLastInformation -Information 'Nuevo contacto'

        Solicita el archivo CSV a través de una interfaz gráfica y adiciona la información 'Nuevo Contacto' a todas las actividades encontradas.
        .EXAMPLE
        Add-CRMActivityLastInformation -Information 'Nuevo contacto' -ActivityDB C:\Users\Prueba\Desktop\activity.csv

        Aadiciona la información 'Nuevo Contacto' a todas las actividades encontradas en el archivo C:\Users\Prueba\Desktop\activity.csv .
        .EXAMPLE
        Add-CRMActivityLastInformation -Information 'Nuevo contacto' -ActivityDB C:\Users\Prueba\Desktop\activity.csv -UpdateScheduleScript 09/06/2016-15_15_00

        Aadiciona la información 'Nuevo Contacto' a todas las actividades encontradas en el archivo C:\Users\Prueba\Desktop\activity.csv y tendrá una fecha de vencimiento de 09/06/2016 a las 15 horas con 15 minutos y 0 segundos.
		
        .PARAMETER VisualInformation
        Parámetro de tipo booleano que permite habilitar o deshabilitar la solicitud de información a través de una interfaz gráfica.

        .PARAMETER Information
        Parámetro de tipo string que permite la inserción directa de la información a adicionar.

        .PARAMETER ActivityType
        Parámetro de tipo string que permite la inserción directa de la información a adicionar.

        .PARAMETER ActivityDB
        Parametro de tipo string que permite la especificacion literal del archivo CSV. Usado en modo de automatización.

        .PARAMETER UpdateSchedule
        Parametro de tipo booleano que permite habilitar o deshabilitar la posibilidad de actualizar la fecha de vencimiento de la actividad.

        .PARAMETER UpdateScheduleScript
        Parametro de tipo string que permite la inserción de fechas en el modo script. 
        Formato dd/MM/yyyy-HH_mm_ss
    #>
    [cmdletbinding(
        DefaultParameterSetName='Visual'
    )]
    param(
            [Parameter(Position=0,Mandatory=$false,ParameterSetName='Visual')]
            [bool]$VisualInformation = $true,
            [Parameter(Position=0,Mandatory=$true,ParameterSetName='Script')]
            [string]$Information,
            [Parameter(Position=1,Mandatory=$false,ParameterSetName='Visual')]
            [Parameter(Position=1,Mandatory=$true,ParameterSetName='Script')]
            [string]$ActivityType,
            [Parameter(Position=2,Mandatory=$false,ParameterSetName='Visual')]
            [Parameter(Position=2,Mandatory=$false,ParameterSetName='Script')]
            [string]$ActivityDB,
            [Parameter(Position=3,Mandatory=$false,ParameterSetName='Visual')]
            [bool]$UpdateSchedule = $false,
            [Parameter(Position=3,Mandatory=$false,ParameterSetName='Script',
            HelpMessage="Ingrese la hora y fecha de la cumpliendo el siguiente formato dd/MM/yyyy-HH_mm_ss")]
            [string]$UpdateScheduleScript
    )
    begin{
        try{
            New-object Microsoft.Crm.Sdk.Messages.WhoAmIRequest | Out-Null
            New-Object Microsoft.Xrm.Sdk.Messages.RetrieveRequest  | Out-Null
        }
        catch{
            Write-Host 'WARNING: No se verifican los módulos de CRM cargados, por favor, ejecute' -ForegroundColor Yellow
            Write-Host '         Import-CRMSDKModule para poder ubicar el SDK de CRM en el sistema,'-ForegroundColor Yellow
            Write-Host '         luego ejecute Connect-CRMOnlineServices con su respectivos datos'  -ForegroundColor Yellow
            Write-Host '         de organización y credenciales para poder iniciar sesión.'-ForegroundColor Yellow
            pause
            break
        }
    }
    process{
        if(!$ActivityType){
            $INT_ACT_TYP = Get-CRMAvailableActivities -Description 'Selecciona el tipo de actividad que modificarás'
        }
        else{
            $ALL_ACT = Get-CRMAvailableActivities -AllActivities $true -ScriptVersion $true
            if($ALL_ACT -contains $ActivityType ){
                $INT_ACT_TYP = $ActivityType
            }
            else{
                Write-Warning 'El tipo de actividad proporcionado no está dentro de las disponibles para modificar'
                break
            }
        }
        if(!$UpdateScheduleScript){
            if($UpdateSchedule){
            $scheduledend = Get-SelectedTime
        }
            else{
            $SW1_1ST = New-Object System.Management.Automation.Host.ChoiceDescription "&Sí", ` "Sí"
            $SW1_2ND = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
            $SW1_OPT = [System.Management.Automation.Host.ChoiceDescription[]]($SW1_1ST, $SW1_2ND)
            $SW1_ASW = $host.ui.PromptForChoice("Información Adicional", "¿Deseas actualizar la fecha de vencimiento", $SW1_OPT, 0) 
            switch ($SW1_ASW){
                0{
                $scheduledend = Get-SelectedTime
                }
                1{
                }
            }
        }
        }
        else{
            $scheduledend = [datetime]::ParseExact($UpdateScheduleScript,"dd/MM/yyyy-HH_mm_ss", [System.Globalization.CultureInfo]::CurrentCulture)
            if(!$scheduledend){
                Write-Warning 'Valor erróneo proporcionado'
                break
            }
        }
        if(!$ActivityDB){
            $DB_ACT_PT = Get-FilePath -Title 'Ubica tu archivo de con los ID de las actividades' -Filter 'Archivos CSV (*.csv) | *.csv' -WarningAction SilentlyContinue
        }
        else{
            $DB_ACT_PT = $ActivityDB
        }
        try{
            $DB_ACT_ID = Import-Csv $DB_ACT_PT
        }
        catch{
            Write-Warning 'No se seleccionó ningún archivo, intente nuevamente.'
            pause
            break
        }
        if(!(Test-CSVHeader -ImportedCSV $DB_ACT_PT -TestValue 'activityid' -ErrorAction SilentlyContinue)){
            Write-Warning 'El archivo proporcionado no posee la cabecera "activityid" en él'
            if(Confirm-InteractiveEnviroment){
                pause
                break
            }
            else{break}
        }

        switch ($PsCmdlet.ParameterSetName) {
            "Visual"{
                $INT_DESC_UP = Get-TextBox -BasicText $true -Width 500 -Title 'Por favor escriba el texto a adicionar'
            }
            "Script"{
                $INT_DESC_UP = $Information
            }
        }
        if(!$INT_DESC_UP){
            Write-Warning 'No se proporcionó texto alguno para actualizar.'
            pause
            break
        }


        $INT_CON_OK = 0
        $INT_DAB_OK = @()
        $INT_CON_ER = 0
        $INT_DAB_ER = @()
        $INT_CON_PR = 1
        foreach($ACC in $DB_ACT_ID){
            $ACT_ID = $ACC.activityid
            if(!($DB_ACT_ID.count)){
                Write-Host '                                          #' 
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
                Write-Host '         Procesando la actividad de ID:  '$ACT_ID -foregroundcolor Cyan
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Cyan
            }
            else{
                Write-Progress -Activity “Actualizando Actividades” -status “Procesando la actividad de ID: $ACT_ID” -percentComplete ($INT_CON_OK / $DB_ACT_ID.count*100)
            }
            $RT_ACT = New-Object Microsoft.Xrm.Sdk.Messages.RetrieveRequest
            $RT_ACT.Target = New-Object Microsoft.Xrm.Sdk.EntityReference($INT_ACT_TYP,$ACC.activityid)
            $RT_ACT.ColumnSet = New-Object Microsoft.Xrm.Sdk.Query.ColumnSet("description")
            $RT_ACT.ColumnSet.AllColumns = $false

            $RT_EXEC = $null
            try{
                $RT_EXEC = $CRMOnlineService.Execute($RT_ACT)
            }
            catch{
                Write-Host
            }
            if(!$RT_EXEC){
                $INT_CON_ER++
                $INT_DAB_ER += $ACC.activityid
            }
            else{
                $RT_EXEC_ATRI = $RT_EXEC.Entity.Attributes

                foreach($INT_ATRIB in $RT_EXEC_ATRI){
                    if($INT_ATRIB.Key -eq 'description'){
                        $UPD_OK = $true
                        $DESC_VAL = $INT_ATRIB.Value
                    }
                }
                if($UPD_OK){
                    $UP_ATRIB_DESC = $INT_DESC_UP + "`n" + $DESC_VAL
            
                    $UPD_ACT_DES = New-Object Microsoft.Xrm.Sdk.Entity($INT_ACT_TYP)
                    $UPD_ACT_DES.Id = [Guid]::Parse($ACC.activityid)
                    $UPD_ACT_DES.Attributes["description"] = $UP_ATRIB_DESC
                    if($UpdateSchedule){
                        $UPD_ACT_DES.Attributes["scheduledend"] = $scheduledend
                    }
                    $CRMOnlineService.Update($UPD_ACT_DES)
                }
                else{
                    $UP_ATRIB_DESC = $INT_DESC_UP
            
                    $UPD_ACT_DES = New-Object Microsoft.Xrm.Sdk.Entity($INT_ACT_TYP)
                    $UPD_ACT_DES.Id = [Guid]::Parse($ACC.activityid)
                    $UPD_ACT_DES.Attributes["description"] = $UP_ATRIB_DESC
                    if($UpdateSchedule){
                        $UPD_ACT_DES.Attributes["scheduledend"] = $scheduledend
                    }
                    $CRMOnlineService.Update($UPD_ACT_DES)
                }
                $INT_DAB_OK += $ACC.activityid
                $INT_CON_OK++
            }
        $INT_CON_PR++
        }
        if(!($DB_ACT_ID.count) -and ($INT_CON_ER -eq 1)){
            Write-Host
            Write-Host '                                          #' 
            Write-Host
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
            Write-Host '                   La actividad no pudo ser procesada correctamente.                  ' -foregroundcolor Yellow
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
        }
        elseif(!($DB_ACT_ID.count)){
            Write-Host
            Write-Host '                                          #' 
            Write-Host
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            Write-Host '                          Actividad procesada correctamente.                          ' -foregroundcolor Green
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
        }
        elseif($DB_ACT_ID.Count -eq ($INT_CON_PR - 1)){
            Write-Host
            Write-Host '                                          #' 
            Write-Host
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            Write-Host '                 ' $INT_CON_OK ' actividad(es) fue(ron) procesada(s) correctamente.                 ' -foregroundcolor Green
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            Write-Host '          Las actividad(es) que se pudieron procesar poseen los siguientes ID         ' -ForegroundColor Green
            Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
            $OUT_OK_CON = 0
            do{
                Write-Host '                         '$INT_DAB_OK[$OUT_OK_CON]'                         ' -ForegroundColor Green
                $OUT_OK_CON++
            }
            while($OUT_OK_CON -lt ($INT_DAB_OK.Count))
            if($INT_CON_ER -ne 0){
                Write-Host
                Write-Host '                                          #' 
                Write-Host
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
                Write-Host '                    ' $INT_CON_ER ' actividad(es) no pudieron ser procesada(s).                     ' -foregroundcolor Yellow
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
                Write-Host '       La(s) actividad(es) que no se pudieron procesar poseen los siguientes ID       ' -ForegroundColor Yellow
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
                $OUT_ERR_CON = 0
                do{
                    Write-Host '                         '$INT_DAB_ER[$OUT_ERR_CON]'                         ' -ForegroundColor Yellow
                    $OUT_ERR_CON++
                }
                while($OUT_ERR_CON -lt ($INT_DAB_ER.Count))
                Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Yellow
            }
        }
    }
    
}

function Get-CRMAvailableActivities{
    <#
        .SYNOPSIS
        Permite la obtención de las actividades disponibles para modificar en Dynamics CRM Online mediante una interfaz visual.
                
        .DESCRIPTION
        Permite la obtención de las actividades disponibles para modificar en Dynamics CRM Online mediante una interfaz visual,
        así como la lista completa de actividades disponibles.
                
        .LINK
        Para mayor información por favor verificar 'Get-CRMAvailableActivities' en
        Dos tercios de shell (http://dosterciosdeshell.wordpress.com)

        .EXAMPLE
        Get-CRMAvailableActivities

        Apertura un cuadro visual de selección para poder elegir las diferentes actividades disponibles para modificar.  
                
        .EXAMPLE
        Get-CRMAvailableActivities -Description 'hola hola'

        Apertura un cuadro visual de selección para poder elegir las diferentes actividades disponibles para modificar con una
        descripción en el cuadro de "hola hola".

        .EXAMPLE
        Get-CRMAvailableActivities -AllActivities $true

        Devuelve todas las actividades disponibles para modificar en Dynamics CRM Online

        .PARAMETER AllActivities
        Parametro de tipo boleano que determina si se desea o no mostrar todas las actividades disponibles

        .PARAMETER Description
        Parametro de tipo String que permite establecer una descripción para guiar al con el fin de esta invocación.
                
        .PARAMETER ScriptVersion
        Parametro de tipo boleano que determina si se desea mostrar todas las actividades en un Arrary. Ideal para scripts.
    #>
    [cmdletbinding(
        DefaultParameterSetName='Full'
    )]
    param(
        [Parameter(Position=0,ParameterSetName='CompleteDB')]
        [bool]$AllActivities,
        [Parameter(Position=1,ParameterSetName='CompleteDB')]
        [bool]$ScriptVersion,
        [string]$Description = 'Selecciona una de las actividades disponibles'
        
    )
    begin{
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    process{
        $DB_AVA_ACT = @('appointment','campaignactivity','campaignresponse','email','fax','letter','phonecall','recurringappointmentmaster','serviceappointment') 
        switch ($PsCmdlet.ParameterSetName) {
            "Full"{
                $GBL_FORM = New-Object Windows.Forms.Form 

                $GBL_FORM.Size = New-Object Drawing.Size @(250,250) 
                $GBL_FORM.StartPosition = "CenterScreen"
                $GBL_FORM.MaximumSize = New-Object Drawing.Size @(260,260) 
                $GBL_FORM.MinimumSize = New-Object Drawing.Size @(260,260) 
                $GBL_FORM.ShowIcon = $false

                $GBL_LBL= New-Object System.Windows.Forms.Label
                $GBL_LBL.Location = New-Object System.Drawing.Point(0, 5)
                $GBL_LBL.Size = New-Object System.Drawing.Point(240, 15)
                $GBL_LBL.TextAlign = "MiddleCenter"
                $GBL_LBL.Margin = 0
                $GBL_LBL.Text = $Description
                $GBL_FORM.Controls.Add($GBL_LBL)

                $GBL_ACT_CMB = New-Object System.Windows.Forms.ComboBox
                $GBL_ACT_CMB.Location = New-Object System.Drawing.Point(0, 25)
                $GBL_ACT_CMB.Size = New-Object System.Drawing.Size(245, 260)
                foreach($Activity in $DB_AVA_ACT)
                {
                  $GBL_ACT_CMB.Items.add($Activity) | Out-Null
                }
                $GBL_FORM.Controls.Add($GBL_ACT_CMB)

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

                $GBL_RESULT = $GBL_FORM.ShowDialog()

                if ($GBL_RESULT -eq [System.Windows.Forms.DialogResult]::OK)
                {
                    if(!$GBL_ACT_CMB.SelectedItem){
                        Write-Warning 'No se seleccionó ninguna de las actividades disponibles.'
                        break
                    }
                    else{
                        $OUTPUT = $GBL_ACT_CMB.SelectedItem
                    }
                }
                else{
                    Write-Warning 'No se seleccionó ninguna de las actividades disponibles.'
                    break
                }
                return $OUTPUT
            }
            "CompleteDB"{
                if($ScriptVersion){
                    return $DB_AVA_ACT
                }
                else{
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
                    Write-Host '         Todas las actividades disponibles para modificar en Dynamics CRM son:        ' -ForegroundColor Green
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
                    $CON_ACT_DB = 0
                    do{
                        Write-Host '                            - '$DB_AVA_ACT[$CON_ACT_DB] -ForegroundColor Green
                        $CON_ACT_DB++
                    }
                    while($CON_ACT_DB -lt ($DB_AVA_ACT.Count-1))
                    Write-Host '--------------------------------------------------------------------------------------' -foregroundcolor Green
                }
            }
        }
    }
}