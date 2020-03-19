#region obtención de información base

<#
Connect-AzAccount 

#>

$COR_AZ_SUB = Get-AzSubscription | Select-Object *

foreach($SUB in $COR_AZ_SUB){
    Select-AzSubscription -Subscription $SUB.SubscriptionId
    $DB_AZ_SUB = Get-AzSubscription -SubscriptionId $SUB.SubscriptionId | Select-Object *

    $DB_AZ_RES = Get-AzResource | Select-Object * | Sort-Object Type

    $DB_AZ_BCK_VLT = Get-AzRecoveryServicesVault
    $DB_AZ_VMS_LST = Get-AzVM | Sort-Object Name
    $DB_AZ_NSG_LST = Get-AzNetworkSecurityGroup
    $DB_AZ_STO_LST = Get-AzStorageAccount
    $DB_AZ_NET_LST = Get-AzVirtualNetwork
    $DB_AZ_PER_LST = $DB_AZ_NET_LST | Foreach-Object { Get-AzVirtualNetworkPeering -ResourceGroupName $_.ResourceGroupName -VirtualNetworkName $_.Name } 
    $DB_AZ_DSK_LST = Get-AzDisk
    $DB_AZ_SQL_LST = Get-AzSqlServer
    $DB_AZ_DBA_LST = $DB_AZ_SQL_LST | Foreach-Object {Get-AzSqlDatabase -ServerName $_.ServerName -ResourceGroupName $_.ResourceGroupName}

    #endregion obtención de información base

    #region Variables, argumentos and functions
    $date= Get-Date -format g
    $version = "5.2" #<<<<<<CHANGE THIS ANYTIME MODIFIED!
    $WarningColor = "#fff200"
    $HTMLHeader = $null
    $HTMLMiddle = $null
    $HTML_Backup = $null
    $HTML_VNET = $null
    $HTML_SQL = $null
    $HTML_NSG = $null
    $HTML_Storage = $null
    $HTML_Details = $null
    $HTMLEnd = $null
    $INT_TXT_CAS = (Get-Culture).TextInfo
    function Convert-HashToString
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [System.Collections.Hashtable]
            $Hash
        )
        $hashstr = "@{"
        $keys = $Hash.keys
        foreach ($key in $keys)
        {
            $v = $Hash[$key]
            if ($key -match "\s")
            {
                $hashstr += "`"$key`"" + "=" + "`"$v`"" + ";"
            }
            else
            {
                $hashstr += $key + "=" + "`"$v`"" + ";"
            }
        }
        $hashstr += "}"
        return $hashstr
    }

    $OBJ_TMP_01 = @()
    if($DB_AZ_BCK_VLT){
        foreach($RS_INT_BCK_OBJ in $DB_AZ_BCK_VLT){
            $DB_AZ_BCK_CNT = Get-AzRecoveryServicesBackupContainer -ContainerType AzureVM -VaultId $RS_INT_BCK_OBJ.ID | Sort-Object FriendlyName
            foreach($RS_INT_BCK_CNT in $DB_AZ_BCK_CNT){
                $DB_AZ_BCK_ITM = Get-AzRecoveryServicesBackupItem -Container $RS_INT_BCK_CNT -WorkloadType AzureVM -VaultId $RS_INT_BCK_OBJ.ID
                $OBJ_TMP_02 = New-Object PSObject
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'VirtualMachine' -Value $DB_AZ_BCK_ITM.VirtualMachineId
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'ProtectionStatus' -Value $DB_AZ_BCK_ITM.ProtectionStatus
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'ProtectionState' -Value $DB_AZ_BCK_ITM.ProtectionState
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'LastBackupStatus' -Value $DB_AZ_BCK_ITM.LastBackupStatus
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'LastBackupTime' -Value $DB_AZ_BCK_ITM.LastBackupTime
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'ProtectionPolicyName' -Value $DB_AZ_BCK_ITM.ProtectionPolicyName
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'LatestRecoveryPoint' -Value $DB_AZ_BCK_ITM.LatestRecoveryPoint
                $OBJ_TMP_02 | Add-Member -type NoteProperty -Name 'ContainerName' -Value $DB_AZ_BCK_ITM.ContainerName
                $OBJ_TMP_01 += $OBJ_TMP_02
            }
        }
    }
    $OBJ_TMP_03 = @()
    $OBJ_TMP_01.VirtualMachine | ForEach-Object {
        $OBJ_TMP_03 += $_.Substring($_.LastIndexOf("/")).Replace("/","")
    }

    #endregion Variables and Arguments

    #region Header
    $HTMLHeader = @"
    <html><head><title>Reporte de recursos en Azure</title>
    <style type="text/css">
    <!--
    body {
            font: 80% Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif; 
            margin-top: 10px;
            margin-left: 10px;
            }     

    h1{
            background-color: inherit;
            color: #606060;
            font-size: 1.6em;
            font-weight: normal;
            letter-spacing: 1px;
            
            }

    h2{
            
            background-color: inherit;
            color: #606060;
            font-size: 1.2em;
            font-weight: normal;
            letter-spacing: 1px;
            
            }
    h3 {

            color: #606060;
            font-size: 0.8em;
            font-weight: normal;
            letter-spacing: 1px;
            }

    p.subtitle{
            background-color: inherit;
            color: #606060;
            font-size: 0.9em;
            }

    .toggle-box {
        display: none;
    }

    .toggle-box + label {
        padding: 10px;
        cursor: pointer;
        display: block;
        clear: both;
        font-size: 16px;
        margin-right: auto;
        margin-bottom:5px;
        text-align: left;
    }

    .toggle-box + label:hover {
        text-shadow:1px 1px 1px rgba(0,0,0,0.1)
    }

    .toggle-box + label + div {
        display: none;
        margin-left: 0px;
        margin-right: auto;
    }

    .toggle-box:checked + label {
        font-style: italic;
    }

    .toggle-box:checked + label + div {
        display: block;
        margin-right: auto;

    }

    .toggle-box + label:before {
        content: "";
        display: block;
        float: left;
        border-right: 2px solid;
        border-bottom: 2px solid;
        width: 5px;
        height: 5px;
        transform: rotate(-45deg);
        margin-top: 6px;
        margin-right: 20px;
        margin-left: auto;
        text-align: left;
        -webkit-transition: all 0.5s;
        transition: all 0.5s;
    }

    .toggle-box:checked + label:before {
        transform: rotate(45deg);
        -webkit-transition: all 0.5s;
        transition: all 0.5s;
    }

    -->
    </style>
    </head>
    <body>
    <h1>Reporte de recursos en Azure</h1></ br><h2>$($DB_AZ_SUB.Name) | $($DB_AZ_SUB.SubscriptionId) </h2><p class="subtitle">Version $version | $date.</p>
"@
    #endregion header

    #region de Storage

    $HTML_Storage+="<input class=""toggle-box"" id=""identifier-Storage"" type=""checkbox"" name=""grouped""><label for=""identifier-Storage"">Storages Accounts (Storage)</label><div>"

    $HTML_Storage+="<table border=""0"" width=""100%"" cellpadding=""4""  style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#000099"">"
    $HTML_Storage+= "<th colspan= ""14"" ><font color=""#FFFFFF"">Detalle de Storage Accounts (Storage)</font></th></tr>"
    $HTML_Storage+="<tr bgcolor=""#0000FF"">"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Grupo de Recursos</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Nombre de Storage</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Ubicación</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">SKU</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Tipo</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Tier de Acceso</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Fecha de creación</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Dominio personalizado</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Último Failover</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Ubicación primaria</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Estado de ubicación primaria</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Ubicación secundaria</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Estado de ubicación secundaria</font></th>"
    $HTML_Storage+="<th align=""center""><font color=""#FFFFFF"">Tráfico sólo por HTTPS</font></th>"
    $HTML_Storage+="</tr>"
    $HTML_Storage+="<tr align=""center"" bgcolor=""#dddddd"">"

    $INT_CNT_HTS_ONN = 0
    $INT_CNT_HTS_OFF = 0
    $INT_ALT_RW = 0

    foreach($INT_TMP_STO in $DB_AZ_STO_LST){
        $HTML_Storage+="<tr"
        #region linea intercaladas
        if ($INT_ALT_RW)
        {
            $HTML_Storage+=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $INT_ALT_RW=1
        }
        #endregion linea intercaladas

        $HTML_Storage+=">"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.ResourceGroupName)</font></td>" 
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.StorageAccountName)</font></td>" 
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.Location)</font></td>" 
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.Sku.Name)</font></td>" 
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.Kind)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.AccessTier)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.CreationTime)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.CustomDomain)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.LastGeoFailoverTime)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.PrimaryLocation)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.StatusOfPrimary)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.SourceAddressPrefix)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.SecondaryLocation)</font></td>"
        $HTML_Storage+="<td align=""center""><font color=""#000000"">$($INT_TMP_STO.EnableHttpsTrafficOnly)</font></td></tr>" 
        if($INT_TMP_STO.EnableHttpsTrafficOnly){
            $INT_CNT_HTS_ONN++
        }
        else{
            $INT_CNT_HTS_OFF++
        }
    }
    $HTML_Storage+="</table>"

    $HTML_Storage+="</div>"

    $HTML_Storage+= "<br />"

    #endregion de Storage

    #region de NSG

    $HTML_NSG+="<input class=""toggle-box"" id=""identifier-NSG"" type=""checkbox"" name=""grouped""><label for=""identifier-NSG"">Network Security Groups (NSG)</label><div>"

    $HTML_NSG+="<table border=""0"" width=""100%"" cellpadding=""4""  style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#000099"">"
    $HTML_NSG+= "<th colspan= ""14"" ><font color=""#FFFFFF"">Detalle de Network Security Groups (NSG)</font></th></tr>"
    $HTML_NSG+="<tr bgcolor=""#0000FF"">"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Grupo de Recursos</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Nombre de NSG</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Ubicación</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Nombre de regla</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Prioridad</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Acceso</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Dirección</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Protocolo</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Puerto</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Red origen</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Red destino</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">NIC asociada</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">VM referente</font></th>"
    $HTML_NSG+="<th align=""center""><font color=""#FFFFFF"">Subnet asociada</font></th>"
    $HTML_NSG+="</tr>"
    $HTML_NSG+="<tr align=""center"" bgcolor=""#dddddd"">"

    $INT_CNT_PRT_ALL = 0
    $INT_CNT_ADD_ALL = 0
    $INT_CNT_PAD_ALL = 0
    $INT_CNT_RUL_ALL = 0
    $INT_ALT_RW = 0
    foreach($INT_TMP_NSG in $DB_AZ_NSG_LST){

        #region linea intercaladas
        if ($INT_ALT_RW)
        {
            $HTML_TMP_01=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $HTML_TMP_01 =""
            $INT_ALT_RW=1
        }
        #endregion linea intercaladas

        foreach($INT_TMP_NSG_01 in $INT_TMP_NSG.SecurityRules){
            $HTML_NSG+="<tr" + $HTML_TMP_01



            $HTML_NSG+=">"
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG.ResourceGroupName)</font></td>" 
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG.Name)</font></td>" 
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG.Location)</font></td>" 
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.Name)</font></td>"
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.Priority)</font></td>"
            if($INT_TMP_NSG_01.Access -eq "Allow"){
                $HTML_NSG+="<td align=""center"" bgcolor=""#008000""><font color=""#000000"">$($INT_TMP_NSG_01.Access)</font></td>"
            }
            else{
                $HTML_NSG+="<td align=""center"" bgcolor=""#ff0000""><font color=""#000000"">$($INT_TMP_NSG_01.Access)</font></td>"
            }
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.Direction)</font></td>"
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.Protocol)</font></td>"
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.DestinationPortRange)</font></td>"
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.SourceAddressPrefix)</font></td>"
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG_01.DestinationAddressPrefix)</font></td>"
            if($INT_TMP_NSG.NetworkInterfaces.id){
                $HTML_NSG+="<td align=""center""><font color=""#000000"">$(($INT_TMP_NSG.NetworkInterfaces.id).Substring(($INT_TMP_NSG.NetworkInterfaces.id).LastIndexOf("/")))</font></td>"
                $INT_TMP_NSG_02 = (Get-AzNetworkInterface -ResourceId $INT_TMP_NSG.NetworkInterfaces.id).VirtualMachine
                if($INT_TMP_NSG_02){
                    $HTML_NSG+="<td align=""center""><font color=""#000000"">$(($INT_TMP_NSG_02.id).Substring(($INT_TMP_NSG_02.id).LastIndexOf("/")))</font></td>" 
                }
                else{
                    $HTML_NSG+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                }
            }
            else{
                $HTML_NSG+="<td align=""center""><font color=""#000000"">$()</font></td>" 
            }
            $HTML_NSG+="<td align=""center""><font color=""#000000"">$($INT_TMP_NSG.Subnets)</font></td></tr>" 
            
            if($INT_TMP_NSG_01.DestinationPortRange -eq '*' -and $INT_TMP_NSG_01.SourceAddressPrefix -eq '*' -and $INT_TMP_NSG_01.Access -eq "Allow" ){
                $INT_CNT_PAD_ALL++
            }
            elseif($INT_TMP_NSG_01.DestinationPortRange -eq '*' -and $INT_TMP_NSG_01.Access -eq "Allow"){
                $INT_CNT_PRT_ALL++
            }
            elseif($INT_TMP_NSG_01.SourceAddressPrefix -eq '*'  -and $INT_TMP_NSG_01.Access -eq "Allow"){
                $INT_CNT_ADD_ALL++
            }
            $INT_CNT_RUL_ALL++
        }

    }
    $HTML_NSG+="</table>"

    $HTML_NSG+="</div>"

    $HTML_NSG+= "<br />"

    #endregion de NSG

    if($DB_AZ_SQL_LST){
        #region de SQL

        $HTML_SQL+="<input class=""toggle-box"" id=""identifier-SQL"" type=""checkbox"" name=""grouped""><label for=""identifier-SQL"">Azure SQL Database (SQL)</label><div>"

        #region de SQL Servers

        $HTML_SQL+="<table border=""0"" width=""100%"" cellpadding=""4""  style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#000099"">"
        $HTML_SQL+= "<th colspan= ""14"" ><font color=""#FFFFFF"">Detalle de Azure SQL Servers (SQL)</font></th></tr>"
        $HTML_SQL+="<tr bgcolor=""#0000FF"">"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Grupo de Recursos</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Nombre de SQL Server</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Ubicación</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Usuario Administrador</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Versión de SQL</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">SQL FQDN</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Nombre de BD</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Edición</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Collation</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Tamaño máximo (GB)</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Estado</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Fecha de creación</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">Zona de redundancia</font></th>"
        $HTML_SQL+="<th align=""center""><font color=""#FFFFFF"">DTU</font></th>"
        $HTML_SQL+="</tr>"
        $HTML_SQL+="<tr align=""center"" bgcolor=""#dddddd"">"

        $INT_CNT_SQL_ALL = 0
        $INT_CNT_DBA_ALL = 0
        $INT_ALT_RW = 0
        foreach($INT_TMP_SQL in $DB_AZ_SQL_LST){

            #region linea intercaladas
            if ($INT_ALT_RW)
            {
                $HTML_TMP_01=" style=""background-color:#dddddd"""
                $INT_ALT_RW=0
            } else
            {
                $HTML_TMP_01 =""
                $INT_ALT_RW=1
            }
            #endregion linea intercaladas
            $INT_VAL_DBA = $true
            $HTML_SQL+="<tr" + $HTML_TMP_01
            $HTML_SQL+=">"
            $INT_TMP_SQL_01 = Get-AzSqlDatabase -ServerName $INT_TMP_SQL.ServerName -ResourceGroupName $INT_TMP_SQL.ResourceGroupName | Where-Object {$_.DatabaseName -ne 'master'}
            if($INT_TMP_SQL_01){
                foreach($INT_TMP_SQL_4 in $INT_TMP_SQL_01){
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.ResourceGroupName)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.ServerName)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.Location)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.SqlAdministratorLogin)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.ServerVersion)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.FullyQualifiedDomainName)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.DatabaseName)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.Edition)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.CollationName)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.MaxSizeBytes/1GB)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.Status)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.CreationDate)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.ZoneRedundant)</font></td>" 
                    $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL_4.Capacity)</font></td>"
                    $INT_CNT_DBA_ALL++
                }
            }
            else{
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.ResourceGroupName)</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.ServerName)</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.Location)</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.SqlAdministratorLogin)</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.ServerVersion)</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$($INT_TMP_SQL.FullyQualifiedDomainName)</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>" 
                $HTML_SQL+="<td align=""center""><font color=""#000000"">$()</font></td>"
            }


            $INT_CNT_SQL_ALL++
            }


        $HTML_SQL+="</tr></table>"
        $HTML_SQL+= "<br />"

        #endregion de SQL Servers



        $HTML_SQL+="</div>"

        $HTML_SQL+= "<br />"

        #endregion de SQL
    }

    #region de VNET

    $HTML_VNET+="<input class=""toggle-box"" id=""identifier-VNET"" type=""checkbox"" name=""grouped""><label for=""identifier-VNET"">Virtual Networks (VNET)</label><div>"

    $HTML_VNET+="<table border=""0"" width=""100%"" cellpadding=""4""  style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#000099"">"
    $HTML_VNET+= "<th colspan= ""14"" ><font color=""#FFFFFF"">Detalle de Virtual Networks (VNET)</font></th></tr>"
    $HTML_VNET+="<tr bgcolor=""#0000FF"">"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Grupo de Recursos</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Nombre de VNET</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Ubicación</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Estado de provisionamiento</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Espacio de direcciones</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">DNS Servers</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Protección DDOS habilitada</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Plan de protección DDOS</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Nombre de subnet</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Estado de provisionamiento de Subnet</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Subnet</font></th>"
    $HTML_VNET+="<th align=""center""><font color=""#FFFFFF"">Tarjetas de red asociadas</font></th>"
    $HTML_VNET+="</tr>"
    $HTML_VNET+="<tr align=""center"" bgcolor=""#dddddd"">"

    $INT_CNT_NET_ALL = 0
    $INT_CNT_SUB_ALL = 0
    $INT_ALT_RW = 0
    foreach($INT_TMP_VNET in $DB_AZ_NET_LST){

        #region linea intercaladas
        if ($INT_ALT_RW)
        {
            $HTML_TMP_01=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $HTML_TMP_01 =""
            $INT_ALT_RW=1
        }
        #endregion linea intercaladas

        foreach($INT_TMP_VNET_01 in $INT_TMP_VNET.Subnets){
            $HTML_VNET+="<tr" + $HTML_TMP_01
            Remove-Variable INT_TMP_VNET_02 -ErrorAction SilentlyContinue
            $INT_TMP_VNET.DhcpOptions.DnsServers | Foreach-Object { $INT_TMP_VNET_02 += $_ + "<br />"}

            $HTML_VNET+=">"
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.ResourceGroupName)</font></td>" 
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.Name)</font></td>" 
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.Location)</font></td>" 
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.ProvisioningState)</font></td>" 
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.AddressSpace.AddressPrefixes)</font></td>"
            if($INT_TMP_VNET_02){
                $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET_02)</font></td>" 
            }
            else{
                $HTML_VNET+="<td align=""center""><font color=""#000000"">Azure Provided</font></td>" 
            }
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.EnableDdosProtection)</font></td>" 
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET.DdosProtectionPlan)</font></td>" 
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET_01.Name)</font></td>"
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET_01.ProvisioningState)</font></td>"
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($INT_TMP_VNET_01.AddressPrefix)</font></td>"

            $TEMPO_01 = $null
            if($INT_TMP_VNET_01.IpConfigurations){
                foreach($NET in $INT_TMP_VNET_01.IpConfigurations){
                    $OUT_01 = $NET.Id.Substring(0,$NET.Id.LastIndexOf("/"))
                    $OUT_02 = $OUT_01.Substring(0,$OUT_01.LastIndexOf("/"))
                    $OUT_03 = ($OUT_02.Substring($OUT_02.LastIndexOf("/"))).Replace("/","")
                    $TEMPO_01 += ($OUT_03 + "<br />")
                }
            }
            $HTML_VNET+="<td align=""center""><font color=""#000000"">$($TEMPO_01)</font></td>"
            $INT_CNT_SUB_ALL++
        }
        $INT_CNT_NET_ALL++
    }
    $HTML_VNET+="</table>"

    $HTML_VNET+="</div>"

    $HTML_VNET+= "<br />"

    #endregion de VNET

    #region de Listado de equipos que posen Backup

    $HTML_Backup+="<input class=""toggle-box"" id=""identifier-backup"" type=""checkbox"" name=""grouped""><label for=""identifier-backup"">Virtual Machines and Backup</label><div>"

    $HTML_Backup+="<table border=""0"" width=""100%"" cellpadding=""4""  style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#000099"">"
    $HTML_Backup+= "<th colspan= ""10"" ><font color=""#FFFFFF"">Detalle de máquinas virtuales y backup</font></th></tr>"
    $HTML_Backup+="<tr bgcolor=""#0000FF"">"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Grupo de Recursos</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Nombre de VM</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Ubicación</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Tamaño</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Sistema Operativo</font></th>"

    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Estado de Proteccion</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Estado de último backup</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Hora y fecha de último backup</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Nombre de política de protección</font></th>"
    $HTML_Backup+="<th align=""center""><font color=""#FFFFFF"">Última fecha de recuperación</font></th>"
    $HTML_Backup+="</tr>"
    $HTML_Backup+="<tr align=""center"" bgcolor=""#dddddd"">"

    $INT_CNT_BCK_ON = 0
    $INT_CNT_BCK_OFF = 0
    $INT_ALT_RW = 0
    foreach($INT_TMP_VM in $DB_AZ_VMS_LST){
        $HTML_Backup+="<tr"
        if ($INT_ALT_RW)
        {
            $HTML_Backup+=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $INT_ALT_RW=1
        }

        If($OBJ_TMP_03 -contains $INT_TMP_VM.Name){
            foreach($INT_TMP_01_10 in $OBJ_TMP_01){
                $INT_TMP_01_03 = $INT_TMP_01_10.VirtualMachine.Replace(("/subscriptions/" + $DB_AZ_SUB.SubscriptionId + "/resourceGroups/" ),"")
                $INT_TMP_01_02 = $INT_TMP_01_03.substring($INT_TMP_01_03.lastIndexOf("/")+1)
                if($INT_TMP_01_02 -eq $INT_TMP_VM.Name){
                    $TMP_04 = $INT_TMP_01_10
                }

            }
            $HTML_Backup+=">"
            $INT_TMP_01_03 = $TMP_04.VirtualMachine.Replace(("/subscriptions/" + $DB_AZ_SUB.SubscriptionId + "/resourceGroups/" ),"")
            $INT_TMP_01_01 = ($INT_TMP_01_03.Substring($INT_TMP_01_03.LastIndexOf("resourceGroups")+15)).substring(0,($INT_TMP_01_03.Substring($INT_TMP_01_03.LastIndexOf("resourceGroups")+15)).indexOf("/"))
            $INT_TMP_01_02 = $INT_TMP_01_03.substring($INT_TMP_01_03.lastIndexOf("/")+1)
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_01_01)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_01_02)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.Location)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.HardwareProfile.VmSize)</font></td>" 
            if($INT_TMP_VM.OSProfile.WindowsConfiguration -or $INT_TMP_VM.LicenseType -like 'Windows_*' ){
                $HTML_Backup+="<td align=""center""><font color=""#000000"">$("Windows")</font></td>"
            }
            else{
                $HTML_Backup+="<td align=""center""><font color=""#000000"">$("Linux")</font></td>"
            }
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.StatusCode)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($TMP_04.LastBackupStatus)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($TMP_04.LastBackupTime)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($TMP_04.ProtectionPolicyName)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($TMP_04.LatestRecoveryPoint)</font></td>" 
            $HTML_Backup+="</tr>"
            $INT_CNT_BCK_ON++
        }
        else{
            $HTML_Backup+=">"
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.ResourceGroupName.ToUpper())</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.Name.ToUpper())</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.Location)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.HardwareProfile.VmSize)</font></td>" 
            if($INT_TMP_VM.OSProfile.WindowsConfiguration -or $INT_TMP_VM.LicenseType -like 'Windows_*' ){
                $HTML_Backup+="<td align=""center""><font color=""#000000"">$("Windows")</font></td>"
            }
            else{
                $HTML_Backup+="<td align=""center""><font color=""#000000"">$("Linux")</font></td>"
            } 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$($INT_TMP_VM.StatusCode)</font></td>" 
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$()</font></td>"
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$()</font></td>"
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$()</font></td>"
            $HTML_Backup+="<td align=""center""><font color=""#000000"">$()</font></td>"
            $HTML_Backup+="</tr>"
            $INT_CNT_BCK_OFF++
        }
    }
    $HTML_Backup+="</table>"

    $HTML_Backup+="</div>"

    $HTML_Backup+= "<br />"

    #endregion de Listado de equipos que posen Backup

    #region de resumen de recursos encontrado

    $HTMLMiddle+="<table border=""0"" width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#009900"">"
    $HTMLMiddle+= "<th colspan= ""10"" ><font color=""#FFFFFF"">Información de resumen</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#00CC00"">"
    $HTMLMiddle+="<th><font color=""#000000"">Grupos de Recursos</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Recursos</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Proveedores de Recursos</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Virtual Machines</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Redes Virtuales (VNET)</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Network Security Group</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Network Security Group rules</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Storages Accounts</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Azure SQL Servers</font></th>"
    $HTMLMiddle+="<th><font color=""#000000"">Azure SQL Databases</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_RES.ResourceGroupName | Select-Object -Unique | Measure-Object).Count)</font></td>" 
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_RES | Measure-Object ).Count)</font></td>" 
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_RES | Select-Object Type -Unique | Measure-Object ).Count)</font></td>"
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_VMS_LST | Select-Object Name -Unique | Measure-Object ).Count)</font></td>"
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_NET_LST | Select-Object Name -Unique | Measure-Object ).Count)</font></td>"
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_NSG_LST | Select-Object Name -Unique | Measure-Object ).Count)</font></td>"
    $HTMLMiddle+="<td><font color=""#000000"">$($INT_CNT_RUL_ALL)</font></td>"
    $HTMLMiddle+="<td><font color=""#000000"">$(($DB_AZ_STO_LST | Measure-Object ).Count)</font></td>" 
    $HTMLMiddle+="<td><font color=""#000000"">$($INT_CNT_SQL_ALL)</font></td>"
    $HTMLMiddle+="<td><font color=""#000000"">$($INT_CNT_DBA_ALL)</font></td>"
    $HTMLMiddle+="</tr></table><tr><td></td></tr>"

    #endregion de resumen de recursos encontrado

    #region de resumen de información

    $HTMLMiddle+="<h2>Información resaltante</h2>"

    $HTMLMiddle+="<table width=""100%""><tr><td valign=""top""  align=""center"">"

    #region de cantidad de recursos por grupo de recursos

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">Grupo de Recursos</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Grupo de Recursos</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad de recursos</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"

    $DB_AZ_RES_LCL = $DB_AZ_RES | Sort-Object ResourceGroupName | Select-Object ResourceGroupName -Unique
    $TMP_AZ_RES_LCL = @()
    $DB_AZ_RES_LCL | Foreach-Object { $TMP_AZ_RES_LCL += $_.ResourceGroupName.ToLower() }
    $DB_AZ_RES_LCL = $TMP_AZ_RES_LCL | Select-Object -Unique
    $INT_ALT_RW = 0
    foreach($AZ_RES_LCL in  $DB_AZ_RES_LCL){
        $HTMLMiddle+="<tr"
        if ($INT_ALT_RW)
        {
            $HTMLMiddle+=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $INT_ALT_RW=1
        }
        $HTMLMiddle+=">"
        $CNT_LCL_MSR = $DB_AZ_RES | Where-Object {$_.ResourceGroupName -eq $AZ_RES_LCL} | Measure-Object
        $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($AZ_RES_LCL)</font></td>" 
        $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($CNT_LCL_MSR.Count)</font></td>" 
        $HTMLMiddle+="</tr>"

    }
    $HTMLMiddle+="</table>"
    #endregion de cantidad de recursos por grupo de recursos

    $HTMLMiddle+= "</td><td valign=""top""  align=""center"">"

    #region de cantidad de recursos por region

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">Recursos por Region</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Región</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad de recursos</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"

    $DB_AZ_RES_LCL = $DB_AZ_RES | Sort-Object Location | Select-Object Location -Unique
    $INT_ALT_RW = 0
    foreach($AZ_RES_LCL in  $DB_AZ_RES_LCL){
        $HTMLMiddle+="<tr"
        if ($INT_ALT_RW)
        {
            $HTMLMiddle+=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $INT_ALT_RW=1
        }
        $HTMLMiddle+=">"
        $CNT_LCL_MSR = $DB_AZ_RES | Where-Object {$_.Location -eq $AZ_RES_LCL.Location} | Measure-Object
        $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($AZ_RES_LCL.Location)</font></td>" 
        $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($CNT_LCL_MSR.Count)</font></td>" 
        $HTMLMiddle+="</tr>"

    }
    $HTMLMiddle+="</table><br />"
    #endregion de cantidad de recursos por region

    $HTMLMiddle+= "</td><td valign=""top""  align=""center"">"

    #region de cantidad de recursos por tags

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">Recursos que poseen Tags</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Posee tag</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad de recursos</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"">"
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">YES</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$(($DB_AZ_RES | Where-Object {$_.Tags -ne $null} | Measure-Object).Count)</font></td>" 
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr style=""background-color:#dddddd"">"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"

    $HTMLMiddle+="<td align=""center""><font color=""#000000"">NO</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$(($DB_AZ_RES | Where-Object {$_.Tags -eq $null} | Measure-Object).Count)</font></td>" 
    $HTMLMiddle+="</tr>"

    $HTMLMiddle+="</table>"
    #endregion de cantidad de recursos por region

    $HTMLMiddle+= "</td><td valign=""top""  align=""center"">"

    #region equipos que tienen backup

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">VMs vs backup</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Posee backup</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad de recursos</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"">"
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">YES</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_BCK_ON)</font></td>" 
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr style=""background-color:#dddddd"">"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"

    $HTMLMiddle+="<td align=""center""><font color=""#000000"">NO</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_BCK_OFF)</font></td>" 
    $HTMLMiddle+="</tr>"

    $HTMLMiddle+="</table>"
    #endregion equipos que tienen backup

    $HTMLMiddle+= "</td><td valign=""top""  align=""center"">"

    #region resumen - NSG abiertos

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">Reglas en NSGs abiertos</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Tipo de apertura</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad reglas</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"">"
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">All Source Network and Protocols</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_PAD_ALL)</font></td>" 
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">Permit all Source Networks (Only)</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_ADD_ALL)</font></td>" 
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"">"
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">Permit all Source Protocols (Only)</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_PRT_ALL)</font></td>" 
    $HTMLMiddle+="</tr>"

    $HTMLMiddle+="<tr align=""center"" bgcolor=""#8E1275"">"
    $HTMLMiddle+="<td align=""center""><font color=""#FFFFFF""># de reglas comprometidas</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#FFFFFF"">$($INT_CNT_PRT_ALL + $INT_CNT_ADD_ALL + $INT_CNT_PAD_ALL)</font></td>" 
    $HTMLMiddle+="</tr>"

    $HTMLMiddle+="</table>"
    #endregion resumen - NSG abiertos

    $HTMLMiddle+= "</td><td valign=""top""  align=""center"">"

    #region storages account

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">Storages Account - HTTPS Only</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Posee backup</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad de Storages</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"">"
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">YES</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_HTS_ONN)</font></td>" 
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr style=""background-color:#dddddd"">"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"

    $HTMLMiddle+="<td align=""center""><font color=""#000000"">NO</font></td>" 
    $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($INT_CNT_HTS_OFF)</font></td>" 
    $HTMLMiddle+="</tr>"

    $HTMLMiddle+="</table>"
    #endregion storages account

    $HTMLMiddle+= "</td><td valign=""top""  align=""center"">"

    #region Cantidad de Storages por Replicacion

    $HTMLMiddle+="<table border=""0""  width=""100%"" cellpadding=""4"" style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;""><tr bgcolor=""#8E1275"">"
    $HTMLMiddle+= "<th colspan= ""2"" ><font color=""#FFFFFF"">Storages por replicación</font></th></tr>"
    $HTMLMiddle+="<tr bgcolor=""#e031ba"">"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Replicación</font></th>"
    $HTMLMiddle+="<th align=""center""><font color=""#000000"">Cantidad de Storages</font></th>"
    $HTMLMiddle+="</tr>"
    $HTMLMiddle+="<tr align=""center"" bgcolor=""#dddddd"">"

    $DB_AZ_RES_LCL = $DB_AZ_STO_LST.Sku | Sort-Object Name | Select-Object Name -Unique
    $INT_ALT_RW = 0
    foreach($AZ_RES_LCL in  $DB_AZ_RES_LCL){
        $HTMLMiddle+="<tr"
        if ($INT_ALT_RW)
        {
            $HTMLMiddle+=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $INT_ALT_RW=1
        }
        $HTMLMiddle+=">"
        $CNT_LCL_MSR = $DB_AZ_STO_LST.Sku | Where-Object {$_.Name -eq $AZ_RES_LCL.Name} | Measure-Object
        $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($AZ_RES_LCL.Name)</font></td>" 
        $HTMLMiddle+="<td align=""center""><font color=""#000000"">$($CNT_LCL_MSR.Count)</font></td>" 
        $HTMLMiddle+="</tr>"

    }
    $HTMLMiddle+="</table>"
    #endregion Cantidad de Storages por Replicacion

    $HTMLMiddle+= "</td></tr></table><br />"

    #endregion de resumen de información

    #region detalle de recursos

    $HTML_Details+="<input class=""toggle-box"" id=""identifier-detalle"" type=""checkbox"" name=""grouped""><label for=""identifier-detalle"">Información detallada de recursos</label><div>"

    $HTML_Details+="<table border=""0"" cellpadding=""4""  style=""font-size:8pt;font-family:Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;"">"
    $HTML_Details+= "<th colspan= ""13"" bgcolor=""#ff8800"" ><font color=""#000000"">Detalle de recursos</font></th>"
    $HTML_Details+="<tr bgcolor=""#ff9d2d"">"
    $HTML_Details+="<th align=""center""><font color=""#000000"">#</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Grupo de Recurso</font></th>"
    $HTML_Details+="<th align=""center""><font color=""#000000"">Ubicación</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Nombre</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Fabricante</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Tipo</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Subtipo</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Sub</font></th>"
    $HTML_Details+="<th><font color=""#000000"">SKU Name</font></th>"
    $HTML_Details+="<th><font color=""#000000"">SKU Tier</font></th>"
    $HTML_Details+="<th><font color=""#000000"">Tags</font></th>"
    $HTML_Details+="<tr>"

    $DB_AZ_RES_RG = $DB_AZ_RES | Sort-Object ResourceGroupName
    $INT_ALT_RW = 0
    $AZ_RES_CNT = 1
    foreach($AZ_RES in  $DB_AZ_RES_RG){

        #region linea intercaladas
        if ($INT_ALT_RW)
        {
            $HTML_TMP_01=" style=""background-color:#dddddd"""
            $INT_ALT_RW=0
        } else
        {
            $HTML_TMP_01 =" style=""background-color:#f2f2f2"""
            $INT_ALT_RW=1
        }
        #endregion linea intercaladas


        $HTML_Details+="<tr" + $HTML_TMP_01 + ">"

        $RS_INT_TYP = $INT_TXT_CAS.ToTitleCase(($AZ_RES.Type.Substring(0,$AZ_RES.Type.IndexOf("/"))))
        $RS_INT_FAB = $RS_INT_TYP.Substring(0,$RS_INT_TYP.IndexOf("."))
        $RS_INT_TPY = $RS_INT_TYP.Substring($RS_INT_TYP.IndexOf(".")+1)
        $RS_INT_STP = $AZ_RES.Type.Substring($AZ_RES.Type.IndexOf("/")+1) 

        
        $HTML_Details+="<td align=""center""><font color=""#000000"">$($AZ_RES_CNT)</font></td>" 
        $HTML_Details+="<td><font color=""#000000"">$($AZ_RES.ResourceGroupName)</font></td>" 
        $HTML_Details+="<td align=""center""><font color=""#000000"">$($AZ_RES.Location)</font></td>" 
        $HTML_Details+="<td><font color=""#000000"">$($AZ_RES.Name)</font></td>" 
        $HTML_Details+="<td align=""center""><font color=""#000000"">$($RS_INT_FAB)</font></td>" 
        $HTML_Details+="<td align=""right""><font color=""#000000"">$($RS_INT_TPY)</font></td>" 
        $HTML_Details+="<td><font color=""#000000"">$($RS_INT_STP)</font></td>" 
        $HTML_Details+="<td><font color=""#000000"">$($AZ_RES.Kind)</font></td>" 
        $HTML_Details+="<td align=""center""><font color=""#000000"">$($AZ_RES.Sku.Name)</font></td>" 
        $HTML_Details+="<td align=""center""><font color=""#000000"">$($AZ_RES.Sku.Tier)</font></td>"
        if($AZ_RES.Tags){
            if($($AZ_RES.Tags | Out-String -Stream) -like '*hidden*'){
                $HTML_Details+="<td><font color=""#000000"">$("System tag")</font></td>" 
            }
            else{
                $HTML_Details+="<td><font color=""#000000"">$(Convert-HashToString $AZ_RES.Tags)</font></td>" 
            }
        }
        else{
            $HTML_Details+="<td><font color=""#000000"">$()</font></td>"
        }
        $HTML_Details+="</tr>"
        $AZ_RES_CNT++
    }

    $HTML_Details+="</table>"

    $HTML_Details+="</div>"

    #endregion detalle de recursos

    #region finalizacion y cierre del archivo

    $HTMLEnd = "</body></html>"
    $HTMLFile = $HTMLHeader + $HTMLMiddle + $HTML_Backup + $HTML_Storage + $HTML_VNET + $HTML_NSG + $HTML_SQL + $HTML_Details + $HTMLEnd
    Add-Type -AssemblyName System.Web
    $PRN_PAT = $env:USERPROFILE + "\OneDrive\GitHub\revolver-ps\propio\microsoft\reports\" + (get-Date).ToShortDateString().Replace("/","") + (get-Date).ToShortTimeString().Replace(":","").Replace(" ","") + "-" + ($SUB.Name).Trim() + "-AzureResourcesReport.html"
    [System.Web.HttpUtility]::HtmlDecode($HTMLFile) | Out-File $PRN_PAT 

    #endregion finalizacion y cierre del archivo
}