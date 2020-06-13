<################################################################################################

Author: Rodolfo Castelo Méndez
Versión: 1.0
Required Modules:
    AzureAD
    Az.Table
    Az

    
################################################################################################>

<#region de inicio de sesión y datos

Connect-AzAccount
$TNT_ID = "1682654c-afff-4760-8e95-b9d57482d7d0"
$COR_AZ_TNT_ALL = Connect-AzureAD -TenantId $TNT_ID
$OUT_TBL_CNN = "DefaultEndpointsProtocol=https;AccountName=orionazreport01;AccountKey=z1HEoa90hHAKPr0hW0SNpymDpjRwJxMCJq/XOXNZfms7w+nu/3tfpktzQ5wu6rTPUzrxDkvW3JG87mVfVJusSg==;EndpointSuffix=core.windows.net"
$OUT_TBL_CTX = New-AzStorageContext -ConnectionString $OUT_TBL_CNN
$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"

region obtención de informacion base #>

#region de definicion de variables bases y de inicialización

$GBL_IN_FOR_CNT = 1
$GBL_IN_SUB_CNT = 2

#endregion de definicion de variables bases y de inicialización

#region obtencion de recursos de subscripción
Clear-Host
$COR_AZ_SUB_ALL = Get-AzSubscription -TenantId $TNT_ID | Select-Object *

#endregion obtencion de recursos de subscripción

#region de limpieza de tablas existentes
Write-Host "0 - Limpieza de tablas maestras de recursos" -ForegroundColor DarkGray
Get-AzStorageTable -Context $OUT_TBL_CTX | ForEach-Object { Remove-AzStorageTable -Name $_.Name -Context $OUT_TBL_CTX}
Write-Host "        Procedimiento de limpieza de tablas preexistentes completado exitosamente" -ForegroundColor Green
#endregion de limpieza de tablas existentes

#region de preparación de tablas maestras
Write-Host "1 - Creacion de tablas maestras de recursos" -ForegroundColor DarkGray

Start-Sleep -Seconds 10
$OUT_DB_TBL_SUB = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amastersubscription"
Write-Host "        Tabla de listado de suscripciones creada exitosamente" -ForegroundColor Green

Start-Sleep -Seconds 10
$OUT_DB_TBL_RSG = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterresourcegroup"
Write-Host "        Tabla de listado de grupo de recursos creada exitosamente" -ForegroundColor Green

Start-Sleep -Seconds 10
$OUT_DB_TBL_REG = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterregions"
Write-Host "        Tabla de listado de regiones creada exitosamente" -ForegroundColor Green

Start-Sleep -Seconds 10
$OUT_DB_TBL_RES = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterresources"
Write-Host "        Tabla de listado de informacion general de recursos creada exitosamente" -ForegroundColor Green

Start-Sleep -Seconds 10
$OUT_DB_TBL_REC = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterrecommendations"
Write-Host "        Tabla de listado de recomendaciones de Azure Advisor creada exitosamente" -ForegroundColor Green

Start-Sleep -Seconds 10
$OUT_DB_TBL_PER = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterpermissions"
Write-Host "        Tabla de listado de permisos sobre subscripciones creada exitosamente" -ForegroundColor Green

#endregion de preparación de tablas maestras

foreach($SUB in $COR_AZ_SUB_ALL){
    Clear-Host
    $WR_BAR = $SUB.Name
    Write-Progress -Activity "Cargando informacion" -status "Revisando: $WR_BAR" -percentComplete ($GBL_IN_SUB_CNT / $COR_AZ_SUB_ALL.Count*100) -ErrorAction SilentlyContinue -Id 100
    Write-Host $GBL_IN_SUB_CNT "- Inicializacion de datos para subscripcion" $SUB.SubscriptionId -ForegroundColor DarkGray
    
    #region selección de subscripción y recursos
    Write-Host "    A. Obtencion de informacion de subscripciones y recursos" -ForegroundColor Cyan
    Select-AzSubscription -Subscription $SUB.SubscriptionId | Out-Null
    $DB_AZ_RES_ALL = Get-AzResource | Select-Object * | Sort-Object Type
    $DB_AZ_RSG_ALL = Get-AzResourceGroup | Select-Object * | Sort-Object Type
    $DB_AZ_REC_ALL = Get-AzAdvisorRecommendation | Select-Object * | Sort-Object Category
    $DB_AZ_PER_ALL = Get-AzRoleAssignment | Select-Object * | Sort-Object Scope -Descending
    Write-Host "        Suscription ID   :" $SUB.SubscriptionId -ForegroundColor Green
    Write-Host "        Suscription Name :" $SUB.Name -ForegroundColor Green
    #endregion selección de subscripción y recursos

    #region provisionamiento de tablas de acceso de recursos
    Write-Host "    B. Creacion de tablas para informacion de recursos" -ForegroundColor Cyan
    $DB_AZ_RES_TYP  = @()
    ($DB_AZ_RES_ALL | Select-Object Type -Unique) | ForEach-Object { $DB_AZ_RES_TYP += $_.Type.Replace("/",".").ToLower()}
    $DB_AZ_RES_TYP = $DB_AZ_RES_TYP | Select-Object -Unique

    foreach($RES in $DB_AZ_RES_TYP){
    
        $FOR_INT_00 = $RES.replace(".","")
        $FOR_INT_01 = Get-AzStorageTable -Context $OUT_TBL_CTX -Name $FOR_INT_00 -ErrorAction SilentlyContinue
        if(!$FOR_INT_01){
            Start-Sleep -Seconds 5
            $FOR_INT_01 = New-AzStorageTable -Context $OUT_TBL_CTX -Name $FOR_INT_00
        }

        $FOR_INT_00 = $null
        $FOR_INT_01 = $null

    }
    Write-Host "        Se crearon " $DB_AZ_RES_TYP.Length "tablas de forma exitosa" -ForegroundColor DarkGreen

    #endregion provisionamiento de tablas de acceso de recursos
    
    #region informacion de las subscripciones
    
    Write-Host "    C. Cargado de informacion de subscripcion" -ForegroundColor Cyan
    Add-AzTableRow `
        -UpdateExisting `
        -Table $OUT_DB_TBL_SUB.CloudTable `
        -PartitionKey $SUB.TenantId `
        -RowKey $SUB.SubscriptionId `
        -Property @{
            "Name" = $SUB.Name;
            "State" = $SUB.State;
            "Environment" = ($SUB.ExtendedProperties | ConvertTo-Json | ConvertFrom-Json).Environment;
            "TenantDomain" = $COR_AZ_TNT_ALL.TenantDomain;
            "TenantName" = (Get-AzureADTenantDetail | Select-Object DisplayName).DisplayName;
            "Regions" = ($DB_AZ_RES_ALL | Select-Object Location -Unique).Count;
            "ResourceGroup" = $DB_AZ_RSG_ALL.Count;
            "Resources" = ($DB_AZ_RES_ALL | Measure-Object).Count;
            "ResourceProviders" = ($DB_AZ_RES_TYP | Measure-Object).Count
        } | Out-Null
    Write-Host "        Se cargo la informacion de la subscripcion " $SUB.SubscriptionId "exitosamente" -ForegroundColor DarkGreen

    #endregion informacion de las subscripciones

    #region informacion de las regiones empleadas
    
    Write-Host "    D. Cargado de informacion de regiones empleadas" -ForegroundColor Cyan
    $DB_AZ_RES_REG = $DB_AZ_RES_ALL | Select-Object Location -Unique
    
    foreach($REG in $DB_AZ_RES_REG){
        Add-AzTableRow `
            -UpdateExisting `
            -Table $OUT_DB_TBL_REG.CloudTable `
            -PartitionKey $SUB.SubscriptionId `
            -RowKey $REG.Location `
            -Property @{
                "ResourcesNumber" = ($DB_AZ_RES_ALL | Where-Object {$_.Location -eq $REG.Location} | Measure-Object).Count;
                "TenantId" = $SUB.TenantId
            } | Out-Null
    }
    Write-Host "        Se cargaron " $DB_AZ_RES_REG.Length " regiones exitosamente" -ForegroundColor DarkGreen

    #endregion informacion de las regiones empleadas

    #region información de grupos de recursos
    
    Write-Host "    E. Cargado de informacion de grupo de recursos empleadas" -ForegroundColor Cyan
       
    foreach($GRP in $DB_AZ_RSG_ALL){
        Add-AzTableRow `
            -UpdateExisting `
            -Table $OUT_DB_TBL_RSG.CloudTable `
            -PartitionKey $SUB.TenantId `
            -RowKey $GRP.ResourceGroupName `
            -Property @{
                "Location" = $GRP.Location;
                "ProvisioningState" = $GRP.ProvisioningState;
                "ResourceId" = $GRP.ResourceId;
                "ResourcesNumber" = ($DB_AZ_RES_ALL | Where-Object {$_.ResourceGroupName -eq $GRP.ResourceGroupName} | Measure-Object).Count;
                "SubscriptionId" = $SUB.SubscriptionId 
            } | Out-Null
    }
    Write-Host "        Se cargaron " $DB_AZ_RSG_ALL.Length " grupos de recursos exitosamente" -ForegroundColor DarkGreen

    #endregion información de grupos de recursos

    #region información general de recursos
    
    Write-Host "    F. Cargado de informacion general de recursos" -ForegroundColor Cyan
    $GBL_IN_FOR_CNT = 1
    foreach($RES in $DB_AZ_RES_ALL){
        $WR_BAR = ($RES.ResourceId.Substring($RES.ResourceId.IndexOf("providers")+10)).Replace("/",".").Replace(" ","_").Replace("#","_")
        If(!$RES.ParentResource){
            $PAR_RES = "Undefined"
        }
        else{
            $PAR_RES = $RES.ParentResource
        }
        Write-Progress -Activity "Cargando informacion" -status "Actualizando: $WR_BAR" -percentComplete ($GBL_IN_FOR_CNT / $DB_AZ_RES_ALL.Count*100) -ErrorAction SilentlyContinue -ParentId 100
        Add-AzTableRow `
            -UpdateExisting `
            -Table $OUT_DB_TBL_RES.CloudTable `
            -PartitionKey $SUB.TenantId `
            -RowKey $WR_BAR `
            -Property @{
                "Location" = $RES.Location;
                "ResourceType" = $RES.ResourceType;
                "ResourceId" = $RES.ResourceId.Substring($RES.ResourceId.IndexOf("resourceGroups")+15);
                "ParentResource" = $PAR_RES;
                "SubscriptionId" = $SUB.SubscriptionId;
                "ResourceGroupName" = $RES.ResourceGroupName
            } | Out-Null
        $GBL_IN_FOR_CNT++
    }
    Write-Host "        Se cargaron " $DB_AZ_RES_ALL.Length " recursos exitosamente" -ForegroundColor DarkGreen
    $GBL_IN_FOR_CNT = 1
    #endregion información general de recursos

    #region información recomendaciones de azure advisor
    
       Write-Host "    G. Cargado de informacion de recomendaciones de Azure Advisor" -ForegroundColor Cyan
       Write-Host "        Limpiando tabla de recomendaciones pasadas" -ForegroundColor DarkGreen
       $REC_OLD =  Get-AzTableRow -Table $OUT_DB_TBL_REC.CloudTable -PartitionKey $SUB.TenantId
       foreach($OLD in $REC_OLD){
            Remove-AzTableRow `
            -Table $OUT_DB_TBL_REC.CloudTable `
            -PartitionKey $OLD.PartitionKey `
            -RowKey $OLD.RowKey | Out-Null
       }
       $GBL_IN_FOR_CNT = 1
       foreach($REC in $DB_AZ_REC_ALL){ 
           $WR_BAR = $REC.Name
           $IM_FLD = $REC.ImpactedField
           if(!$IM_FLD){
            $IM_FLD = "Undefined"
            $IM_PRO = "Undefined"
           }
           else{
               $IM_PRO = $REC.ImpactedField.Substring(0,$REC.ImpactedField.IndexOf("."))
               $IM_TYP = $REC.ImpactedField.Substring($REC.ImpactedField.IndexOf(".")+1,($REC.ImpactedField.Substring($REC.ImpactedField.IndexOf(".")+1).IndexOf("/")))
               $IM_SBT = $REC.ImpactedField.Substring($REC.ImpactedField.IndexOf("/")+1)
           }

           Write-Progress -Activity "Cargando informacion" -status "Actualizando: $WR_BAR" -percentComplete ($GBL_IN_FOR_CNT / $DB_AZ_REC_ALL.Count*100) -ErrorAction SilentlyContinue -ParentId 100
           Add-AzTableRow `
               -UpdateExisting `
               -Table $OUT_DB_TBL_REC.CloudTable `
               -PartitionKey $SUB.TenantId `
               -RowKey $WR_BAR `
               -Property @{
                   "Problem" = $REC.ShortDescription.Problem;
                   "Solution" = $REC.ShortDescription.Solution;
                   "ImpactedValue" = $REC.ImpactedValue
                   "ImpactedField" = $IM_FLD;
                   "ImpactedProvider" = $IM_PRO;
                   "ImpactedType" = $IM_TYP;
                   "ImpactedSubType" = $IM_SBT;
                   "Impact" = $REC.Impact;
                   "LastUpdated" = $REC.LastUpdated.DateTime;
                   "SubscriptionId" = $SUB.SubscriptionId;
                   "Category" = $REC.Category
               } | Out-Null
           $GBL_IN_FOR_CNT++
       }
       Write-Host "        Se cargaron " $DB_AZ_REC_ALL.Length " recomendaciones exitosamente" -ForegroundColor DarkGreen
       $GBL_IN_FOR_CNT = 1
    #endregion información recomendaciones de azure advisor

    #region información de informacion de usuarios administrativos de Azure
        
    Write-Host "    H. Cargado de informacion de usuarios administrativos de Azure" -ForegroundColor Cyan
    $GBL_IN_FOR_CNT = 1
    foreach($USR in $DB_AZ_PER_ALL){
        $WR_BAR = $USR.DisplayName
        $FD_SCP = $USR.Scope
        $FD_SIG = $USR.SignInName
        $FD_ROW = $USR.RoleAssignmentId.Substring($USR.RoleAssignmentId.LastIndexOf("/")+1)
        if($USR.ObjectType -eq "User"){
            $FS_TYP = Get-AzureADUser -ObjectId $USR.ObjectId | Select-Object *
            if($FS_TYP.UserType){
                $FS_TYP =  $FS_TYP.UserType
            }
            elseif($FS_TYP.ImmutableId){
                $FS_TYP = "Member"
            }
            else{
                $FS_TYP = "Undefined"
            }
        }
        elseif ($USR.ObjectType -eq "ServicePrincipal") {
            $FS_TYP = (Get-AzureADServicePrincipal -ObjectId $USR.ObjectId | Select-Object ServicePrincipalType).ServicePrincipalType
        }
        else {
            $FS_TYP = $USR.ObjectType
        }
        
        if($FD_SCP -like "*providers*" ){
            $FD_SCP = "Resource"
        }
        elseif ($FD_SCP -like "*resourcegroups*") {
            $FD_SCP = "ResourceGroup"
        }
        elseif ($FD_SCP -like "*subscriptions*") {
            $FD_SCP = "Subscription"
        }
        else{
            $FD_SCP = "Root"
        }

        if(!$FD_SIG){
            $FD_SIG = "Undefined"
        }

        Write-Progress -Activity "Cargando informacion" -status "Actualizando: $WR_BAR" -percentComplete ($GBL_IN_FOR_CNT / $DB_AZ_PER_ALL.Count*100) -ErrorAction SilentlyContinue  -ParentId 100
        Add-AzTableRow `
            -UpdateExisting `
            -Table $OUT_DB_TBL_PER.CloudTable `
            -PartitionKey $SUB.TenantId `
            -RowKey $FD_ROW `
            -Property @{
                "RoleAssignmentId" = $USR.RoleAssignmentId;
                "Scope" = $USR.Scope;
                "ScopeLevel" = $FD_SCP;
                "DisplayName" = $USR.DisplayName;
                "SignInName" = $FD_SIG;
                "RoleDefinitionName" = $USR.RoleDefinitionName;
                "RoleDefinitionId" = $USR.RoleDefinitionId;
                "CanDelegate" = $USR.CanDelegate;
                "ObjectType" = $USR.ObjectType;
                "UserType" = $FS_TYP;
                "SubscriptionId" = $SUB.SubscriptionId
            } | Out-Null
        $GBL_IN_FOR_CNT++
        
    }
    Write-Host "        Se cargaron " $DB_AZ_PER_ALL.Length " usuarios administrativos exitosamente" -ForegroundColor DarkGreen
    $GBL_IN_FOR_CNT = 1
    #endregion información de informacion de usuarios administrativos de Azure

    #region informacion de Storages Accounts
        
    Write-Host "    I. Cargado de informacion de cuentas de almacenamiento" -ForegroundColor Cyan
    $DB_AZ_STO_ALL = $DB_AZ_RES_ALL | Where-Object {$_.ResourceType -eq 'Microsoft.Storage/storageAccounts'} | Select-Object *
    $OUT_DB_TBL_STO = Get-AzStorageTable -Context $OUT_TBL_CTX -Name ((($DB_AZ_STO_ALL | Select-Object ResourceType -Unique).ResourceType).ToLower().replace(".","").replace("/",""))
    $GBL_IN_FOR_CNT = 1
    foreach($STO in $DB_AZ_STO_ALL){
        $STO_INF = Get-AzStorageAccount -ResourceGroupName $STO.ResourceGroupName -Name $STO.Name
        $WR_BAR = $STO_INF.StorageAccountName
        Write-Progress -Activity "Cargando informacion" -status "Actualizando: $WR_BAR" -percentComplete ($GBL_IN_FOR_CNT / $DB_AZ_STO_ALL.Count*100) -ErrorAction SilentlyContinue  -ParentId 100
        if(!$STO_INF.CustomDomain){
            $CST_DMN = "Undefined"
        }
        else{
            $CST_DMN = $STO_INF.CustomDomain
        }
        if(!$STO_INF.LargeFileSharesState){
            $SHR_LRG = "Undefined"
        }
        else{
            $SHR_LRG = $STO_INF.LargeFileSharesState
        }
        If(!$STO_INF.NetworkRuleSet.IpRules){
            $RL_IPS = "Empty"
        }
        else{
            $RL_IPS = "Defined"
        }
        If(!$STO_INF.NetworkRuleSet.VirtualNetworkRules){
            $RL_IPS = "Empty"
        }
        else{
            $RL_IPS = ($STO_INF.NetworkRuleSet.VirtualNetworkRules | Measure-Object).Count
        }
        If(!$STO_INF.AccessTier){
            $ACC_TIR = "Undefined"
        }
        else{
            $ACC_TIR = $STO_INF.AccessTier.ToString()
        }
        Add-AzTableRow `
            -UpdateExisting `
            -Table $OUT_DB_TBL_STO.CloudTable `
            -PartitionKey $SUB.SubscriptionId `
            -RowKey $STO_INF.StorageAccountName `
            -Property @{
                "ResourceGroupName" = $STO_INF.ResourceGroupName;
                "Id" = $STO_INF.Id;
                "Location" = $STO_INF.Location;
                "SkuName" = $STO_INF.Sku.Name.ToString();
                "SkuTier" = $STO_INF.Sku.Tier.ToString();
                "Kind" = $STO_INF.Kind.ToString();
                "AccessTier" = $ACC_TIR;
                "CreationTime" = $STO_INF.CreationTime;
                "CustomDomain" = $CST_DMN;
                "PrimaryLocation" = $STO_INF.PrimaryLocation.ToString();
                "ProvisioningState" = $STO_INF.ProvisioningState.ToString();
                "StatusOfPrimary" = $STO_INF.StatusOfPrimary.ToString();
                "EnableHttpsTrafficOnly" = $STO_INF.EnableHttpsTrafficOnly;
                "LargeFileSharesState" = $SHR_LRG;
                "NetworkRuleSetBypass" = $STO_INF.NetworkRuleSet.Bypass.ToString();
                "NetworkRuleSetDefaultAction" = $STO_INF.NetworkRuleSet.DefaultAction.ToString();
                "NetworkRuleSetIpRules" = $RL_IPS;
                "NetworkRuleSetVirtualNetworkRules" = $RL_IPS;
                "TenantId" = $SUB.TenantId
            } | Out-Null
        $GBL_IN_FOR_CNT++

    }
    Write-Host "        Se cargaron " $DB_AZ_STO_ALL.Length " cuentas de almacenamiento exitosamente" -ForegroundColor DarkGreen
    $GBL_IN_FOR_CNT = 1
    #endregion informacion de Storages Accounts

    $GBL_IN_SUB_CNT++
}