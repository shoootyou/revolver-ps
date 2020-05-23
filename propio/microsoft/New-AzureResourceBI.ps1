<################################################################################################

Author: Rodolfo Castelo Méndez
Versión: 1.0
Required Modules:
    AzTable
    Az

################################################################################################>

<#region de inicio de sesión y datos

Connect-AzAccount
$TNT_ID = "fca6d03e-0144-4abb-9215-05ebbce29cb0"
$COR_AZ_TNT_ALL = Connect-AzureAD -TenantId $TNT_ID
$OUT_TBL_CNN = "DefaultEndpointsProtocol=https;AccountName=azrsrcbi001;AccountKey=YrHybuuJJJiskDFLcjGXjR/4s+44b0fA5lo0/xj+GFXQoBjd55dgET0KkaLLC06bL7tIWQq8QthmhpC+EoJCXQ==;EndpointSuffix=core.windows.net"
$OUT_TBL_CTX = New-AzStorageContext -ConnectionString $OUT_TBL_CNN
$GBL_AZ_SUB_CNT = 1
$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"

region obtención de informacion base #>

#region obtencion de recursos de subscripción
Clear-Host
$COR_AZ_SUB_ALL = Get-AzSubscription -TenantId $TNT_ID | Select-Object *

#endregion obtencion de recursos de subscripción

#region de preparación de tablas maestras
Write-Host "0 - Creacion de tablas maestras de recursos" -ForegroundColor DarkGray
$OUT_DB_TBL_ACC =  Get-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterresourcestableurl"
if(!$OUT_DB_TBL_ACC){
    $OUT_DB_TBL_ACC = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterresourcestableurl"
}
Write-Host "        Tabla de acceso a los recursos creada exitosamente" -ForegroundColor Green

$OUT_DB_TBL_SUB =  Get-AzStorageTable -Context $OUT_TBL_CTX -Name "amastersubscription"
if(!$OUT_DB_TBL_SUB){
    $OUT_DB_TBL_SUB = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amastersubscription"
}
Write-Host "        Tabla de listado de suscripciones creada exitosamente" -ForegroundColor Green

$OUT_DB_TBL_RSG =  Get-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterresourcegroup"
if(!$OUT_DB_TBL_RSG){
    $OUT_DB_TBL_RSG = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterresourcegroup"
}
Write-Host "        Tabla de listado de grupo de recursos creada exitosamente" -ForegroundColor Green

$OUT_DB_TBL_REG =  Get-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterregions"
if(!$OUT_DB_TBL_REG){
    $OUT_DB_TBL_REG = New-AzStorageTable -Context $OUT_TBL_CTX -Name "amasterregions"
}
Write-Host "        Tabla de listado de regiones creada exitosamente" -ForegroundColor Green

#endregion de preparación de tablas maestras

foreach($SUB in $COR_AZ_SUB_ALL){
    Write-Host $GBL_AZ_SUB_CNT "- Inicializacion de datos para subscripcion" $SUB.SubscriptionId -ForegroundColor DarkGray
    #region selección de subscripción y recursos
    Write-Host "    A. Obtencion de informacion de subscripciones y recursos" -ForegroundColor Cyan
    Select-AzSubscription -Subscription $SUB.SubscriptionId | Out-Null
    $DB_AZ_RES_ALL = Get-AzResource | Select-Object * | Sort-Object Type
    $DB_AZ_RSG_ALL = Get-AzResourceGroup | Select-Object * | Sort-Object Type
    Write-Host "        Suscription ID   :" $SUB.SubscriptionId -ForegroundColor Green
    Write-Host "        Suscription Name :" $SUB.Name -ForegroundColor Green
    #endregion selección de subscripción y recursos

    #region provisionamiento de tablas de acceso de recursos
    Write-Host "    B. Creacion de tablas para informacion de recursos" -ForegroundColor Cyan
    $DB_AZ_RES_TYP  = @()
    ($DB_AZ_RES_ALL | Select-Object Type -Unique) | ForEach-Object { $DB_AZ_RES_TYP += $_.Type.Substring(0,$_.Type.IndexOf("/"))}
    $DB_AZ_RES_TYP = $DB_AZ_RES_TYP | Select-Object -Unique

    foreach($RES in $DB_AZ_RES_TYP){
    
        $FOR_INT_00 = (($RES).ToLower()).replace(".","")
        $FOR_INT_01 = Get-AzStorageTable -Context $OUT_TBL_CTX -Name $FOR_INT_00 -ErrorAction SilentlyContinue
        if(!$FOR_INT_01){
            $FOR_INT_01 = New-AzStorageTable -Context $OUT_TBL_CTX -Name $FOR_INT_00
        }
        
        Add-AzTableRow `
            -UpdateExisting `
            -Table $OUT_DB_TBL_ACC.CloudTable `
            -PartitionKey $FOR_INT_00 `
            -RowKey $FOR_INT_00 `
            -Property @{
                "Uri" = $FOR_INT_01.Uri.AbsoluteUri;
                "Context" = $FOR_INT_01.Context.ConnectionString
            } | Out-Null

        $FOR_INT_00 = $null
        $FOR_INT_01 = $null

    }
    Write-Host "        Se crearon " $DB_AZ_RES_TYP.Length "tablas de forma exitosa" -ForegroundColor DarkGreen
    Write-Host "        Se cargo la informacion de acceso de" $DB_AZ_RES_TYP.Length "tablas de forma exitosa" -ForegroundColor DarkGreen
    

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
            "ResourceGroup" = $DB_AZ_RSG_ALL.Count
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
            -PartitionKey $SUB.TenantId `
            -RowKey $REG.Location `
            -Property @{
                "ResourcesNumber" = ($DB_AZ_RES_ALL | Where-Object {$_.Location -eq $REG.Location} | Measure-Object).Count
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
                "ResourceId" = $GRP.ResourceId
                "ResourcesNumber" = ($DB_AZ_RES_ALL | Where-Object {$_.ResourceGroupName -eq $GRP.ResourceGroupName} | Measure-Object).Count
            } | Out-Null
    }
    Write-Host "        Se cargaron " $DB_AZ_RSG_ALL.Length " grupos de recursos exitosamente" -ForegroundColor DarkGreen

    #endregion información de grupos de recursos

}
