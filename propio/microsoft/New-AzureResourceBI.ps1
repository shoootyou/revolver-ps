<#region de inicio de sesión y datos

Connect-AzAccount 
$TNT_ID = "fca6d03e-0144-4abb-9215-05ebbce29cb0"
$OUT_TBL_CNN = "DefaultEndpointsProtocol=https;AccountName=azrsrcbi001;AccountKey=YrHybuuJJJiskDFLcjGXjR/4s+44b0fA5lo0/xj+GFXQoBjd55dgET0KkaLLC06bL7tIWQq8QthmhpC+EoJCXQ==;EndpointSuffix=core.windows.net"
$OUT_TBL_CTX = New-AzStorageContext -ConnectionString $OUT_TBL_CNN
$OUT_INT_TBL_NAM = "resourcestableurl"
$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"

region obtención de información base #>

#region obtencion de recursos de subscripción

$COR_AZ_SUB_ALL = Get-AzSubscription -TenantId $TNT_ID | Select-Object *

#endregion obtencion de recursos de subscripción

foreach($SUB in $COR_AZ_SUB_ALL){
    #region selección de subscripción y recursos
    
    Select-AzSubscription -Subscription $SUB.SubscriptionId
    $DB_AZ_SUB_INF = Get-AzSubscription -SubscriptionId $SUB.SubscriptionId | Select-Object *
    $DB_AZ_RES_ALL = Get-AzResource | Select-Object * | Sort-Object Type

    #endregion selección de subscripción y recursos

    #region de preparación de tabla de accesos a tabla de recursos

    $OUT_DB_TBL_ACC =  Get-AzStorageTable -Context $OUT_TBL_CTX -Name $OUT_INT_TBL_NAM  # Tabla general de acceso a tabla de recursos
    if(!$OUT_DB_TBL_ACC){
        $OUT_DB_TBL_ACC = New-AzStorageTable -Context $OUT_TBL_CTX -Name $OUT_INT_TBL_NAM  # Tabla general de acceso a tabla de recursos
    }
    
    #endregion de preparación de tabla de accesos a tabla de recursos
    
    #region provisionamiento de tablas de recursos
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

        $FOR_INT_01 = $null
    }
    #endregion provisionamiento de tablas de recursos
    

}


