#$sourcevhdName = "CubeBackup-Data.vhd"
#$sourcevhdName = "CubeBackup03.vhd"
$sourcevhdName = "TempCube-03.vhd"

#$destinationvhdname = "Compilado-01.vhd"
#$destinationvhdname = "Compilado-02.vhd"
$destinationvhdname = "Compilado-03.vhd"
<#
$sourceSAName = "cubebackupsysdisk01"
$sourceSAKey = "BKpOHLoldIpI9NzHb8Ae7wl5yyz8Z0Kxt9XqPrYoQMtyjBfdO81WuIUWLfhlgArnT0kfUHepVqMlOiOWVJZXsw=="
$sourceSAContainerName = "vhds"

$sourceSAName = "cubebackupsysdisk02"
$sourceSAKey = "92hedu4HcIzdb8LkWWE8W6l+TPLT09N/SLMRwIi4T43htbEDzHhXeE65hnGhqF2rsjurkVbqj3xxVyd4+gWZNQ=="
$sourceSAContainerName = "vhds"
#>
$sourceSAName = "cubebackupsysdisk03"
$sourceSAKey = "dlBcKmXskwpIvGpuqW48PzrD9xT3cgjv5vdfwpJb2Q3G3rxqYZnar2+LEBqXGLLzdwWKVy36YcBeF28wdWrW3Q=="
$sourceSAContainerName = "vhds"

$destinationSAName = "cubeconsolidados"
$destinationSAKey = "ZHZfc/MRWsDHsXiA+9JQT3Ovs9Qwj01uEu2D+clCO85XsTHxc3zoTODHLc/bV2HSI1J27jMxu2ISg3E/UAc4ew=="
$destinationContainerName = "cubeconsolida"

$sourceContext = New-AzureStorageContext -StorageAccountName $sourceSAName -StorageAccountKey $sourceSAKey
$destinationContext = New-AzureStorageContext –StorageAccountName $destinationSAName -StorageAccountKey $destinationSAKey

$blobCopy = Start-AzureStorageBlobCopy -DestContainer $destinationContainerName -DestContext $destinationContext -SrcBlob $sourcevhdName -Context $sourceContext -SrcContainer $sourceSAContainerName -DestBlob $destinationvhdname

($blobCopy | Get-AzureStorageBlobCopyState).Status

#$blobCopy | Stop-AzureStorageBlobCopy

Get-AzureStorageBlobCopyState -Blob "cubeconsolidados" -Container 'vhds'



$destinationvhdname |  Get-AzureStorageContainer 
Get-AzureStorageBlob -Context $sourceContext


$AnonContext = New-AzureStorageContext -StorageAccountName $destinationSAName  -Anonymous ;
Get-AzureStorageBlob -Context $AnonContext -Container $ContainerName;

Get-AzureStorageBlobCopyState -Blob CubeBackup-Data.vhd -Container vhds

Get-AzureStorageBlob -Context $destinationContext -Container $destinationContainerName | Get-AzureStorageBlobCopyState

Stop-AzureStorageBlobCopy -Context $sourceContext -Container $sourceSAContainerName -CopyId 'f38535bf-88d9-4c14-9b57-ca0a22c14a54' -Blob "CubeBackup03.vhd"