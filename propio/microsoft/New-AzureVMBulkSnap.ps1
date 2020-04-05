$resourceGroupName = "RG-HANA-EASTUS2-QA-001"
$resourceGroupNameTarget = "rg-hana-qas-100"
$vmName="QAS-hana-vm"
$Location = "EastUS2"
$i  = 0

# Obtencion SO
$vm = get-azureRmVm -ResourceGroupName $resourceGroupName -Name $vmName

#Captura Disco 
$vmOSDisk = (Get-AzureRmVM -ResourceGroupName $resourceGroupName -Name $vmName).StorageProfile.OsDisk.Name
$Disk = Get-AzureRmDisk -ResourceGroupName $resourceGroupName -DiskName $vmOSDisk
Write-Host "Iniciando Snap de SO"
#Snap SO
$SnapshotConfig = New-AzureRmSnapshotConfig -SourceUri $Disk.Id -CreateOption Copy -Location $Location
$Snapshot = New-AzureRmSnapshot -Snapshot $snapshotConfig -SnapshotName ($vmOSDisk + "-snap") -ResourceGroupName $resourceGroupNameTarget
Write-Host "Finalizado snapshot SO"
#Snap Data Disk
$vmdataDisk = (Get-AzureRmVM -ResourceGroupName $resourceGroupName -Name $vmName).StorageProfile.DataDisks
foreach($diskData in $vmdataDisk){
    Write-Progress -Id 1 -Activity "Generando snapshots" -status ("Trabajando en " + $diskData.Name) -percentComplete ($i / $vmdataDisk.count*100)
    #Obtencion Disk Data
    $datadisk_int = Get-AzureRmDisk -ResourceGroupName $resourceGroupName -DiskName $diskData.Name

    #Snap SO
    $SnapshotConfig = New-AzureRmSnapshotConfig -SourceUri $datadisk_int.Id -CreateOption Copy -Location $Location
    $Snapshot = New-AzureRmSnapshot -Snapshot $snapshotConfig -SnapshotName ($datadisk_int.Name + "-snap") -ResourceGroupName $resourceGroupNameTarget
    $i++
}