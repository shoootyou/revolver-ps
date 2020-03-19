$resourceGroupName = "rg-hana-qas-100"
$HannSnap = Get-AzSnapshot -ResourceGroupName $resourceGroupName
$SKU = "Premium_LRS"
$vmtargetname = "vmhanaqas100"
foreach($DiskInt in $HannSnap){
    
    if($DiskInt.OsType){
        $diskConfig = New-AzDiskConfig -SkuName $SKU -Location $DiskInt.Location -CreateOption Copy `
        -SourceResourceId $DiskInt.Id -DiskSizeGB $DiskInt.DiskSizeGB -Zone 1 `
        -OsType $DiskInt.OsType 
        New-AzDisk -Disk $diskConfig -ResourceGroupName $resourceGroupName -DiskName ($DiskInt.Name.Replace("vmhana01",$vmtargetname)).replace("-snap","")
        ($DiskInt.Name.Replace("vmhana01",$vmtargetname)).replace("-snap","")
        
    }
    else{
        $diskConfig = New-AzDiskConfig -SkuName $SKU -Location $DiskInt.Location -CreateOption Copy `
        -SourceResourceId $DiskInt.Id -DiskSizeGB $DiskInt.DiskSizeGB -Zone 1
        New-AzDisk -Disk $diskConfig -ResourceGroupName $resourceGroupName -DiskName ($DiskInt.Name.Replace("vmhana01",$vmtargetname)).replace("-snap","") 
        ($DiskInt.Name.Replace("vmhana01",$vmtargetname)).replace("-snap","")
    }

}

 
