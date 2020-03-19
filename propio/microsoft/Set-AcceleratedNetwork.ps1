$nic = Get-AzureRmNetworkInterface -ResourceGroupName “rg-hana-prod-001” -Name “vmhana01-secondary”

$nic.EnableAcceleratedNetworking = $false

$nic | Set-AzureRmNetworkInterface