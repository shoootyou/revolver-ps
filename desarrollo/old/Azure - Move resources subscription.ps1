<#$resource = Get-AzureRmResource -ResourceName ExampleApp -ResourceGroupName OldRG
Move-AzureRmResource -DestinationResourceGroupName NewRG -ResourceId $resource.ResourceId -DestinationSubscriptionId

$webapp = Get-AzureRmResource -ResourceGroupName BKL-Metropolis -ResourceName ExampleSite
$plan = Get-AzureRmResource -ResourceGroupName OldRG -ResourceName ExamplePlan
Move-AzureRmResource -DestinationResourceGroupName NewRG -ResourceId $webapp.ResourceId, $plan.ResourceId
#>

$DB_RESO = Get-AzureRmResource | Where {$_.Name -like '*cube*'} 
$new_subs_id = 'c0724ee0-e9e2-48a5-a3a5-3acf854f27ad'
$new_reso_gr = 'Americatel-CubeBackup'
foreach($Reso in $DB_RESO){
    $resource = Get-AzureRmResource -ResourceName $Reso.ResourceName -ResourceGroupName $Reso.ResourceGroupName
    Move-AzureRmResource -DestinationResourceGroupName $new_reso_gr -ResourceId $Reso.ResourceId -DestinationSubscriptionId $new_subs_id -Force
}



