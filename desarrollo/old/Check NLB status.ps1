#Define Nodes 
$nodeserver1='Server1' 
$nodeserver2='Server2' 
 
#get NLB status on NLB Nodes  
$Node1 = Get-WmiObject -Class MicrosoftNLB_Node -computername $nodeserver1 -namespace root\MicrosoftNLB | Select-Object __Server, statuscode 
if($node1.statuscode -eq "1008" -or $node1.statuscode -eq "1007"){ 
    write-host "NLB Status of $node1 is: Converged"   -fore green
}else{ 
    write-host "NLB Status of $node1 is: Error" -fore red
} 
$Node2 = Get-WmiObject -Class MicrosoftNLB_Node -computername $nodeserver2 -namespace root\MicrosoftNLB | Select-Object __Server, statuscode 
if($node2.statuscode -eq "1008" -or $node2.statuscode -eq "1007"){ 
    write-host "NLB Status of $node2 is: Converged"   -fore green
}else{ 
     write-host "NLB Status of $node2 is: Error"  -fore red
}