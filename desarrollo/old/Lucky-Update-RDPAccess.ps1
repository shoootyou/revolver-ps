Login-AzureRmAccount -SubscriptionId '2f5687a8-4b58-4212-92ca-b538edc10e2e'

$URI = “https://www.cual-es-mi-ip.net/ “
$HTML = Invoke-WebRequest -Uri $URI
$IP = ($HTML.ParsedHtml.getElementsByTagName("span") | Where{ $_.className -eq 'big-text font-arial' } ).textContent

Write-Host '------------------------------------------------------' -ForegroundColor Green

$DB_NSG = Get-AzureRmNetworkSecurityGroup
foreach($NSG in $DB_NSG){
    $RDP_RL = $NSG | Select SecurityRules -ExpandProperty SecurityRules | Where {$_.Name -eq 'Torioux-RDP'}

    Set-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG `
    -Name $RDP_RL.Name `
    -Access $RDP_RL.Access `
    -Protocol $RDP_RL.Protocol `
    -Direction $RDP_RL.Direction `
    -Priority $RDP_RL.Priority `
    -SourceAddressPrefix $IP `
    -SourcePortRange $RDP_RL.SourcePortRange `
    -DestinationAddressPrefix $RDP_RL.DestinationAddressPrefix `
    -DestinationPortRange $RDP_RL.DestinationPortRange | Out-Null

     Set-AzureRmNetworkSecurityGroup -NetworkSecurityGroup $NSG | Out-Null

     Write-Host 'Change for ' $NSG.Name " it's completed" -ForegroundColor Green
     Write-Host '------------------------------------------------------' -ForegroundColor Green
}
pause