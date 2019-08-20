#----------------------------------------------------------------------------------------------------------------------------
#############################################################################################################################
Set-ExecutionPolicy Unrestricted
#############################################################################################################################
$OneVariable = New-Object System.Management.Automation.Host.ChoiceDescription "&Dinámica", ` "Dinámica"
$TwoVariable = New-Object System.Management.Automation.Host.ChoiceDescription "&Estática", ` "Estática"
$LoadPrompt = [System.Management.Automation.Host.ChoiceDescription[]]($OneVariable, $TwoVariable)
$Prompt = $host.ui.PromptForChoice("Confirmación", "¿Qué tipo de conexión deseas?", $LoadPrompt, 0) 
switch ($Prompt){
#############################################################################################################################
#####################################              Primera Opción              ##############################################
#############################################################################################################################
0 {
#############################################################################################################################
$IPType = "IPv4"
$adapter = Get-NetAdapter | ? {$_.Status -eq "up"}
$interface = $adapter | Get-NetIPInterface -AddressFamily $IPType
If ($interface.Dhcp -eq "Disabled") {
    # Remove existing gateway
    If (($interface | Get-NetIPConfiguration).Ipv4DefaultGateway) {
        $interface | Remove-NetRoute -Confirm:$false
    }

    # Enable DHCP
    $interface | Set-NetIPInterface -DHCP Enabled

    # Configure the  DNS Servers automatically
    $interface | Set-DnsClientServerAddress -ResetServerAddresses
}
ipconfig /renew
#############################################################################################################################                      
}
#############################################################################################################################
#####################################              Segunda Opción              ##############################################
#############################################################################################################################
1 {
$PersonalIP = Read-Host "Cual es la IP que te corresponde?"
$IP= "172.16.0." + $PersonalIP
#############################################################################################################################
$Claro = New-Object System.Management.Automation.Host.ChoiceDescription "&Claro", ` "Claro"
$Movistar = New-Object System.Management.Automation.Host.ChoiceDescription "&Movistar", ` "Movistar"
$LoadConType = [System.Management.Automation.Host.ChoiceDescription[]]($Claro, $Movistar)
$ConType = $host.ui.PromptForChoice("Confirmación", "¿Qué tipo de conexión deseas?", $LoadConType, 0) 
switch ($ConType){
0{
 #############################################################################################################################
        $MaskBits = 24 # This means subnet mask = 255.255.255.0
        $Gateway = "172.16.0.1"
        $Dns = "172.16.0.240"
        $IPType = "IPv4"

        # Retrieve the network adapter that you want to configure
        $adapter = Get-NetAdapter | ? {$_.Status -eq "up"}

        # Remove any existing IP, gateway from our ipv4 adapter
        If (($adapter | Get-NetIPConfiguration).IPv4Address.IPAddress) {
            $adapter | Remove-NetIPAddress -AddressFamily $IPType -Confirm:$false
        }

        If (($adapter | Get-NetIPConfiguration).Ipv4DefaultGateway) {
            $adapter | Remove-NetRoute -AddressFamily $IPType -Confirm:$false
        }

         # Configure the IP address and default gateway
        $adapter | New-NetIPAddress `
            -AddressFamily $IPType `
            -IPAddress $IP `
            -PrefixLength $MaskBits `
            -DefaultGateway $Gateway

        # Configure the DNS client server IP addresses
        $adapter | Set-DnsClientServerAddress -ServerAddresses $DNS
 #############################################################################################################################
}
1{
 #############################################################################################################################
        $MaskBits = 24 # This means subnet mask = 255.255.255.0
        $Gateway = "172.16.0.246"
        $Dns = "172.16.0.240"
        $IPType = "IPv4"

        # Retrieve the network adapter that you want to configure
        $adapter = Get-NetAdapter | ? {$_.Status -eq "up"}

        # Remove any existing IP, gateway from our ipv4 adapter
        If (($adapter | Get-NetIPConfiguration).IPv4Address.IPAddress) {
            $adapter | Remove-NetIPAddress -AddressFamily $IPType -Confirm:$false
        }

        If (($adapter | Get-NetIPConfiguration).Ipv4DefaultGateway) {
            $adapter | Remove-NetRoute -AddressFamily $IPType -Confirm:$false
        }

         # Configure the IP address and default gateway
        $adapter | New-NetIPAddress `
            -AddressFamily $IPType `
            -IPAddress $IP `
            -PrefixLength $MaskBits `
            -DefaultGateway $Gateway

        # Configure the DNS client server IP addresses
        $adapter | Set-DnsClientServerAddress -ServerAddresses $DNS
 #############################################################################################################################
}
}
#############################################################################################################################
}
#############################################################################################################################
}
#----------------------------------------------------------------------------------------------------------------------------