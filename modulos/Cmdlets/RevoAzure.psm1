﻿<#MIT License

Copyright (c) 2017 Rodolfo Castelo Méndez

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#
# Generated by: Rodolfo Castelo Méndez
#
# Generated on: 10/31/2017
#>

Function Register-RevoProvider {
    Param(
        [string]$ResourceProviderNamespace
    )
    process{
        Write-RevoMessageConsole ("Registering resource provider '$ResourceProviderNamespace'") -Type Verbose;
        Register-AzureRmResourceProvider -ProviderNamespace $ResourceProviderNamespace;
    }
}

Function Get-RevoResourceGroup{
    <#
        .SYNOPSIS
        Permite validar si el Grupo de Recursos proporcionado existe.
        
        .DESCRIPTION
        Permite validar si el Grupo de Recursos proporcionado existe, en caso no existe, procede con su creación en la región especificada.
                
        .EXAMPLE
        Get-RevoResourceGroup -ResourceGroupName 'migrupoderecursos' -Region 'East US' -Deployment

        Validará si el grupo de Recursos indicado existe, de no existirlo, procederá con su creación. El switch Deployment generará mensajes sobre la creación correcta.
        
        .EXAMPLE
        Get-RevoResourceGroup -ResourceGroupName 'migrupoderecursos' -Region 'East US'

        Validará si el grupo de Recursos indicado existe, de no existirlo, procederá con su creación. 
        
        .PARAMETER ResourceGroupName
        Tipo: Requerido
        Nombre del Grupo de Recursos a Utilizar
      
        .PARAMETER Region
        Tipo: Requerido
        Región en la que se crearán los recursos

        .PARAMETER Deployment      
        Tipo: Opcional
        Validador de si se empleará la función para un despliegue o sólo validación.

        .PARAMETER Silent      
        Tipo: Opcional
        Validador de si se retornará algún mensaje de error.
    #>
    [CmdletBinding(DefaultParametersetName='None')] 
    param(
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ResourceGroupName,
        [Parameter(Mandatory=$True)]
        [ValidateSet('South Central US','North Europe','West Europe','Southeast Asia','Korea Central','Korea South','West US','East US','Japan West','Japan East','East Asia','East US 2','North Central US','Central US','Brazil South','Australia East','Australia Southeast','West India','Central India','South India','Canada Central','Canada East','West Central US','UK West','UK South','West US 2')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Region,
        [Parameter(Mandatory=$false)]      
        [switch]
        $Deployment,
        [Parameter(Mandatory=$false)]      
        [switch]
        $Silent

    )
    begin{
        
        $ErrorActionPreference = "SilentlyContinue"

    }
    process{

        $BEG_INT_RG = Get-AzureRMResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue;

        if(!$BEG_INT_RG -and $Deployment -and !$Silent){
            Write-RevoMessageConsole 'El grupo de recursos proporcionado no existe, se procederá con su creación.' -Type Warning;
            $BEG_INT_RG = New-AzureRmResourceGroup -Name $ResourceGroupName -Location $Region;
            Write-RevoMessageConsole ('El grupo de recursos ' +$ResourceGroupName +' fue creado de forma exitosa.') -Type Confirmation;
            return $BEG_INT_RG
        }
        elseif(!$BEG_INT_RG -and $Deployment -and $Silent){
            $BEG_INT_RG = New-AzureRmResourceGroup -Name $ResourceGroupName -Location $Region;
            return $BEG_INT_RG
        }
        elseif(!$BEG_INT_RG -and !$Silent){
            Write-RevoMessageConsole ('El grupo de recursos ' +$ResourceGroupName +' no existe. Emplee el parámetro -Deployment para proceder con su creación automáticamente.') -Type Confirmation;
            return $null
        }
        elseif(!$BEG_INT_RG -and $Silent){
            return $null
        }

        else{
            if($Deployment){
                Write-RevoMessageConsole ('El grupo de recursos ' +$ResourceGroupName +' existe y se empleará en el proceso de despliegue.') -Type Confirmation;
                return $BEG_INT_RG
            }
            else{
                Write-RevoMessageConsole ('El grupo de recursos ' +$ResourceGroupName +' existe.') -Type Confirmation;
                return $BEG_INT_RG
            }
            
        }

    }
}

Function New-RevoAppService{
    <#
        .SYNOPSIS
        Permite crear un nuevo App Service con App Service Plan Incluído
        
        .DESCRIPTION
        Permite crear un nuevo App Service con App Service Plan incluído además de la posibilidad de emplear un Grupo de Recursos existente o crear uno nuevo. En caso existe el App Service Plan con el nombre indicado, se procederá con el uso del mismo.
                
        .EXAMPLE
        New-RevoAppService -WebsiteName 'examplewebsite' -ResourceGroupName 'Example-ResourceGroup' -AppServiceTier Standard -Region 'West US' -ServicePlanName 'ExampleAppServicePlan' -UseNewServicePlan 

        Creará un app service nuevo con el nombre 'examplewebsite' con los valores indicados. El swtich UseNewServicePlan generará la creación del App Service Plan de nombre 'ExampleAppServicePlan'
        
        .EXAMPLE
        New-RevoAppService -WebsiteName 'examplewebsite' -ResourceGroupName 'Example-ResourceGroup' -AppServiceTier Standard -Region 'West US' -ServicePlanName 'ExampleAppServicePlan' -UseExistingServicePlan 

        Creará un app service nuevo con el nombre 'examplewebsite' con los valores indicados. El swtich UseExistingServicePlan buscará y empleará el App Service Plan de nombre 'ExampleAppServicePlan'
        
        .PARAMETER WebsiteName
        Tipo: Requerido
        Nombre del App Service a crear

        .PARAMETER ResourceGroupName
        Tipo: Requerido
        Nombre del Grupo de Recursos a Utilizar

        .PARAMETER AppServiceTier
        Tipo: Requerido
        Tier empleado por el App Service Plan 
        
        .PARAMETER Region
        Tipo: Requerido
        Región en la que se crearán los recursos

        .PARAMETER ServicePlanName
        Tipo: Requerido only if UseExistingServicePlan is TRUE
        Nombre del App Service Plan que se empleará o creará

        .PARAMETER UseExistingServicePlan
        Tipo: Opcional
        Validador de si se usará un App Service Plan existente

        .PARAMETER UseNewServicePlan      
        Tipo: Opcional
        Validador de si se procederá con la creación de un App Service Plan nuevo

    #>
    [CmdletBinding(DefaultParametersetName='None')] 
    param(
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string]
        $WebsiteName,
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ResourceGroupName,
        [Parameter(Mandatory=$True)]
        [ValidateSet('Free','Shared','Basic','Standard','Premium','PremiumV2')]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppServiceTier,
        [Parameter(Mandatory=$True)]
        [ValidateSet('South Central US','North Europe','West Europe','Southeast Asia','Korea Central','Korea South','West US','East US','Japan West','Japan East','East Asia','East US 2','North Central US','Central US','Brazil South','Australia East','Australia Southeast','West India','Central India','South India','Canada Central','Canada East','West Central US','UK West','UK South','West US 2')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Region,
        [Parameter(Mandatory=$False)]
        [Parameter(ParameterSetName='Existing',Mandatory=$true)]
        [Parameter(ParameterSetName='New',Mandatory=$false)]
        [string]
        $ServicePlanName,
        [Parameter(ParameterSetName='Existing')]
        [switch]
        $UseExistingServicePlan,
        [Parameter(ParameterSetName='New')]
        [switch]
        $UseNewServicePlan      
    )
    begin{
        $ErrorActionPreference = "Stop"
       
        # Register RPs
        $resourceProviders = @("microsoft.web");
        if($resourceProviders.length) {
            foreach($resourceProvider in $resourceProviders) {
                Register-RevoProvider($resourceProvider) | Out-Null;
            }
        }

        if($WebsiteName -match ' '){
            Write-RevoMessageConsole 'El nombre del sitio no puede poseer espacios, por favor reintente nuevamente sin espacios.' -Type Error
            break
        }

    }
    process{

        <#
            $BEG_INT_RG    | Información del Grupo de Recursos a utilizar.
            $BEG_APP_SVP   | Información del Service Plan a utilizar.
            $BEG_APP_TIER  | Información del Plan escogido para App Service
            $BEG_WEB_NAME  | Información del nombre del AppService
        
        #>
        
        $BEG_APP_TIER = $AppServiceTier;

        $BEG_INT_RG = Get-RevoResourceGroup -ResourceGroupName $ResourceGroupName -Region $Region -Deployment

        if(!$ServicePlanName){
            $ServicePlanName = $WebsiteName.ToLower()
        }

        if($UseExistingServicePlan -and $ServicePlanName){
            $BEG_APP_SVP = Get-AzureRmAppServicePlan -Name $ServicePlanName -ErrorAction Stop;
        }
        elseif(($UseNewServicePlan -and $ServicePlanName) -or (!$BEG_APP_SVP)){
            Write-RevoMessageConsole ('Se procederá con la creación de App Service Plan '+ $ServicePlanName) -Type Verbose;
            $BEG_APP_SVP = New-AzureRmAppServicePlan -Name $ServicePlanName -Location $Region -ResourceGroupName $BEG_INT_RG.ResourceGroupName -Tier $BEG_APP_TIER;
            $BEG_APP_SVP = Get-AzureRmAppServicePlan -Name $ServicePlanName -ErrorAction Stop;
            Write-RevoMessageConsole ('El App Service Plan ' +$ServicePlanName +' fue creado de forma exitosa.') -Type Confirmation;
        }

        $BEG_WEB_NAME = $WebsiteName.ToLower()

        Write-RevoMessageConsole ('Se procederá con la creación de App Service '+ $BEG_WEB_NAME) -Type Verbose;
        $RTRN_FIN = New-AzureRmWebApp -ResourceGroupName $BEG_INT_RG.ResourceGroupName -Name $BEG_WEB_NAME -Location $Region -AppServicePlan $BEG_APP_SVP.Id 
        Write-RevoMessageConsole ('El App Service ' +$BEG_WEB_NAME +' fue creado de forma exitosa.') -Type Confirmation;

        return $RTRN_FIN
    }
}

Function Get-RevoVirtualNetworkSubnet{
    <#
        .SYNOPSIS
        Permite validar y crear una subred en una Virtual Network.
        
        .DESCRIPTION
        Permite validar y crear una subred en una Virtual Network, de no existir la red Virtual se procederá con su creación.
                
        .EXAMPLE
        New-RevoAppService -WebsiteName 'examplewebsite' -ResourceGroupName 'Example-ResourceGroup' -AppServiceTier Standard -Region 'West US' -ServicePlanName 'ExampleAppServicePlan' -UseNewServicePlan 

        Creará un app service nuevo con el nombre 'examplewebsite' con los valores indicados. El swtich UseNewServicePlan generará la creación del App Service Plan de nombre 'ExampleAppServicePlan'
        
        .EXAMPLE
        New-RevoAppService -WebsiteName 'examplewebsite' -ResourceGroupName 'Example-ResourceGroup' -AppServiceTier Standard -Region 'West US' -ServicePlanName 'ExampleAppServicePlan' -UseExistingServicePlan 

        Creará un app service nuevo con el nombre 'examplewebsite' con los valores indicados. El swtich UseExistingServicePlan buscará y empleará el App Service Plan de nombre 'ExampleAppServicePlan'

        .PARAMETER ResourceGroupName
        Tipo: Requerido
        Nombre del Grupo de Recursos a Utilizar

        .PARAMETER Region
        Tipo: Requerido
        Región en la que se crearán los recursos

        .PARAMETER VNETName
        Tipo: Requerido
        Nombre de la Red Virtual a utilizar y/o crear.

        .PARAMETER UseExistingVNET
        Tipo: Opcional
        Validador de si se usará una red virtual existente en Azure.

        .PARAMETER UseNewVNET
        Tipo: Opcional
        Validador de si se usará una nueva red virtual existente en Azure la cual deberá ser creada en el proceso.

        .PARAMETER VNETGlobalAddressPrefix      
        Tipo: Requerido si se utiliza UseNewVNET
        Campo que debe contener la red principal de la Red Virtual. Ej.: 10.0.0.0/16

        .PARAMETER VNETSubnetAddressPrefix      
        Tipo: Requerido
        Campo que debe contener la red de la subnet a utilizar. Ej.: 10.0.0.0/28

        .PARAMETER Deployment      
        Tipo: Opcional
        Validador de si se empleará la función para un despliegue o sólo validación.

    #>
    [CmdletBinding(DefaultParametersetName='None')] 
    param(
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ResourceGroupName,
        [Parameter(Mandatory=$True)]
        [ValidateSet('South Central US','North Europe','West Europe','Southeast Asia','Korea Central','Korea South','West US','East US','Japan West','Japan East','East Asia','East US 2','North Central US','Central US','Brazil South','Australia East','Australia Southeast','West India','Central India','South India','Canada Central','Canada East','West Central US','UK West','UK South','West US 2')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Region,
        [Parameter(Mandatory=$True)]
        [Parameter(ParameterSetName='Existing',Mandatory=$true)]
        [Parameter(ParameterSetName='New',Mandatory=$true)]
        [string]
        $VNETName,
        [Parameter(ParameterSetName='Existing')]
        [switch]
        $UseExistingVNET,
        [Parameter(ParameterSetName='New')]
        [switch]
        $UseNewVNET,
        [Parameter(ParameterSetName='New',Mandatory=$True)]
        [string]
        $VNETGlobalAddressPrefix,
        [Parameter(ParameterSetName='New',Mandatory=$True)]
        [Parameter(ParameterSetName='Existing',Mandatory=$true)]
        [string]
        $VNETSubnetAddressPrefix,
        [Parameter(Mandatory=$false)]
        [switch]
        $Deployment
    )
    begin{
        $ErrorActionPreference = "Stop"
        $WarningActionPreference = "SilentlyContinue"
    }
    process{

        <#
            $BEG_VNET          | Información de la virtual Network a Emplear.
            $BEG_VNET_SUBNET   | Información de la subnet a emplear para el WAF.
        #>

        $TMP_VAL = Get-RevoResourceGroup -ResourceGroupName $ResourceGroupName -Region $Region -Silent
        if(!$TMP_VAL -and !$Deployment){
            Write-RevoMessageConsole "El grupo de Recursos indicado no existe por favor verifique la información proporcionada en el valor -ResourceGroupName " -Type Error
            break;
        }
        elseif(!$TMP_VAL -and $Deployment){
            Get-RevoResourceGroup -ResourceGroupName $ResourceGroupName -Region $Region -Deployment -Silent
            Write-RevoMessageConsole 'El grupo de Recursos indicado no existía pero debido a que se especificó el parámetro -Deployment, se procedió con su creación' -Type Confirmation
        }

        if($UseExistingVNET){
            $BEG_VNET = Get-AzureRmVirtualNetwork -Name $VNETName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
            if($BEG_VNET){
                
                $BEG_VNET_SUBNET = $BEG_VNET | Test-RevoSubnetInformation -SubnetInformation $VNETSubnetAddressPrefix

                # Información al Usuario de si se encontró la información o no se encontró
                if($BEG_VNET_SUBNET){
                    Write-RevoMessageConsole "Se logró encontrar la Subnet '$VNETSubnetAddressPrefix' en la VNET" -Type Confirmation
                    return $BEG_VNET_SUBNET
                }
                else{
                    Write-RevoMessageConsole "No se logró encontrar la Subnet indicada, se intentará crear la misma" -Type Warning
                    
                    try{
                        Add-AzureRmVirtualNetworkSubnetConfig -Name ("WAFNetwork-"+ (Get-Random -Minimum 100 -Maximum 999)) -VirtualNetwork $BEG_VNET -AddressPrefix $VNETSubnetAddressPrefix | Out-Null;
                        $BEG_VNET = Set-AzureRmVirtualNetwork -VirtualNetwork $BEG_VNET -ErrorAction Stop;
                        $BEG_VNET_SUBNET = $BEG_VNET | Test-RevoSubnetInformation -SubnetInformation $VNETSubnetAddressPrefix
                        Write-RevoMessageConsole 'Se creó la red de forma exitosa' -Type Confirmation
                        return $BEG_VNET_SUBNET
                    }
                    catch{
                        Write-RevoMessageConsole 'No se pudo crear la red especificada, intente nuevamente con una red diferente' -Type Error;
                        return $null
                        break;
                    }
                }
            }
            else{
                Write-RevoMessageConsole 'No se ha podido encontrar la Red Virtual especificada, por favor emplee el parámetro UseNewVNET o valide el nombre ingresado e intente nuevamente' -Type Error 
                return $null
            }
        }
        elseif($UseNewVNET){
            $BEG_VNET = Get-AzureRmVirtualNetwork -Name $VNETName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
            if($BEG_VNET){
                Write-RevoMessageConsole 'Se especificó el parámetro para una nueva red pero ya exista una con el mismo nombre, por favor revise la información proporcionada' -Type Warning;
                break
            }
            else{
                $BEG_VNET_SUBNET = New-AzureRmVirtualNetworkSubnetConfig -Name ("WAFNetwork-"+ (Get-Random -Minimum 100 -Maximum 999)) -AddressPrefix $VNETSubnetAddressPrefix -WarningAction SilentlyContinue ;
                try{
                    $BEG_VNET = New-AzureRmVirtualNetwork -Name $VNETName -ResourceGroupName $ResourceGroupName -Location $Region -AddressPrefix $VNETGlobalAddressPrefix -Subnet $BEG_VNET_SUBNET -ErrorAction Stop -WarningAction SilentlyContinue;
                }
                catch{
                    Write-RevoMessageConsole 'No se pudo crear la red indicada, por favor validar los valores de -VNETGlobalAddressPrefix y -VNETSubnetAddressPrefix ' -Type Error
                    return $null
                }

                $BEG_VNET_SUBNET = $null
                $BEG_VNET_SUBNET = $BEG_VNET | Test-RevoSubnetInformation -SubnetInformation $VNETSubnetAddressPrefix

                if($BEG_VNET_SUBNET){
                    Write-RevoMessageConsole "La Subnet '$VNETSubnetAddressPrefix' fue creada de forma exitosa" -Type Confirmation
                    return $BEG_VNET_SUBNET
                }
            }
        }

    }
}



        <# Create a public IP address
        $publicip = New-AzureRmPublicIpAddress -ResourceGroupName $rg.ResourceGroupName -name publicIP01 -location EastUs -AllocationMethod Dynamic

        # Create a new IP configuration
        $gipconfig = New-AzureRmApplicationGatewayIPConfiguration -Name gatewayIP01 -Subnet $subnet

        # Create a backend pool with the hostname of the web app
        $pool = New-AzureRmApplicationGatewayBackendAddressPool -Name appGatewayBackendPool -BackendFqdns $webapp.HostNames

        # Define the status codes to match for the probe
        $match = New-AzureRmApplicationGatewayProbeHealthResponseMatch -StatusCode 200-399

        # Create a probe with the PickHostNameFromBackendHttpSettings switch for web apps
        $probeconfig = New-AzureRmApplicationGatewayProbeConfig -name webappprobe -Protocol Http -Path / -Interval 30 -Timeout 120 -UnhealthyThreshold 3 -PickHostNameFromBackendHttpSettings -Match $match

        # Define the backend http settings
        $poolSetting = New-AzureRmApplicationGatewayBackendHttpSettings -Name appGatewayBackendHttpSettings -Port 80 -Protocol Http -CookieBasedAffinity Disabled -RequestTimeout 120 -PickHostNameFromBackendAddress -Probe $probeconfig

        # Create a new front-end port
        $fp = New-AzureRmApplicationGatewayFrontendPort -Name frontendport01  -Port 80

        # Create a new front end IP configuration
        $fipconfig = New-AzureRmApplicationGatewayFrontendIPConfig -Name fipconfig01 -PublicIPAddress $publicip

        # Create a new listener using the front-end ip configuration and port created earlier
        $listener = New-AzureRmApplicationGatewayHttpListener -Name listener01 -Protocol Http -FrontendIPConfiguration $fipconfig -FrontendPort $fp

        # Create a new rule
        $rule = New-AzureRmApplicationGatewayRequestRoutingRule -Name rule01 -RuleType Basic -BackendHttpSettings $poolSetting -HttpListener $listener -BackendAddressPool $pool 

        # Define the application gateway SKU to use
        $sku = New-AzureRmApplicationGatewaySku -Name Standard_Small -Tier Standard -Capacity 2

        # Create the application gateway
        $appgw = New-AzureRmApplicationGateway -Name ContosoAppGateway -ResourceGroupName $rg.ResourceGroupName -Location EastUs -BackendAddressPools $pool -BackendHttpSettingsCollection $poolSetting -Probes $probeconfig -FrontendIpConfigurations $fipconfig  -GatewayIpConfigurations $gipconfig -FrontendPorts $fp -HttpListeners $listener -RequestRoutingRules $rule -Sku $sku
        #>
