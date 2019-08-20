function Get-ADSitesandSubnets{
    $CFG_DN = (Get-ADRootDSE).ConfigurationNamingContext
    $SIT_DN = ("CN=Sites," + $CFG_DN)
    $SIT_OB_DB = Get-ADObject -SearchBase $SIT_DN -filter {objectClass -eq "site"} -properties "siteObjectBL", name
    foreach ($SIT_OB_IN in $SIT_OB_DB) {
        $SUB_AR = New-Object -Type string[] -ArgumentList $SIT_OB_IN.siteObjectBL.Count
        $COU_01 = 0
        $RTN_AL = New-Object PSCustomObject | Select SiteName, Subnets, DistinguishedName
        foreach ($SUB_DN in $SIT_OB_IN.siteObjectBL) {
            $SUB_NA = $SUB_DN.SubString(3, $SUB_DN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
            $SUB_AR[$COU_01] = $SUB_NA
            $COU_01++
        }
        $RTN_AL.SiteName = $SIT_OB_IN.Name
        $RTN_AL.Subnets = $SUB_AR
        $RTN_AL.DistinguishedName = $SIT_OB_IN.DistinguishedName
        Return $RTN_AL | fl *
    }
}

Get-ADSitesandSubnets > $Env:UserProfile\Desktop\Get-ADSitesandSubnets.txt

function Get-ADSubnetDetails{
    $siteName = 'Default-First-Site-Name'
    $CFG_DN = (Get-ADRootDSE).ConfigurationNamingContext
    $SIT_DN = ("CN=Sites," + $CFG_DN)
    $siteDN = "CN=" + $siteName + "," + $SIT_DN
    $SIT_OB_IN = Get-ADObject -Identity $siteDN -properties "siteObjectBL", "description", "location" 
    $RTN_AL = New-Object PSCustomObject | Select Description, DistinguishedName, Name,ObjectClass,ObjectGUID,siteObject
    foreach ($SUB_DN_IN in $SIT_OB_IN.siteObjectBL) {
        $SUB_IN_IN = Get-ADObject -Identity $SUB_DN_IN -properties "siteObject", "description", "location" 
        $RTN_AL.Description = $SUB_IN_IN.Description
        $RTN_AL.DistinguishedName = $SUB_IN_IN.DistinguishedName
        $RTN_AL.Name = $SUB_IN_IN.Name
        $RTN_AL.ObjectClass = $SUB_IN_IN.ObjectClass
        $RTN_AL.ObjectGUID = $SUB_IN_IN.ObjectGUID
        $RTN_AL.siteObject = $SUB_IN_IN.siteObject
        Return $RTN_AL | fl *
    }
}
Get-ADSubnetDetails > $Env:UserProfile\Desktop\Get-ADSubnetDetails.txt

function Get-ADInformationDetails{
    $RTN_AL = New-Object PSCustomObject | Select configurationNamingContext, defaultNamingContext,dnsHostName,domainControllerFunctionality,domainFunctionality,forestFunctionality,isGlobalCatalogReady
    $SUB_IN_IN = Get-ADRootDSE | Select *
    $RTN_AL.configurationNamingContext = $SUB_IN_IN.configurationNamingContext
    $RTN_AL.defaultNamingContext = $SUB_IN_IN.defaultNamingContext
    $RTN_AL.dnsHostName = $SUB_IN_IN.dnsHostName
    $RTN_AL.domainControllerFunctionality = $SUB_IN_IN.domainControllerFunctionality
    $RTN_AL.domainFunctionality = $SUB_IN_IN.domainFunctionality
    $RTN_AL.forestFunctionality = $SUB_IN_IN.forestFunctionality
    $RTN_AL.isGlobalCatalogReady = $SUB_IN_IN.isGlobalCatalogReady
    Return $RTN_AL | fl *


}

Get-ADInformationDetails > $Env:UserProfile\Desktop\Get-ADInformationDetails.txt