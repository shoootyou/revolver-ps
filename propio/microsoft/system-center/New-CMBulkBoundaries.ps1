$DB_PAT = "E:\Scripts\SPSA_BoundayExample.csv"
$DB_NET = Import-Csv $DB_PAT -Encoding UTF8
$CN_SED = "NoSede"
$CN_NUM = 1
$BN_DBS = @()
foreach($NET in $DB_NET){
    #Write-Host $NET.Sede
    #Write-Host $CN_SED
    #Write-Host $CN_NUM
    #Write-Host "-------------------------------------------------------------"
    if(($NET.Sede -ne $CN_SED)-or ($DB_NET.Length -eq $CN_NUM)){
        if($BN_DBS.Length -gt 0){
            $VL_BND = Get-CMBoundaryGroup -Name $NA_GRP
            if(!($VL_BND)){
                $BN_GRP = New-CMBoundaryGroup -Name $NA_GRP -Description $NA_DRP -ErrorAction SilentlyContinue
                Write-Host "Creando nuevo Boundary Group" $NA_GRP
                Remove-Variable -Name VL_BND
            }
            foreach($BN in $BN_DBS){
                Add-CMBoundaryToGroup -BoundaryGroupId $BN_GRP.GroupID -BoundaryId $BN.BoundaryID
                Write-Host "Agregando Boundary: " $BN.DisplayName "al grupo " $BN_GRP.Name
            }
            $BN_DBS = @()
        }
        $CN_SED = $NET.Sede
    }
    $BN_SGL = New-CMBoundary -Name $NET.Network -Type IPSubnet -Value $NET.Network 
    Write-Host "Creado Boundary: " $NET.Network
    $NA_GRP = $NET.BDGRP
    $NA_DRP = $NET.Sede
    $BN_DBS += $BN_SGL
    $CN_NUM++
}

Remove-Variable -Name BN
Remove-Variable -Name BN_DBS
Remove-Variable -Name BN_GRP
Remove-Variable -Name BN_SGL
Remove-Variable -Name CN_SED
Remove-Variable -Name DB_NET
Remove-Variable -Name DB_PAT
Remove-Variable -Name NA_DRP
Remove-Variable -Name NA_GRP
Remove-Variable -Name NET 
Remove-Variable -Name CN_NUM