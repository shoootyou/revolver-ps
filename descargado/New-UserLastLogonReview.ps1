<########################################################################### 
    The purpose of this PowerShell script is to collect the last logon  
    for user accounts on each DC in the domain, evaluate, and return the 
    most recent logon value. 
 
        Author:   Jeremy Reeves 
        Modified: 02/14/2018 
        Notes:    Must have RSAT Tools if running on a workstation 
 
###########################################################################> 
 
 
Import-Module ActiveDirectory 
 
function Get-ADUsersLastLogon($Username="*") { 
 
    $FilePath_Prefix = "C:\Slush\UserLastLogon-" 
 
    function Msg ($Txt="") { 
        Write-Host "$([DateTime]::Now)    $Txt" 
    } 
 
    #Cycle each DC and gather user account lastlogon attributes 
     
    $List = @() #Define Array 
    (Get-ADDomain).ReplicaDirectoryServers | Sort | % { 
 
        $DC = $_ 
        Msg "Reading $DC" 
        $List += Get-ADUser -Server $_ -Filter "samaccountname -like '$Username'" -Properties LastLogon | Select samaccountname,lastlogon,@{n='DC';e={$DC}} 
 
    } 
 
    Msg "Sorting for most recent lastlogon" 
     
    $LatestLogOn = @() #Define Array 
    $List | Group-Object -Property samaccountname | % { 
 
        $LatestLogOn += ($_.Group | Sort -prop lastlogon -Descending)[0] 
 
    } 
     
    $List.Clear() 
 
    if ($Username -eq "*") { #$Username variable was not set.    Running against all user accounts and exporting to a file. 
 
        $FileName = "$FilePath_Prefix$([DateTime]::Now.ToString("yyyyMMdd-HHmmss")).csv" 
         
        try { 
 
            $LatestLogOn | Select samaccountname, lastlogon, @{n='lastlogondatetime';e={[datetime]::FromFileTime($_.lastlogon)}}, DC | Export-CSV -Path $FileName -NoTypeInformation -Force 
            Msg "Exported results. $FileName" 
 
        } catch { 
 
            Msg "Export Failed. $FileName" 
 
        } 
 
    } else { #$Username variable was set, and may refer to a single user account. 
 
        if ($LatestLogOn) { $LatestLogOn | Select samaccountname, @{n='lastlogon';e={[datetime]::FromFileTime($_.lastlogon)}}, DC | FT } else { Msg "$Username not found." } 
 
    } 
 
    $LatestLogon.Clear() 
 
} 