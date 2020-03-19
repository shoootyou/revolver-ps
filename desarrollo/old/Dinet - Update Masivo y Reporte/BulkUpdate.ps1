Import-Module ActiveDirectory
function Get-FilePath{
    [cmdletbinding()]
    param(
            [string]$Title,
            [string]$Filter =  "CSV Files (*.csv)|*.csv",
            [string]$Path = $env:USERPROFILE
    )
    process{
        if($Title -eq $null){$Title = "Ubica la ruta del Archivo CSV "}
	    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	    $OBJ_IMP_PT = New-Object System.Windows.Forms.OpenFileDialog
	    $OBJ_IMP_PT.InitialDirectory = $Path
	    $OBJ_IMP_PT.Filter = $Filter
	    $OBJ_IMP_PT.Title = $Title
	    $Show = $OBJ_IMP_PT.ShowDialog()
	    If ($Show -eq "OK")
	    {
		    Return $OBJ_IMP_PT.FileName
	    }
        else{
            Write-Warning "No seleccionaste ningún archivo"
        }
    }
}

$GBL_PATH = Get-FilePath
$DB = Import-Csv -Path $GBL_PATH -Delimiter ','
Write-Host ''
Write-Host '--------------------------------------------------------------------' -ForegroundColor Green
foreach($USR_CSV in $DB){
    $USR_AD = Get-ADUser -Identity $USR_CSV.SamAccountName -Properties telephoneNumber,wWWHomePage,employeeID,msDS-cloudExtensionAttribute1,proxyAddresses
    Set-ADUser -Identity $USR_CSV.SamAccountName -Description $USR_CSV.Description -City $USR_CSV.City -State $USR_CSV.State -PostalCode $USR_CSV.PostalCode -Country $USR_CSV.Country -Office $USR_CSV.Office
 #----------------------------------------------------------------------------------------------------
    if($USR_CSV.Department){
        if($USR_AD.Department){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'Department'=$USR_CSV.Department}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'Department'=$USR_CSV.Department}
        }
    }
 #----------------------------------------------------------------------------------------------------
    if($USR_CSV.JobTitle){
        if($USR_AD.Title){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'Title'=$USR_CSV.JobTitle}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Title $USR_CSV.JobTitle
        }
    }
 #----------------------------------------------------------------------------------------------------
    if($USR_CSV.Company){
        if($USR_AD.Company){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'Company'=$USR_CSV.Company}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'Company'=$USR_CSV.Company}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.TelephoneNumber){
        if($USR_AD.telephoneNumber){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'telephoneNumber'=$USR_CSV.TelephoneNumber}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'telephoneNumber'=$USR_CSV.TelephoneNumber}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.WebPage){
        if($USR_AD.wWWHomePage){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'wWWHomePage'=$USR_CSV.WebPage}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'wWWHomePage'=$USR_CSV.WebPage}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.employeeID){
        if($USR_AD.employeeID){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'employeeID'=$USR_CSV.employeeID}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'employeeID'=$USR_CSV.employeeID}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dNI){
        if($USR_AD.dNI){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dNI'=$USR_CSV.dNI}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dNI'=$USR_CSV.dNI}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dntCeCo){
        if($USR_AD.dntCeCo){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dntCeCo'=$USR_CSV.dntCeCo}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dntCeCo'=$USR_CSV.dntCeCo}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dntDesCeco){
        if($USR_AD.dntDesCeco){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dntDesCeco'=$USR_CSV.dntDesCeco}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dntDesCeco'=$USR_CSV.dntDesCeco}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dntDesOrdenInterna){
        if($USR_AD.dntDesOrdenInterna){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dntDesOrdenInterna'=$USR_CSV.dntDesOrdenInterna}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dntDesOrdenInterna'=$USR_CSV.dntDesOrdenInterna}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dntFechaIngreso){
        if($USR_AD.dntFechaIngreso){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dntFechaIngreso'=$USR_CSV.dntFechaIngreso}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dntFechaIngreso'=$USR_CSV.dntFechaIngreso}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dntOrdenInterna){
        if($USR_AD.dntOrdenInterna){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dntOrdenInterna'=$USR_CSV.dntOrdenInterna}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dntOrdenInterna'=$USR_CSV.dntOrdenInterna}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.dntPuesto){
        if($USR_AD.dntPuesto){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'dntPuesto'=$USR_CSV.dntPuesto}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'dntPuesto'=$USR_CSV.dntPuesto}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.msDScloudExtensionAttribute1){
        if($USR_AD.'msDS-cloudExtensionAttribute1'){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'msDS-cloudExtensionAttribute1'=$USR_CSV.msDScloudExtensionAttribute1}
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'msDS-cloudExtensionAttribute1'=$USR_CSV.msDScloudExtensionAttribute1}
        }
    }
#----------------------------------------------------------------------------------------------------
    if($USR_CSV.PrincipalProxyAddress){
        if(($USR_AD.proxyAddresses) -and ($USR_AD.proxyAddresses -like 'SMTP:*')){
            Set-ADUser -Identity $USR_CSV.SamAccountName -Replace @{'proxyAddresses'=('SMTP:'+$USR_CSV.PrincipalProxyAddress)}
            if($USR_CSV.ProxyAddress1){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress1)}
            }
            if($USR_CSV.ProxyAddress2){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress2)}
            }
            if($USR_CSV.ProxyAddress3){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress3)}
            }
            if($USR_CSV.ProxyAddress4){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress4)}
            }
            if($USR_CSV.ProxyAddress5){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress5)}
            }
            if($USR_CSV.ProxyAddress6){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress6)}
            }
            if($USR_CSV.ProxyAddress7){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress7)}
            }
            if($USR_CSV.ProxyAddress8){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress8)}
            }
            if($USR_CSV.ProxyAddress9){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress9)}
            }
        }
        else{
            Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('SMTP:'+$USR_CSV.PrincipalProxyAddress)}
            if($USR_CSV.ProxyAddress1){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress1)}
            }
            if($USR_CSV.ProxyAddress2){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress2)}
            }
            if($USR_CSV.ProxyAddress3){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress3)}
            }
            if($USR_CSV.ProxyAddress4){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress4)}
            }
            if($USR_CSV.ProxyAddress5){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress5)}
            }
            if($USR_CSV.ProxyAddress6){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress6)}
            }
            if($USR_CSV.ProxyAddress7){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress7)}
            }
            if($USR_CSV.ProxyAddress8){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress8)}
            }
            if($USR_CSV.ProxyAddress9){
                Set-ADUser -Identity $USR_CSV.SamAccountName -Add @{'proxyAddresses'=('smtp:'+$USR_CSV.ProxyAddress9)}
            }
        }
    }
#------------------------------------
#            OUTPUT
#------------------------------------

    Write-Host 'Se completó la modificación para el usuario: ' -ForegroundColor Green -NoNewline
    Write-Host  $USR_CSV.SamAccountName -ForegroundColor Yellow 
    Write-Host '--------------------------------------------------------------------' -ForegroundColor Green
}
pause