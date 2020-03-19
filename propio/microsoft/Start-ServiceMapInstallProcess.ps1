$DB_COM = Import-Csv C:\Scripts\QROMA-Computers.txt
$TL_PAT = "C:\Scripts\"
$IN_DOM = ".cppq.com.pe"
$RM_PAT = "\admin$\CloudMigration\"
$AG_MOM = "MMASetup-AMD64.exe"
$AG_DEP = "InstallDependencyAgent-Windows.exe"
$LG_NAM = (Get-Date -Format "yyyyMMdd") + "-Deploymentlog.log"
Remove-Item ($TL_PAT + $LG_NAM) -Force  -ErrorAction SilentlyContinue
"Computador,Sesion,Copia,EstadoRDP" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
$IN_CON = 1
foreach($INT in $DB_COM){
    $CP_NAM = $INT.ComputerName
    $CP_DOM = ($CP_NAM + $IN_DOM)
    Write-Progress -Activity "Instalando Agentes" -Status "Instalando agente en equipo $CP_NAM" -Id 100 -PercentComplete ($IN_CON / $DB_COMP.Count*100)
    Write-Progress -Activity ($CP_NAM.ToUpper()) -Status "Validaciones iniciales" -ParentId 100 -PercentComplete 25
    $CP_SES = New-PSSession -ComputerName $CP_DOM -ErrorAction SilentlyContinue
    $CP_RDP = Test-NetConnection -ComputerName ($CP_NAM + $IN_DOM) -Port 3389 -InformationLevel Quiet -WarningAction SilentlyContinue
    $CP_PAT = Test-Path ("\\" + $CP_NAM + $IN_DOM + "\C$\CloudMigration") -ErrorAction SilentlyContinue
    Write-Progress -Activity ($CP_NAM.ToUpper()) -Status "Creación de directorio" -ParentId 100 -PercentComplete 50
    if($CP_SES){
        $JOB_CRE = Invoke-Command -ComputerName $CP_DOM -AsJob -ScriptBlock { 
            New-Item -Path C:\ -Name CloudMigration -type directory -Force
        }
        $SESSION = $true
    }
    else{
        $SESSION = $false
    }
    Start-Sleep 10
    Write-Progress -Activity ($CP_NAM.ToUpper()) -Status "Copia de archivos" -ParentId 100 -PercentComplete 75
    if($SESSION){
        try{
            Copy-Item -Path "C:\Agent\*" -Destination ("\\" + $CP_NAM + $IN_DOM + "\C$\CloudMigration") -Force -Verbose
            $COPIED = $true
        }
        catch{
            $COPIED = $false
        }
    }
    else{
        try{
            Remove-Item ("\\" + $CP_NAM + $IN_DOM + "\C$\CloudMigration") -Force -Recurse
            xcopy.exe C:\Agent\* ("\\" + $CP_NAM + $IN_DOM + "\C$\CloudMigration\*") /E /Y
            $COPIED = $true
        }
        catch{
            $COPIED = $false
        }
    }

    Write-Progress -Activity ($CP_NAM.ToUpper()) -Status "Instalación de agentes" -ParentId 100 -PercentComplete 95
    if($SESSION -and $COPIED){     
        Invoke-Command -ComputerName $CP_DOM -ScriptBlock {
            Set-Location "C:\CloudMigration\"
            Start-Process "InstallDependencyAgent-Windows.exe" -ArgumentList '/C:"InstallDependencyAgent-Windows.exe /S /AcceptEndUserLicenseAgreement:1"'
        } -InDisconnectedSession
        Start-Sleep 20
        Invoke-Command -ComputerName $CP_DOM -ScriptBlock {
            Set-Location "C:\CloudMigration\"
            Start-Process "MMASetup-AMD64.exe" -ArgumentList '/C:"setup.exe /qn ADD_OPINSIGHTS_WORKSPACE=1 OPINSIGHTS_WORKSPACE_ID=8121b54c-f42b-48de-a029-f0cae99f0b01 OPINSIGHTS_WORKSPACE_KEY=6wRLFY6bUO5H5rXRrAq3TgPdBp4m+5XkAXVF+jNYPYYa2ccBhiuOdKOmwOvbbggozL20K2A3S3u1GzujuxJpOA== AcceptEndUserLicenseAgreement=1"'
        } -InDisconnectedSession
        $INSTALLED = $true
    }
    else{
        $INSTALLED = $false
    }
    Start-Sleep 10
    Write-Progress -Activity ($CP_NAM.ToUpper()) -Status "Cierre y finalización" -ParentId 100 -PercentComplete 100
    if(!$INSTALLED){
        if($COPIED){
            if($CP_RDP){
                $CP_DOM + ",Sesion no posible,Archivos copiados,RDP disponible" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
            }
            else{
                $CP_DOM + ",Sesion no posible,Archivos copiados,RDP fallido" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
            }
        }
        else{
            if($CP_RDP){
                $CP_DOM + ",Sesion no posible,Archivos no copiados,RDP disponible" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
            }
            else{
                $CP_DOM + ",Sesion no posible,Archivos no copiados,RDP fallido" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
            }
        }
    }
    else{
        if($COPIED){
            $CP_DOM + ",Agentes instalados,Archivos copiados,RDP no necesario" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
        }
        else{
            $CP_DOM + ",Sesion posible,Archivos no copiados,RDP no necesario" | Out-File -FilePath ($TL_PAT + $LG_NAM) -Append -Encoding utf8
        }
    }
    
    $IN_CON++
}
Get-PSSession | % { Remove-PSSession -Session $_}

<#
https://download.microsoft.com/download/3/d/b/3db49584-aa1e-403d-99b3-1083fcf931b5/MMASetup-AMD64.exe
https://aka.ms/dependencyagentwindows
#>