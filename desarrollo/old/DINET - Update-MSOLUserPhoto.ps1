function Get-FolderPath{
    [cmdletbinding()]
    param(
            [bool]$NewFolder = $true,
            [string]$Description
    )
    process{
        if($Description -eq $null){ $Description =  "Encuentra tu carpeta" }
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        $OBJ_IMP_PT = New-Object System.Windows.Forms.FolderBrowserDialog
        $OBJ_IMP_PT.ShowNewFolderButton = $NewFolder
        $OBJ_IMP_PT.Description = $Description
        $Show = $OBJ_IMP_PT.ShowDialog()
        If ($Show -eq "OK")
        {
	        Return $OBJ_IMP_PT.SelectedPath
        }
        else{
            Write-Warning "No seleccionaste ninguna carpeta"
        }
    }
}

Import-Module ActiveDirectory
$USR_DB_LCL = GET-ADUser –filter * -properties thumbnailphoto | Where {$_.UserPrincipalName -notlike '*dinet.local'}
$GBL_FL = Get-FolderPath -Description 'Por favor ubicar la carpeta en donde se almacenarán las imágenes'

Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '                Se da inicio a la exportación de fotos del AD Local               ' -ForegroundColor Green
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Green

if($GBL_FL){
    $EXP_FL_LCL = $GBL_FL + '\ExportADLocal'
    $LOG_EXP_LCL = $GBL_FL + '\LogExportLocal.txt'
    
    if(Test-Path $GBL_FL){
        New-Item $EXP_FL_LCL -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    }
    $VAL_01 = Get-ChildItem $EXP_FL_LCL
    if(!($VAL_01)){
        Write-Host '                                        ##'
        Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write ExportedUsers | Out-File $LOG_EXP_LCL -Append
        $CON_LCL = 1
        foreach($USR_LCL in $USR_DB_LCL){
            Write-Progress -Activity “Exportando Fotos - AD Local” -status “Procesando el usuario $USR_LCL_UPN” -percentComplete ($CON_LCL / $USR_DB_LCL.count*100)
            If ($USR_LCL.thumbnailphoto){
                $USR_LCL_UPN = $USR_LCL.UserPrincipalName
                $FILE_LCL_EXP = $EXP_FL_LCL + '\' + $USR_LCL_UPN +'.jpg'
                [System.Io.File]::WriteAllBytes($FILE_LCL_EXP, $USR_LCL.Thumbnailphoto)
                Write $USR_LCL_UPN | Out-File $LOG_EXP_LCL -Append
              }
            $CON_LCL++
        }
        Write-Host '            Se culminó con la exportación de las fotos en el AD local             ' -ForegroundColor Cyan 
        Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Cyan
    }
    else{
	Write-Host '                                        ##'
        Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Cyan
        Write-Host '             Se encontró un proceso anterior, se continuará el proceso            ' -ForegroundColor Cyan
        Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Cyan
    }

}
else{
    Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Yellow
    Write-Host '            No se pudo encontrar la carpeta para dar inicio al proceso            ' -ForegroundColor Yellow
    Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Yellow
    Break
}

Write-Host '                                        ##'
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '              Iniciando servicios en Microsoft Office 365 y Azure AD              ' -ForegroundColor Green
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Green

###################################################################################################################################################################################################

$USER = 'admin365cloud@dinetcorp.onmicrosoft.com'
$PASS = ConvertTo-SecureString –String 'T2UhaT_&wA7e' –AsPlainText -Force

$CRED = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $USER, $PASS
$SESS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $CRED -Authentication Basic -AllowRedirection

Import-PSSession $SESS
Connect-MsolService -Credential $CRED

###################################################################################################################################################################################################

Write-Host '                                        ##'
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Green
Write-Host '             Se da inicio a la verificación y actualización de fotos              ' -ForegroundColor Green
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Green

$USR_DB_AZ = Get-MsolUser -All | Select UserPrincipalName | Sort UserPrincipalName

$TIME = Get-Date
$LOG_EXP_AZ = $GBL_FL + '\LogExportAzure-' + $TIME.Day + '-' + $TIME.Month + '-' + $TIME.Year + '-' + $TIME.Hour + '-' + $TIME.Minute + '-' + $TIME.Second + '.txt'

$UPLO_FL = $GBL_FL + '\UploadedPhotos'
New-Item $UPLO_FL -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

Write-Host '                                        ##'
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Cyan
Write 'UserPrincipalName,Status' | Out-File $LOG_EXP_AZ -Append

$CON_AZ = 1

foreach($USR_AZ in $USR_DB_AZ){
Write-Progress -Activity “Exportando Fotos - AD Azure” -status “Procesando el usuario $USR_AZ_UPN” -percentComplete ($CON_AZ / $USR_DB_AZ.count*100)
    $USR_AZ_UPN = $USR_AZ.UserPrincipalName
    $FIND_USR = $EXP_FL_LCL + '\' + $USR_AZ_UPN +'.jpg'
    
    
    if(Test-Path $FIND_USR){
        $AZ_USR_PH = Get-UserPhoto -Identity $USR_AZ.UserPrincipalName | Select PictureData
        $TMP_AZ_PH = $AZ_USR_PH.PictureData
        $ERRO_01 = $false

        $TMP_PH = $GBL_FL + '\TMP_PH.jpeg'
        [System.Io.File]::WriteAllBytes($TMP_PH, $TMP_AZ_PH)

        $COMP_AZ = ([Byte[]] $(Get-Content -Path $TMP_PH -Encoding Byte -ReadCount 0))
        $COMP_LCL = ([Byte[]] $(Get-Content -Path $FIND_USR -Encoding Byte -ReadCount 0))
        if(($COMP_AZ -ne $COMP_LCL)){
            Try{
                Set-UserPhoto -Identity $USR_AZ_UPN -PictureData $COMP_LCL -Confirm:$false
            }
            Catch{
                $OUT_00 = $USR_AZ_UPN + ',No se pudo cargar la foto'
                Write $OUT_00 | Out-File $LOG_EXP_AZ -Append
                $ERRO_01 = $true
            }
            $OUT_01 = $USR_AZ_UPN + ',Foto Actualizada'
            Write $OUT_01 | Out-File $LOG_EXP_AZ -Append
            if(!$ERRO_01){
                Move-Item $FIND_USR $UPLO_FL
            }
        }
       else{
            $OUT_02 = $USR_AZ_UPN + ',No requiere de Actualizacion'
            Write $OUT_02 | Out-File $LOG_EXP_AZ -Append
       }

    }
    else{
        $OUT_03 = $USR_AZ_UPN + ',No se encontró en AD Local'
        Write $OUT_03 | Out-File $LOG_EXP_AZ -Append

    }
$CON_AZ++
sleep 15
}
Write-Host '            Se culminó con la exportación de las fotos en el AD local             ' -ForegroundColor Cyan 
Write-Host '----------------------------------------------------------------------------------' -ForegroundColor Cyan
pause