Write-host 'Escriba la letra de la unidad del VHD atachado'
$LETT = Read-Host
$GBL_SRC = $LETT + ':\'

$DB_USRS = Get-ChildItem -Path $GBL_SRC | Where {($_.Attributes -eq 'Directory') -and ($_.FullName -like '*@*')} # Obtener todos las carpetas de cada usuario
$PROG = 1
foreach($FOR_01 in $DB_USRS){
    
    $FOR_01_MAIL_FOLDER = $FOR_01.FullName + '\mail'
    $DB_MAIL = Get-ChildItem -Path $FOR_01_MAIL_FOLDER | Where {($_.Attributes -eq 'Directory')} # obtener la carpeta mail
    foreach($FOR_02 in $DB_MAIL){
    $PRO_Nam = $FOR_02.name
    
    Write-Progress -Activity “Ordenando Carpetas” -status “Ejecutando sobre carpeta: $PRO_Nam” -percentComplete ($PROG / $DB_MAIL.count*100)


        $DB_DATE = Get-ChildItem -Path $FOR_02.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -match '[1-2][0-9][0-9][0-9][0-9][0-9]')} 
        $DB_NO_DATE = Get-ChildItem -Path $FOR_02.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -notmatch '[1-2][0-9][0-9][0-9][0-9][0-9]')} 
       
       
       
       
        if($DB_DATE -ne $null){
            foreach($FOR_03 in $DB_DATE){
                $TMP_01 = $FOR_03.FullName
                Get-ChildItem -Path "$TMP_01\*.eml" | Move-Item -Destination $FOR_02.FullName -Force -Confirm:$false
                Remove-Item -Path $TMP_01 -Force -Confirm:$false
            }
        }
        

        if($DB_NO_DATE -ne $null){
            foreach($FOR_04 in $DB_NO_DATE){
                $TMP_02 = $FOR_04.FullName

                $DB_DATE_2 = Get-ChildItem -Path $FOR_04.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -match '[1-2][0-9][0-9][0-9][0-9][0-9]')} 
                $DB_NO_DATE_2 = Get-ChildItem -Path $FOR_04.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -notmatch '[1-2][0-9][0-9][0-9][0-9][0-9]')} 

                if($DB_DATE_2 -ne $null){
                    foreach($FOR_04_01 in $DB_DATE_2){
                        $TMP_03 = $FOR_04_01.FullName
                        Get-ChildItem -Path "$TMP_03\*.eml" | Move-Item -Destination $FOR_04.FullName -Force -Confirm:$false
                        Remove-Item -Path $TMP_03 -Force -Confirm:$false
                    }
                }

                if($DB_NO_DATE_2 -ne $null){
                    foreach($FOR_05 in $DB_NO_DATE_2){
                        $TMP_04 = $FOR_05.FullName

                        $DB_DATE_3 = Get-ChildItem -Path $FOR_05.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -match '[1-2][0-9][0-9][0-9][0-9][0-9]')} 
                        $DB_NO_DATE_3 = Get-ChildItem -Path $FOR_05.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -notmatch '[1-2][0-9][0-9][0-9][0-9][0-9]')} 

                        if($DB_DATE_3 -ne $null){
                            foreach($FOR_05_01 in $DB_DATE_3){
                                $TMP_05 = $FOR_05_01.FullName
                                Get-ChildItem -Path "$TMP_05\*.eml" | Move-Item -Destination $FOR_05.FullName -Force -Confirm:$false
                                Remove-Item -Path $TMP_05 -Force -Confirm:$false
                            }
                        }

                        if($DB_NO_DATE_3 -ne $null){
                            foreach($FOR_06 in $DB_NO_DATE_3){
                                $TMP_05 = $FOR_06.FullName

                                $DB_DATE_4 = Get-ChildItem -Path $FOR_06.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -match '[1-2][0-9][0-9][0-9][0-9][0-9]')} 
                                $DB_NO_DATE_4 = Get-ChildItem -Path $FOR_06.FullName | Where {($_.Attributes -eq 'Directory') -and ($_.Name -notmatch '[1-2][0-9][0-9][0-9][0-9][0-9]')} 

                                if($DB_DATE_4 -ne $null){
                                    foreach($FOR_06_01 in $DB_DATE_4){
                                        $TMP_06 = $FOR_06_01.FullName
                                        Get-ChildItem -Path "$TMP_06\*.eml" | Move-Item -Destination $FOR_06.FullName -Force -Confirm:$false
                                        Remove-Item -Path $TMP_06 -Force -Confirm:$false
                                    }
                                }

                                if($DB_NO_DATE_4 -ne $null){


                                }
                            }
                        }
                    }
                }
            }
        }

    
   $PROG++
   }
Write-Host '---------------------------------------------------------------------'
Write-Host '   Ordenamiento de carpetas de' $FOR_01.Name 'completada'
Write-Host '---------------------------------------------------------------------'
}
pause