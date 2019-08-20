$ContentFolder = Get-ChildItem -Path C:\Users\Rodolfo\Downloads\www.pyme.pe\* –Include *.htm
foreach($FileHTML in $ContentFolder){
#------------------------------------ Variables Generales ------------------------------------
        $FullPath = $FileHTML.FullName
#--------------------------------- Variables Nombre Empresa ----------------------------------
        $EMP_NAME_1 = Get-Content $FullPath | Select -Index 975
        $EMP_NAME_2 = Get-Content $FullPath | Select -Index 974
        $EMP_NAME_3 = Get-Content $FullPath | Select -Index 976
#------------------------------- Variables Correo Electronico --------------------------------
        $EMP_MX_1 = Get-Content $FullPath | Select -Index 1039
        $EMP_MX_2 = Get-Content $FullPath | Select -Index 1040
        $EMP_MX_3 = Get-Content $FullPath | Select -Index 1046
#---------------------------------------------------------------------------------------------
#---------------------------------  Nombre Empresa --------------------------------------------
        if($EMP_NAME_1 -like "*href*"){
            $EMP_NAME_POS_1_1 = $EMP_NAME_1.IndexOf(">")
            $EMP_NAME_PH_1_1 = $EMP_NAME_1.Substring($EMP_NAME_POS_1_1+1)
            $EMP_NAME_POS_2_1 = $EMP_NAME_PH_1_1.IndexOf("<")
            $EMP_NAME_PH_2_1 = $EMP_NAME_PH_1_1.Substring(0, $EMP_NAME_POS_2_1)
            Write-Host "------------------------"
            Write-Host $EMP_NAME_PH_2_1
        }
        elseif($EMP_NAME_2 -like "*href*"){
            $EMP_NAME_POS_1_2 = $EMP_NAME_2.IndexOf(">")
            $EMP_NAME_PH_1_2 = $EMP_NAME_2.Substring($EMP_NAME_POS_1_2+1)
            $EMP_NAME_POS_2_2 = $EMP_NAME_PH_1_2.IndexOf("<")
            $EMP_NAME_PH_2_2 = $EMP_NAME_PH_1_2.Substring(0, $EMP_NAME_POS_2_2)
            Write-Host "------------------------"
            Write-Host $EMP_NAME_PH_2_2
        }
        elseif($EMP_NAME_3 -like "*href*"){
            $EMP_NAME_POS_1_3 = $EMP_NAME_3.IndexOf(">")
            $EMP_NAME_PH_1_3 = $EMP_NAME_3.Substring($EMP_NAME_POS_1_3+1)
            $EMP_NAME_POS_2_3 = $EMP_NAME_PH_1_3.IndexOf("<")
            $EMP_NAME_PH_2_3 = $EMP_NAME_PH_1_3.Substring(0, $EMP_NAME_POS_2_3)
            Write-Host "------------------------"
            Write-Host $EMP_NAME_PH_2_3
        }


#------------------------------- Correo Electrónico  ------------------------------------------
        if($EMP_MX_1 -like "*email_plantilla_dashboard*"){
            $EMP_MX_POS_1_1 = $EMP_MX_1.IndexOf('href')
            $EMP_MX_PH_1_1 = $EMP_MX_1.Substring($EMP_MX_POS_1_1+1)
            $EMP_MX_POS_2_1 = $EMP_MX_PH_1_1.IndexOf('"')
            $EMP_MX_PH_2_1 = $EMP_MX_PH_1_1.Substring($EMP_MX_POS_2_1+1)
            $EMP_MX_POS_3_1 = $EMP_MX_PH_2_1.IndexOf('"')
            $EMP_MX_PH_3_1 = $EMP_MX_PH_2_1.Substring(0, $EMP_MX_POS_3_1)
            Write-Host "------------------------"
            Write-Host $EMP_MX_PH_3_1
        }
        elseif($EMP_MX_2 -like "*email_plantilla_dashboard*"){
            $EMP_MX_POS_1_2 = $EMP_MX_2.IndexOf('href')
            $EMP_MX_PH_1_2 = $EMP_MX_2.Substring($EMP_MX_POS_1_2+1)
            $EMP_MX_POS_2_2 = $EMP_MX_PH_1_2.IndexOf('"')
            $EMP_MX_PH_2_2 = $EMP_MX_PH_1_2.Substring($EMP_MX_POS_2_2+1)
            $EMP_MX_POS_3_2 = $EMP_MX_PH_2_2.IndexOf('"')
            $EMP_MX_PH_3_2 = $EMP_MX_PH_2_2.Substring(0, $EMP_MX_POS_3_2)
            Write-Host "------------------------"
            Write-Host $EMP_MX_PH_3_2
        }
        elseif($EMP_MX_3 -like "*email_plantilla_dashboard*"){
            $EMP_MX_POS_1_3 = $EMP_MX_3.IndexOf('href')
            $EMP_MX_PH_1_3 = $EMP_MX_3.Substring($EMP_MX_POS_1_3+1)
            $EMP_MX_POS_2_3 = $EMP_MX_PH_1_3.IndexOf('"')
            $EMP_MX_PH_2_3 = $EMP_MX_PH_1_3.Substring($EMP_MX_POS_2_3+1)
            $EMP_MX_POS_3_3 = $EMP_MX_PH_2_3.IndexOf('"')
            $EMP_MX_PH_3_3 = $EMP_MX_PH_2_3.Substring(0, $EMP_MX_POS_3_3)
            Write-Host "------------------------"
            Write-Host $EMP_MX_PH_3_3
        }
}