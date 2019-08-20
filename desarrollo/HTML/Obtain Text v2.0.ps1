#-----------------------------------------------------------------------------------------------------------------
# Detalle:
# - Las variables contienen detalle la cual seguirá luego de un " _ " de la misma, ejemplo:
#                      MTX_FILE_PATH
# Dicha variable hace referencia al archivo de la matrix, por "MTX_FILE" y "_PATH" ya que es enfocado a la ruta del mismo
#-----------------------------------------------------------------------------------------------------------------
# Definición de Variables:
# $SRC_XXX = Referencia a una variable para la búsqueda de un valor. 
#       _XXX = Tipo de Valor a Buscar.
# $OBJ_OUT = Variable creada para el almacenamiento de los datos. Orientada a la exportación de información.
# $MTX_FOLDER = Variable matrix de ruta principal de los elementos HTML a analizar.
# $MTX_FILE = Variable representativa por el elemento de la matrix principal.
#-----------------------------------------------------------------------------------------------------------------
$OBJ_OUT = @()
$MTX_FOLDER = Get-ChildItem -Path C:\Users\Rodolfo\Downloads\www.pyme.pe\* –Include *.htm
foreach($MTX_FILE in $MTX_FOLDER){
#---------------------------------------------- Variables Generales ----------------------------------------------
        $MTX_FILE_PATH = $MTX_FILE.FullName
        $OBJ_OUT_PRO = New-Object PSObject
#---------------------------------------- Búsqueda del Nombre de Empresa -----------------------------------------
$SRC_NAME = Select-String -Path $MTX_FILE_PATH -pattern "<!-- NOMBRE COMERCIAL -->"
if ($SRC_NAME){ 
#------------------------------------------ Busqueda de Variable -------------------------------------------------
    $SRC_NAME_STRG = $SRC_NAME.ToString()
    $SRC_NAME_STRG_POS_1 = $SRC_NAME_STRG.IndexOf(':')
    $SRC_NAME_STRG_PH_1 = $SRC_NAME_STRG.Substring($SRC_NAME_STRG_POS_1+1)
    $SRC_NAME_STRG_POS_2 = $SRC_NAME_STRG_PH_1.IndexOf(':')
    $SRC_NAME_STRG_PH_2 = $SRC_NAME_STRG_PH_1.Substring($SRC_NAME_STRG_POS_2+1)
    $SRC_NAME_STRG_POS_3 = $SRC_NAME_STRG_PH_2.IndexOf(':')
    $SRC_NAME_STRG_PH_3 = $SRC_NAME_STRG_PH_2.Substring(0, $SRC_NAME_STRG_POS_3)
    [int]$SRC_NAME_INTE = [convert]::ToInt32($SRC_NAME_STRG_PH_3, 10)
    $SRC_NAME_LINE = $SRC_NAME_INTE + 1            
#----------------------------------------- Procesamiento de Búsqueda  --------------------------------------------
    $EMP_NAME_BRUTE = Get-Content $MTX_FILE_PATH | Select -Index $SRC_NAME_LINE
    $EMP_NAME_BRUTE_POS_1 = $EMP_NAME_BRUTE.IndexOf(">")
    $EMP_NAME_BRUTE_PH_1 = $EMP_NAME_BRUTE.Substring($EMP_NAME_BRUTE_POS_1+1)
    $EMP_NAME_BRUTE_POS_2 = $EMP_NAME_BRUTE_PH_1.IndexOf("<")
    $EMP_NAME = $EMP_NAME_BRUTE_PH_1.Substring(0, $EMP_NAME_BRUTE_POS_2)
#------------------------------------------ Adición de Propiedad OUT ---------------------------------------------
    Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name "Nombre Comercial" -Value $EMP_NAME
#-----------------------------------------------------------------------------------------------------------------
}
#------------------------------------- Búsqueda delos móviles de la Empresa --------------------------------------
$SRC_MOV = Select-String -Path $MTX_FILE_PATH -pattern "<!-- TELEFONOS CELULARES -->"
if ($SRC_MOV){ 
#------------------------------------------ Busqueda de Variable -------------------------------------------------
    $SRC_MOV_STRG = $SRC_MOV.ToString()
    $SRC_MOV_STRG_POS_1 = $SRC_MOV_STRG.IndexOf(':')
    $SRC_MOV_STRG_PH_1 = $SRC_MOV_STRG.Substring($SRC_MOV_STRG_POS_1+1)
    $SRC_MOV_STRG_POS_2 = $SRC_MOV_STRG_PH_1.IndexOf(':')
    $SRC_MOV_STRG_PH_2 = $SRC_MOV_STRG_PH_1.Substring($SRC_MOV_STRG_POS_2+1)
    $SRC_MOV_STRG_POS_3 = $SRC_MOV_STRG_PH_2.IndexOf(':')
    $SRC_MOV_STRG_PH_3 = $SRC_MOV_STRG_PH_2.Substring(0, $SRC_MOV_STRG_POS_3)
    [int]$SRC_MOV_INTE = [convert]::ToInt32($SRC_MOV_STRG_PH_3, 10)
    $SRC_MOV_LINE = $SRC_MOV_INTE            
#----------------------------------------- Procesamiento de Búsqueda  --------------------------------------------
    $EMP_MOV_BRUTE = Get-Content $MTX_FILE_PATH | Select -Index $SRC_MOV_LINE
    $EMP_MOV_BRUTE_POS_1 = $EMP_MOV_BRUTE.IndexOf('/i>')
    $EMP_MOV_BRUTE_PH_1 = $EMP_MOV_BRUTE.Substring($EMP_MOV_BRUTE_POS_1+1)
    $EMP_MOV_BRUTE_POS_2 = $EMP_MOV_BRUTE_PH_1.IndexOf('>')
    $EMP_MOV_BRUTE_PH_2 = $EMP_MOV_BRUTE_PH_1.Substring($EMP_MOV_BRUTE_POS_2+1)
    $EMP_MOV_BRUTE_POS_3 = $EMP_MOV_BRUTE_PH_2.IndexOf('<')
    $EMP_MOV = $EMP_MOV_BRUTE_PH_2.Substring(0, $EMP_MOV_BRUTE_POS_3)
#------------------------------------------ Adición de Propiedad OUT ---------------------------------------------
    Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name "Móviles" -Value $EMP_MOV
#-----------------------------------------------------------------------------------------------------------------
}
#----------------------------------- Búsqueda de los teléfonos de la Empresa -------------------------------------
$SRC_FIJO = Select-String -Path $MTX_FILE_PATH -pattern "<!-- TELEFONOS FIJOS -->"
if ($SRC_FIJO){ 
#------------------------------------------ Busqueda de Variable -------------------------------------------------
    $SRC_FIJO_STRG = $SRC_FIJO.ToString()
    $SRC_FIJO_STRG_POS_1 = $SRC_FIJO_STRG.IndexOf(':')
    $SRC_FIJO_STRG_PH_1 = $SRC_FIJO_STRG.Substring($SRC_FIJO_STRG_POS_1+1)
    $SRC_FIJO_STRG_POS_2 = $SRC_FIJO_STRG_PH_1.IndexOf(':')
    $SRC_FIJO_STRG_PH_2 = $SRC_FIJO_STRG_PH_1.Substring($SRC_FIJO_STRG_POS_2+1)
    $SRC_FIJO_STRG_POS_3 = $SRC_FIJO_STRG_PH_2.IndexOf(':')
    $SRC_FIJO_STRG_PH_3 = $SRC_FIJO_STRG_PH_2.Substring(0, $SRC_FIJO_STRG_POS_3)
    [int]$SRC_FIJO_INTE = [convert]::ToInt32($SRC_FIJO_STRG_PH_3, 10)
    $SRC_FIJO_LINE = $SRC_FIJO_INTE            
#----------------------------------------- Procesamiento de Búsqueda  --------------------------------------------
    $EMP_FIJO_BRUTE = Get-Content $MTX_FILE_PATH | Select -Index $SRC_FIJO_LINE
    $EMP_FIJO_BRUTE_POS_1 = $EMP_FIJO_BRUTE.IndexOf('/i>')
    $EMP_FIJO_BRUTE_PH_1 = $EMP_FIJO_BRUTE.Substring($EMP_FIJO_BRUTE_POS_1+1)
    $EMP_FIJO_BRUTE_POS_2 = $EMP_FIJO_BRUTE_PH_1.IndexOf('>')
    $EMP_FIJO = $EMP_FIJO_BRUTE_PH_1.Substring($EMP_FIJO_BRUTE_POS_2+1)
#------------------------------------------ Adición de Propiedad OUT ---------------------------------------------
    Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name "Fijos" -Value $EMP_FIJO
#-----------------------------------------------------------------------------------------------------------------
}
#--------------------------------------- Búsqueda del Correo de la Empresa ---------------------------------------
$SRC_MAIL = Select-String -Path $MTX_FILE_PATH -pattern "<!-- EMAIL -->"
if ($SRC_MAIL){ 
#------------------------------------------ Busqueda de Variable -------------------------------------------------
    $SRC_MAIL_STRG = $SRC_MAIL.ToString()
    $SRC_MAIL_STRG_POS_1 = $SRC_MAIL_STRG.IndexOf(':')
    $SRC_MAIL_STRG_PH_1 = $SRC_MAIL_STRG.Substring($SRC_MAIL_STRG_POS_1+1)
    $SRC_MAIL_STRG_POS_2 = $SRC_MAIL_STRG_PH_1.IndexOf(':')
    $SRC_MAIL_STRG_PH_2 = $SRC_MAIL_STRG_PH_1.Substring($SRC_MAIL_STRG_POS_2+1)
    $SRC_MAIL_STRG_POS_3 = $SRC_MAIL_STRG_PH_2.IndexOf(':')
    $SRC_MAIL_STRG_PH_3 = $SRC_MAIL_STRG_PH_2.Substring(0, $SRC_MAIL_STRG_POS_3)
    [int]$SRC_MAIL_INTE = [convert]::ToInt32($SRC_MAIL_STRG_PH_3, 10)
    $SRC_MAIL_LINE = $SRC_MAIL_INTE            
#----------------------------------------- Procesamiento de Búsqueda  --------------------------------------------
    $EMP_MAIL_BRUTE = Get-Content $MTX_FILE_PATH | Select -Index $SRC_MAIL_LINE
    $EMP_MAIL_BRUTE_POS_1 = $EMP_MAIL_BRUTE.IndexOf('href')
    $EMP_MAIL_BRUTE_PH_1 = $EMP_MAIL_BRUTE.Substring($EMP_MAIL_BRUTE_POS_1+1)
    $EMP_MAIL_BRUTE_POS_2 = $EMP_MAIL_BRUTE_PH_1.IndexOf('"')
    $EMP_MAIL_BRUTE_PH_2 = $EMP_MAIL_BRUTE_PH_1.Substring($EMP_MAIL_BRUTE_POS_2+1)
    $EMP_MAIL_BRUTE_POS_3 = $EMP_MAIL_BRUTE_PH_2.IndexOf('"')
    $EMP_MAIL = $EMP_MAIL_BRUTE_PH_2.Substring(0, $EMP_MAIL_BRUTE_POS_3)
#------------------------------------------ Adición de Propiedad OUT ---------------------------------------------
    Add-Member -InputObject $OBJ_OUT_PRO -MemberType NoteProperty -Name "Correo Electrónico" -Value $EMP_MAIL
#-----------------------------------------------------------------------------------------------------------------
}


$OBJ_OUT += $OBJ_OUT_PRO

}
$OBJ_OUT | Out-GridView -Title "Información Personal"
$OBJ_OUT | Export-Csv C:\Users\Rodolfo\Desktop\Exportación.csv 
