##############################################################################################################################
## Autor: Rodolfo Castelo Méndez                                                                                            ##
## Script para la creación de la estructura de la oganización                                                               ##
## Nombre de la organización: G3CAJA.COM                                                                                    ##
## Nombres de los Computadores: (Nomenclatura del área)-(Número correlativo)                                                ##
## Versión 2.05                                                                                                             ##
##############################################################################################################################
################                         Importación del Módulo del Directorio Activo                         ################        
##############################################################################################################################
Import-Module ActiveDirectory
##############################################################################################################################
################                               Creación de la estructura de OUs                               ################        
##############################################################################################################################
New-ADOrganizationalUnit -Path "DC=G3CAJA,DC=COM" Internal
# Área de Servidores
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3CAJA,DC=COM" Servidores
New-ADOrganizationalUnit -Path "OU=Servidores,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Servidores,OU=Internal,DC=G3CAJA,DC=COM" Secundario
# Area de la DMZ
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3CAJA,DC=COM" DMZ
New-ADOrganizationalUnit -Path "OU=DMZ,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=DMZ,OU=Internal,DC=G3CAJA,DC=COM" Secundario
# Área de ingenieria
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3CAJA,DC=COM" Ingenieria
New-ADOrganizationalUnit -Path "OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM" Secundario
# Área de Contabilidad
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3CAJA,DC=COM" Contabilidad
New-ADOrganizationalUnit -Path "OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM" Secundario
# Área de Administración
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3CAJA,DC=COM" Administracion
New-ADOrganizationalUnit -Path "OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM" Secundario
# Área de Ventas
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3CAJA,DC=COM" Ventas
New-ADOrganizationalUnit -Path "OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM" Secundario
##############################################################################################################################
################                   Creación de los objetos de computador de la Organización                   ################        
##############################################################################################################################
# Area de Ingenieria
New-Adcomputer –name "ING-PC-01" –SamAccountName "ING-PC-01" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-02" –SamAccountName "ING-PC-02" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-03" –SamAccountName "ING-PC-03" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-04" –SamAccountName "ING-PC-04" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-05" –SamAccountName "ING-PC-05" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-06" –SamAccountName "ING-PC-06" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-07" –SamAccountName "ING-PC-07" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-08" –SamAccountName "ING-PC-08" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-09" –SamAccountName "ING-PC-09" -Path "OU=Secundario,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
New-Adcomputer –name "ING-PC-10" –SamAccountName "ING-PC-10" -Path "OU=Secundario,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
##############################################################################################################################
# Area de Contabilidad
$CONADM = 1
do{
$SAM = "CON-PC-" + $CONADM; $CONADM++
New-Adcomputer –name $SAM –SamAccountName $SAM -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM"
}
while($CONADM -le 15);
$CONSEC = 16
do{
$SAM = "CON-PC-" + $CONSEC; $CONSEC++
New-Adcomputer –name $SAM –SamAccountName $SAM -Path "OU=Secundario,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM"
}
while($CONSEC -le 20);
##############################################################################################################################
#  Area de Administración
$ADMADM = 1
do{
$SAM = "ADM-PC-" + $ADMADM; $ADMADM++
New-Adcomputer –name $SAM –SamAccountName $SAM -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM"
}
while($ADMADM -le 70);
$ADMSEC = 71
do{
$SAM = "ADM-PC-" + $ADMSEC; $ADMSEC++
New-Adcomputer –name $SAM –SamAccountName $SAM -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM"
}
while($ADMSEC -le 90);
##############################################################################################################################
# Area de Ventas
$VENADM = 1
do{
$SAM = "VEN-PC-" + $VENADM; $VENADM++
New-Adcomputer –name $SAM –SamAccountName $SAM -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM"
}
while($VENADM -le 36);
$VENSEC = 37
do{
$SAM = "VEN-PC-" + $VENSEC; $VENSEC++
New-Adcomputer –name $SAM –SamAccountName $SAM -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM"
}
while($VENSEC -le 50);
##############################################################################################################################
################                    Creación de la estructura de Grupos de la Organización                    ################        
##############################################################################################################################
# Área de ingenieria
New-ADGroup -Name ING-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3CAJA,DC=COM"
# Área de Contabilidad
New-ADGroup -Name CON-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM"
New-ADGroup -Name CON-Usuarios -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3CAJA,DC=COM"
# Área de Administración
New-ADGroup -Name ADM-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM"
New-ADGroup -Name ADM-Usuarios -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Administracion,OU=Internal,DC=G3CAJA,DC=COM"
# Área de Ventas
New-ADGroup -Name VEN-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM"
New-ADGroup -Name VEN-Usuarios -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Ventas,OU=Internal,DC=G3CAJA,DC=COM"
############################################################################################################################
################           Creación de la estructura de Usuarios Administrativos de la Organización         ################        
############################################################################################################################
$Administrativo = Import-CSV -Delimiter "," -Path ".\admin.csv"
foreach ($Usuario in $Administrativo)
{
    $Area = $Usuario.area
    $CodArea = $Area.substring(0,3)
    $Grupo = $CodArea + "-Administrators"
    $Ruta = $Usuario.Ruta
    $Clave = $Usuario.clave
    $Nombre = $Usuario.Nombre
    $Apellido = $Usuario.Apellido
    $NombreCompleto = $Usuario.Nombre + " " + $Usuario.Apellido
    $SAM = $Nombre + "." + $Apellido
    $UPN = $SAM + "@G3CAJA.com"
    New-AdUser $SAM -GivenName $Nombre -Surname $Apellido -DisplayName $NombreCompleto -Path $Ruta
    Set-ADAccountPassword $SAM -Reset -NewPassword (ConvertTO-SecureString -ASPlainText $Clave -Force)
    Add-ADGroupMember $Grupo -member $SAM
    Enable-ADAccount -identity $SAM
    Set-ADUser -identity $SAM -UserPrincipalName $UPN -SamAccountName $SAM
}
############################################################################################################################
################           Creación de la estructura de Usuarios Convencionales  de la Organización         ################        
############################################################################################################################
$Convencional = Import-CSV -Delimiter "," -Path ".\users.csv"
foreach ($Usuario in $Convencional)
{
    $Area = $Usuario.area
    $CodArea = $Area.substring(0,3)
    $Grupo = $CodArea + "-Usuarios"
    $Ruta = $Usuario.Ruta
    $Clave = $Usuario.clave
    $Nombre = $Usuario.Nombre
    $Apellido = $Usuario.Apellido
    $NombreCompleto = $Usuario.Nombre + " " + $Usuario.Apellido
    $SAM = $Nombre + "." + $Apellido
    $UPN = $SAM + "@G3CAJA.com"
    New-AdUser $SAM -GivenName $Nombre -Surname $Apellido -DisplayName $NombreCompleto -Path $Ruta
    Set-ADAccountPassword $SAM -Reset -NewPassword (ConvertTO-SecureString -ASPlainText $Clave -Force)
    Add-ADGroupMember $Grupo -member $SAM
    Enable-ADAccount -identity $SAM
    Set-ADUser -identity $SAM -UserPrincipalName $UPN -SamAccountName $SAM
}
