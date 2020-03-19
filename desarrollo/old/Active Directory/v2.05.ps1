##############################################################################################################################
## Autor: Rodolfo Castelo Méndez                                                                                            ##
## Script para la creación de la estructura de la oganización                                                               ##
## Nombre de la organización: G3BANK.COM                                                                                    ##
## Nombres de los Computadores: (Nomenclatura del área)-(Número correlativo)                                                ##
## Versión 2.05                                                                                                             ##
##############################################################################################################################
################                         Importación del Módulo del Directorio Activo                         ################        
##############################################################################################################################
Import-Module ActiveDirectory
##############################################################################################################################
################                               Creación de la estructura de OUs                               ################        
##############################################################################################################################
New-ADOrganizationalUnit -Path "DC=G3BANK,DC=COM" Internal
# Área de Servidores
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3BANK,DC=COM" Servidores
New-ADOrganizationalUnit -Path "OU=Servidores,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Servidores,OU=Internal,DC=G3BANK,DC=COM" Secundario
# Area de la DMZ
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3BANK,DC=COM" DMZ
New-ADOrganizationalUnit -Path "OU=DMZ,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=DMZ,OU=Internal,DC=G3BANK,DC=COM" Secundario
# Área de ingenieria
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3BANK,DC=COM" Ingenieria
New-ADOrganizationalUnit -Path "OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM" Secundario
# Área de Contabilidad
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3BANK,DC=COM" Contabilidad
New-ADOrganizationalUnit -Path "OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" Secundario
# Área de Administración
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3BANK,DC=COM" Administracion
New-ADOrganizationalUnit -Path "OU=Administracion,OU=Internal,DC=G3BANK,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Administracion,OU=Internal,DC=G3BANK,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM" Secundario
# Área de Ventas
New-ADOrganizationalUnit -Path "OU=Internal,DC=G3BANK,DC=COM" Ventas
New-ADOrganizationalUnit -Path "OU=Ventas,OU=Internal,DC=G3BANK,DC=COM" Usuarios
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Usuarios,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM" Secundario
New-ADOrganizationalUnit -Path "OU=Ventas,OU=Internal,DC=G3BANK,DC=COM" Computadores
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM" Principal
New-ADOrganizationalUnit -Path "OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM" Secundario
##############################################################################################################################
################                   Creación de los objetos de computador de la Organización                   ################        
##############################################################################################################################
# Area de Ingenieria
New-Adcomputer –name "ING-PC-01" –SamAccountName "ING-PC-01" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-02" –SamAccountName "ING-PC-02" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-03" –SamAccountName "ING-PC-03" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-04" –SamAccountName "ING-PC-04" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-05" –SamAccountName "ING-PC-05" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-06" –SamAccountName "ING-PC-06" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-07" –SamAccountName "ING-PC-07" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-08" –SamAccountName "ING-PC-08" -Path "OU=Principal,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-09" –SamAccountName "ING-PC-09" -Path "OU=Secundario,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ING-PC-10" –SamAccountName "ING-PC-10" -Path "OU=Secundario,OU=Computadores,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
# Area de Contabilidad
New-Adcomputer –name "CON-PC-01" –SamAccountName "CON-PC-01" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-02" –SamAccountName "CON-PC-02" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-03" –SamAccountName "CON-PC-03" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-04" –SamAccountName "CON-PC-04" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-05" –SamAccountName "CON-PC-05" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-06" –SamAccountName "CON-PC-06" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-07" –SamAccountName "CON-PC-07" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-08" –SamAccountName "CON-PC-08" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-09" –SamAccountName "CON-PC-09" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-10" –SamAccountName "CON-PC-10" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-11" –SamAccountName "CON-PC-11" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-12" –SamAccountName "CON-PC-12" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-13" –SamAccountName "CON-PC-13" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-14" –SamAccountName "CON-PC-14" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-15" –SamAccountName "CON-PC-15" -Path "OU=Principal,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-16" –SamAccountName "CON-PC-16" -Path "OU=Secundario,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-17" –SamAccountName "CON-PC-17" -Path "OU=Secundario,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-18" –SamAccountName "CON-PC-18" -Path "OU=Secundario,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-19" –SamAccountName "CON-PC-19" -Path "OU=Secundario,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "CON-PC-20" –SamAccountName "CON-PC-20" -Path "OU=Secundario,OU=Computadores,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
#  Area de Administración
New-Adcomputer –name "ADM-PC-01" –SamAccountName "ADM-PC-01" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-02" –SamAccountName "ADM-PC-02" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-03" –SamAccountName "ADM-PC-03" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-04" –SamAccountName "ADM-PC-04" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-05" –SamAccountName "ADM-PC-05" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-06" –SamAccountName "ADM-PC-06" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-07" –SamAccountName "ADM-PC-07" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-08" –SamAccountName "ADM-PC-08" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-09" –SamAccountName "ADM-PC-09" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-10" –SamAccountName "ADM-PC-10" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-11" –SamAccountName "ADM-PC-11" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-12" –SamAccountName "ADM-PC-12" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-13" –SamAccountName "ADM-PC-13" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-14" –SamAccountName "ADM-PC-14" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-15" –SamAccountName "ADM-PC-15" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-16" –SamAccountName "ADM-PC-16" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-17" –SamAccountName "ADM-PC-17" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-18" –SamAccountName "ADM-PC-18" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-19" –SamAccountName "ADM-PC-19" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-20" –SamAccountName "ADM-PC-20" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-21" –SamAccountName "ADM-PC-21" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-22" –SamAccountName "ADM-PC-22" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-23" –SamAccountName "ADM-PC-23" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-24" –SamAccountName "ADM-PC-24" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-25" –SamAccountName "ADM-PC-25" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-26" –SamAccountName "ADM-PC-26" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-27" –SamAccountName "ADM-PC-27" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-28" –SamAccountName "ADM-PC-28" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-29" –SamAccountName "ADM-PC-29" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-30" –SamAccountName "ADM-PC-30" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-31" –SamAccountName "ADM-PC-31" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-32" –SamAccountName "ADM-PC-32" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-33" –SamAccountName "ADM-PC-33" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-34" –SamAccountName "ADM-PC-34" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-35" –SamAccountName "ADM-PC-35" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-36" –SamAccountName "ADM-PC-36" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-37" –SamAccountName "ADM-PC-37" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-38" –SamAccountName "ADM-PC-38" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-39" –SamAccountName "ADM-PC-39" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-40" –SamAccountName "ADM-PC-40" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-41" –SamAccountName "ADM-PC-41" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-42" –SamAccountName "ADM-PC-42" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-43" –SamAccountName "ADM-PC-43" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-44" –SamAccountName "ADM-PC-44" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-45" –SamAccountName "ADM-PC-45" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-46" –SamAccountName "ADM-PC-46" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-47" –SamAccountName "ADM-PC-47" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-48" –SamAccountName "ADM-PC-48" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-49" –SamAccountName "ADM-PC-49" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-50" –SamAccountName "ADM-PC-50" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-51" –SamAccountName "ADM-PC-51" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-52" –SamAccountName "ADM-PC-52" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-53" –SamAccountName "ADM-PC-53" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-54" –SamAccountName "ADM-PC-54" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-55" –SamAccountName "ADM-PC-55" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-56" –SamAccountName "ADM-PC-56" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-57" –SamAccountName "ADM-PC-57" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-58" –SamAccountName "ADM-PC-58" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-59" –SamAccountName "ADM-PC-59" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-60" –SamAccountName "ADM-PC-60" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-61" –SamAccountName "ADM-PC-61" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-62" –SamAccountName "ADM-PC-62" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-63" –SamAccountName "ADM-PC-63" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-64" –SamAccountName "ADM-PC-64" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-65" –SamAccountName "ADM-PC-65" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-66" –SamAccountName "ADM-PC-66" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-67" –SamAccountName "ADM-PC-67" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-68" –SamAccountName "ADM-PC-68" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-69" –SamAccountName "ADM-PC-69" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-70" –SamAccountName "ADM-PC-70" -Path "OU=Principal,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-71" –SamAccountName "ADM-PC-71" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-72" –SamAccountName "ADM-PC-72" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-73" –SamAccountName "ADM-PC-73" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-74" –SamAccountName "ADM-PC-74" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-75" –SamAccountName "ADM-PC-75" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-76" –SamAccountName "ADM-PC-76" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-77" –SamAccountName "ADM-PC-77" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-78" –SamAccountName "ADM-PC-78" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-79" –SamAccountName "ADM-PC-79" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-80" –SamAccountName "ADM-PC-80" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-81" –SamAccountName "ADM-PC-81" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-82" –SamAccountName "ADM-PC-82" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-83" –SamAccountName "ADM-PC-83" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-84" –SamAccountName "ADM-PC-84" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-85" –SamAccountName "ADM-PC-85" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-86" –SamAccountName "ADM-PC-86" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-87" –SamAccountName "ADM-PC-87" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-88" –SamAccountName "ADM-PC-88" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-89" –SamAccountName "ADM-PC-89" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "ADM-PC-90" –SamAccountName "ADM-PC-90" -Path "OU=Secundario,OU=Computadores,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
# Area de Ventas
New-Adcomputer –name "VEN-PC-01" –SamAccountName "VEN-PC-01" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-02" –SamAccountName "VEN-PC-02" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-03" –SamAccountName "VEN-PC-03" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-04" –SamAccountName "VEN-PC-04" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-05" –SamAccountName "VEN-PC-05" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-06" –SamAccountName "VEN-PC-06" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-07" –SamAccountName "VEN-PC-07" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-08" –SamAccountName "VEN-PC-08" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-09" –SamAccountName "VEN-PC-09" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-10" –SamAccountName "VEN-PC-10" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-11" –SamAccountName "VEN-PC-11" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-12" –SamAccountName "VEN-PC-12" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-13" –SamAccountName "VEN-PC-13" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-14" –SamAccountName "VEN-PC-14" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-15" –SamAccountName "VEN-PC-15" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-16" –SamAccountName "VEN-PC-16" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-17" –SamAccountName "VEN-PC-17" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-18" –SamAccountName "VEN-PC-18" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-19" –SamAccountName "VEN-PC-19" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-20" –SamAccountName "VEN-PC-20" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-21" –SamAccountName "VEN-PC-21" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-22" –SamAccountName "VEN-PC-22" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-23" –SamAccountName "VEN-PC-23" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-24" –SamAccountName "VEN-PC-24" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-25" –SamAccountName "VEN-PC-25" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-26" –SamAccountName "VEN-PC-26" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-27" –SamAccountName "VEN-PC-27" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-28" –SamAccountName "VEN-PC-28" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-29" –SamAccountName "VEN-PC-29" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-30" –SamAccountName "VEN-PC-30" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-31" –SamAccountName "VEN-PC-31" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-32" –SamAccountName "VEN-PC-32" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-33" –SamAccountName "VEN-PC-33" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-34" –SamAccountName "VEN-PC-34" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-35" –SamAccountName "VEN-PC-35" -Path "OU=Principal,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-36" –SamAccountName "VEN-PC-36" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-37" –SamAccountName "VEN-PC-37" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-38" –SamAccountName "VEN-PC-38" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-39" –SamAccountName "VEN-PC-39" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-40" –SamAccountName "VEN-PC-40" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-41" –SamAccountName "VEN-PC-41" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-42" –SamAccountName "VEN-PC-42" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-43" –SamAccountName "VEN-PC-43" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-44" –SamAccountName "VEN-PC-44" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-45" –SamAccountName "VEN-PC-45" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-46" –SamAccountName "VEN-PC-46" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-47" –SamAccountName "VEN-PC-47" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-48" –SamAccountName "VEN-PC-48" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-49" –SamAccountName "VEN-PC-49" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-Adcomputer –name "VEN-PC-50" –SamAccountName "VEN-PC-50" -Path "OU=Secundario,OU=Computadores,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
##############################################################################################################################
################                    Creación de la estructura de Grupos de la Organización                    ################        
##############################################################################################################################
# Área de ingenieria
New-ADGroup -Name ING-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
# Área de Contabilidad
New-ADGroup -Name CON-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
New-ADGroup -Name CON-Usuarios -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
# Área de Administración
New-ADGroup -Name ADM-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
New-ADGroup -Name ADM-Usuarios -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Administracion,OU=Internal,DC=G3BANK,DC=COM"
# Área de Ventas
New-ADGroup -Name VEN-Administrators -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
New-ADGroup -Name VEN-Usuarios -GroupScope Global -Path "OU=Principal,OU=Usuarios,OU=Ventas,OU=Internal,DC=G3BANK,DC=COM"
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
    $UPN = $SAM + "@g3bank.com"
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
    $UPN = $SAM + "@g3bank.com"
    New-AdUser $SAM -GivenName $Nombre -Surname $Apellido -DisplayName $NombreCompleto -Path $Ruta
    Set-ADAccountPassword $SAM -Reset -NewPassword (ConvertTO-SecureString -ASPlainText $Clave -Force)
    Add-ADGroupMember $Grupo -member $SAM
    Enable-ADAccount -identity $SAM
    Set-ADUser -identity $SAM -UserPrincipalName $UPN -SamAccountName $SAM
}
