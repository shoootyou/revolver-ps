##############################################################################################################################
## Autor: Rodolfo Castelo Méndez                                                                                            ##
## Script para la creación de la estructura de la oganización                                                               ##
## Nombre de la organización: G3BANK.COM                                                                                    ##
## Nombres de los Computadores: (Nomenclatura del área)-(Número correlativo)                                                ##
## Versión 1.02                                                                                                             ##
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
##############################################################################################################################
################                    Creación de la estructura de Usuarios de la Organización                  ################        
##############################################################################################################################
### Area de Ingenieria
### Principal
##############################################################################################################################
#1
New-ADUser rodolfo.castelo -GivenName Rodolfo -Surname Castelo -DisplayName "Rodolfo Castelo" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "rodolfo.castelo"
Set-ADAccountPassword rodolfo.castelo -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "rodolfo.castelo"
Set-ADUser -identity "rodolfo.castelo" -UserPrincipalName rodolfo.castelo@g3bank.com
#2
New-ADUser roy.puente -GivenName Roy -Surname Puente -DisplayName "Roy Perez" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "roy.puente"
Set-ADAccountPassword roy.puente -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "roy.puente"
Set-ADUser -identity "roy.puente" -UserPrincipalName roy.puente@g3bank.com
#3
New-ADUser yvel.cruz -GivenName Yvel -Surname Cruz -DisplayName "Yvel Cruz" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "yvel.cruz"
Set-ADAccountPassword yvel.cruz -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "yvel.cruz"
Set-ADUser -identity "yvel.cruz" -UserPrincipalName yvel.cruz@g3bank.com
#4
New-ADUser alex.garcia -GivenName Alex -Surname Garcia -DisplayName "Alex Garcia" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "alex.garcia"
Set-ADAccountPassword alex.garcia -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "alex.garcia"
Set-ADUser -identity "alex.garcia" -UserPrincipalName alex.garcia@g3bank.com
#5
New-ADUser spencer.camacho -GivenName Spencer -Surname Camacho -DisplayName "Spencer Camacho" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "spencer.camacho"
Set-ADAccountPassword spencer.camacho -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "spencer.camacho"
Set-ADUser -identity "spencer.camacho" -UserPrincipalName spencer.camacho@g3bank.com
#6
New-ADUser eddy.saldaña -GivenName Eddy -Surname Saldaña -DisplayName "Eddy Saldaña" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "eddy.saldaña"
Set-ADAccountPassword eddy.saldaña -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "eddy.saldaña"
Set-ADUser -identity "eddy.saldaña" -UserPrincipalName eddy.saldaña@g3bank.com
#7
New-ADUser raul.mendez -GivenName Raul -Surname Mendez -DisplayName "Raul Mendez" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "raul.mendez"
Set-ADAccountPassword raul.mendez -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "raul.mendez"
Set-ADUser -identity "raul.mendez" -UserPrincipalName raul.mendez@g3bank.com
#8
New-ADUser ana.mendez -GivenName Ana -Surname Mendez -DisplayName "Ana Mendez" -Path "OU=Principal,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ING-Administrators" -member "ana.mendez"
Set-ADAccountPassword ana.mendez -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "ana.mendez"
Set-ADUser -identity "ana.mendez" -UserPrincipalName ana.mendez@g3bank.com
##############################################################################################################################
### Area de Ingenieria
### Area Secundaria
##############################################################################################################################
#1
New-ADUser gerardo.avila -GivenName Gerardo -Surname Avila -DisplayName "Gerardo Avila" -Path "OU=Secundario,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ADM-Administrators" -member "gerardo.avila"
Set-ADAccountPassword gerardo.avila -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "gerardo.avila"
Set-ADUser -identity "gerardo.avila" -UserPrincipalName gerardo.avila@g3bank.com
#2
New-ADUser juan.mendoza -GivenName Juan -Surname Mendoza -DisplayName "Juan Mendoza" -Path "OU=Secundario,OU=Usuarios,OU=Ingenieria,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "ADM-Administrators" -member "juan.mendoza"
Set-ADAccountPassword juan.mendoza -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "juan.mendoza"
Set-ADUser -identity "juan.mendoza" -UserPrincipalName juan.mendoza@g3bank.com
##############################################################################################################################
##############################################################################################################################
### Area de Contabilidad
### Area Principal
##############################################################################################################################
#1
New-ADUser adonis.perez -GivenName Adonis -Surname Perez -DisplayName "Adonis Perez" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Administrators" -member "adonis.perez"
Set-ADAccountPassword adonis.perez -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "adonis.perez"
Set-ADUser -identity "adonis.perez" -UserPrincipalName adonis.perez@g3bank.com
#2
New-ADUser luis.guzman -GivenName Luis -Surname Guzman -DisplayName "Luis Guzman" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "luis.guzman"
Set-ADAccountPassword luis.guzman -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "luis.guzman"
Set-ADUser -identity "luis.guzman" -UserPrincipalName luis.guzman@g3bank.com
#3
New-ADUser rossana.reccio -GivenName Rossana -Surname Reccio -DisplayName "Rossana Reccio" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "rossana.reccio"
Set-ADAccountPassword rossana.reccio -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "rossana.reccio"
Set-ADUser -identity "rossana.reccio" -UserPrincipalName rossana.reccioz@g3bank.com
#4
New-ADUser patricia.navarro -GivenName Patricia -Surname Navarro -DisplayName "Patricia Navarro" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "patricia.navarro"
Set-ADAccountPassword patricia.navarro -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "patricia.navarro"
Set-ADUser -identity "patricia.navarro" -UserPrincipalName patricia.navarro@g3bank.com
#5
New-ADUser roberto.robles -GivenName Roberto -Surname Robles -DisplayName "Roberto Robles" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "roberto.robles"
Set-ADAccountPassword roberto.robles -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "roberto.robles"
Set-ADUser -identity "roberto.robles" -UserPrincipalName roberto.robles@g3bank.com
#6
New-ADUser asuncion.gonzales -GivenName Asunción -Surname Gonzales -DisplayName "Asunción Gonzales" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "asuncion.gonzales"
Set-ADAccountPassword asuncion.gonzales -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "asuncion.gonzales"
Set-ADUser -identity "asuncion.gonzales" -UserPrincipalName asuncion.gonzales@g3bank.com
#7
New-ADUser gina.padilla -GivenName Gina -Surname Padilla -DisplayName "Gina Padilla" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "gina.padilla"
Set-ADAccountPassword gina.padilla -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "gina.padilla"
Set-ADUser -identity "gina.padilla" -UserPrincipalName gina.padilla@g3bank.com
#8
New-ADUser rosio.yabar -GivenName Rosio -Surname Yabar -DisplayName "Rosio Yabar" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "rosio.yabar"
Set-ADAccountPassword rosio.yabar -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "rosio.yabar"
Set-ADUser -identity "rosio.yabar" -UserPrincipalName rosio.yabar@g3bank.com
#9
New-ADUser jhanet.arango -GivenName Jhanet -Surname Arango -DisplayName "Jhanet Arango" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "jhanet.arango"
Set-ADAccountPassword jhanet.arango -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "jhanet.arango"
Set-ADUser -identity "jhanet.arango" -UserPrincipalName jhanet.arango@g3bank.com
#10
New-ADUser milagros.perez -GivenName Milagros -Surname Perez -DisplayName "Milagros Perez" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "milagros.perez"
Set-ADAccountPassword milagros.perez -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "milagros.perez"
Set-ADUser -identity "milagros.perez" -UserPrincipalName milagros.perez@g3bank.com
#11
New-ADUser helen.robles -GivenName Helen -Surname Robles -DisplayName "Helen Robles" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "helen.robles"
Set-ADAccountPassword helen.robles -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "helen.robles"
Set-ADUser -identity "helen.robles" -UserPrincipalName helen.robles@g3bank.com
#12
New-ADUser mery.vilca -GivenName Mery -Surname Vilca -DisplayName "Mery Vilca" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "mery.vilca"
Set-ADAccountPassword mery.vilca -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "mery.vilca"
Set-ADUser -identity "mery.vilca" -UserPrincipalName mery.vilca@g3bank.com
#13
New-ADUser carlos.farach -GivenName Carlos -Surname Farach -DisplayName "Carlos Farach" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "carlos.farach"
Set-ADAccountPassword carlos.farach -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "carlos.farach"
Set-ADUser -identity "carlos.farach" -UserPrincipalName carlos.farach@g3bank.com
#14
New-ADUser diego.patino -GivenName Diego -Surname Patino -DisplayName "Diego Patino" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "diego.patino"
Set-ADAccountPassword diego.patino -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "diego.patino"
Set-ADUser -identity "diego.patino" -UserPrincipalName diego.patino@g3bank.com
#15
New-ADUser jessica.cordova -GivenName Jessica -Surname Cordova -DisplayName "Jessica Cordova" -Path "OU=Principal,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "jessica.cordova"
Set-ADAccountPassword jessica.cordova -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "jessica.cordova"
Set-ADUser -identity "jessica.cordova" -UserPrincipalNamejessica.cordovar@g3bank.com
##############################################################################################################################
### Area de Contabilidad
### Area Secundaria
##############################################################################################################################
#1
New-ADUser larry.concha -GivenName Larry -Surname Concha -DisplayName "Larry Concha" -Path "OU=Secundario,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Administrators" -member "larry.concha"
Set-ADAccountPassword larry.concha -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "larry.concha"
Set-ADUser -identity "larry.concha" -UserPrincipalName larry.concha@g3bank.com
#2
New-ADUser malcom.quiroz -GivenName Malcom -Surname Quiroz -DisplayName "Malcom Quiroz" -Path "OU=Secundario,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "malcom.quiroz"
Set-ADAccountPassword malcom.quiroz -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "malcom.quiroz"
Set-ADUser -identity "malcom.quiroz" -UserPrincipalName malcom.quiroz@g3bank.com
#3
New-ADUser fiama.ochoa -GivenName Fiama -Surname Ochoa -DisplayName "Fiama Ochoa" -Path "OU=Secundario,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "fiama.ochoa"
Set-ADAccountPassword fiama.ochoa -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "fiama.ochoa"
Set-ADUser -identity "fiama.ochoa" -UserPrincipalName fiama.ochoa@g3bank.com
#4
New-ADUser alfredo.cautin -GivenName Alfredo -Surname Cautin -DisplayName "Alfredo Cautin" -Path "OU=Secundario,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM"
Add-ADGroupMember "CON-Usuarios" -member "alfredo.cautin"
Set-ADAccountPassword alfredo.cautin -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "alfredo.cautin"
Set-ADUser -identity "alfredo.cautin" -UserPrincipalName alfredo.cautin@g3bank.com
#5
New-ADUser rafael.zagarra -GivenName Rafael -Surname Zagarra -DisplayName "Rafael Zagarra" -Path "OU=Secundario,OU=Usuarios,OU=Contabilidad,OU=Internal,DC=G3BANK,DC=COM" 
Add-ADGroupMember "CON-Usuarios" -member "rafael.zagarra"
Set-ADAccountPassword rafael.zagarra -Reset -NewPassword (ConvertTO-SecureString -ASPlainText "P@ssw0rd" -Force)
Enable-ADAccount -identity "rafael.zagarra"
Set-ADUser -identity "rafael.zagarra" -UserPrincipalName rafael.zagarra@g3bank.com
##############################################################################################################################
##############################################################################################################################
### Area de Administración
### Area Principal
##############################################################################################################################