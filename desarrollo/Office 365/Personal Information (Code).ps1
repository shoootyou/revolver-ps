##################################################
#### Información de ubicación variables
#################################################
$SiUbica = New-Object System.Management.Automation.Host.ChoiceDescription "&Sí", ` "Sí"
$NoUbica = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "No"
$OpcionesUbica = [System.Management.Automation.Host.ChoiceDescription[]]($SiUbica, $NoUbica)
$ResultadosUbica = $host.ui.PromptForChoice("Confirmación", "¿Desea verificar la información de ubicación?", $OpcionesUbica, 0) 
#################################################
$UsuarioPrimario = Get-User -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -like "$VariableConsulta")} | Select *
$OUserOut = @()
foreach ($OUser in $UsuarioPrimario) {

    $ObjProperties = New-Object PSObject
             
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Identity -Value $OUser.Identity
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name DisplayName -Value $OUser.DisplayName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name FirstName -Value $OUser.FirstName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name LastName -Value $OUser.LastName
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Office -Value $OUser.office
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Phone -Value $OUser.Phone
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name WindowsEmailAddress -Value $OUser.WindowsEmailAddress
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Department -Value $OUser.Department
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Title -Value $OUser.Title
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Manager -Value $OUser.Manager
    Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name Company -Value $OUser.Company

    switch ($ResultadosUbica){
            0 {
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name PostalCode -Value $OUser.PostalCode
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name StreetAddress -Value $OUser.StreetAddress        
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name City -Value $OUser.City
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name StateOrProvince -Value $OUser.StateOrProvince
            Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name CountryOrRegion -Value $OUser.CountryOrRegion
            }
            1 {
        
            }
        }
    $OUserOut += $ObjProperties
}
$OUserOut | Out-GridView -Title "Información Personal"
$OUserOut |  Export-Csv C:\Users\Rodolfo\Desktop\Exportación.csv    