$Identidad = Get-User -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -like "*tcreatividad*") } | Select *
foreach ($Usuario in $Identidad) {

        Set-User -Identity $Usuario.WindowsEmailAddress -MobilePhone ""

}