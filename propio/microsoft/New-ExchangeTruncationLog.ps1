#LIMMBX01
$MOV_DB_01 = Get-ChildItem G:\Logs\EMPPANDB11 | ? {$_.Extension -eq '.log' -and $_.LastAccessTime -lt (Get-Date).AddDays(-15)}
$MOV_DB_01 | % {Write-Progress -Id 1 -Activity “Copiando Información” -status (“Trabajando en " + $_.Name) -percentComplete ($i / $MOV_DB_01.count*100) ; Move-Item -Path $_.FullName -Destination T:\EMPPANDB11 }


#LIMMBX03
$MOV_DB_01 = Get-ChildItem G:\Logs\ADMNLIMDB05 | ? {$_.Extension -eq '.log' -and $_.LastAccessTime -lt (Get-Date).AddDays(-15)}
$MOV_DB_01 | % {Write-Progress -Id 1 -Activity "Limpiando Información” -status (“Trabajando en " + $_.Name) -percentComplete ($i / $MOV_DB_01.count*100) ;  Remove-Item -Path $_.FullName }