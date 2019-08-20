function Get-SharedFolders{

    $DB_SHA = Get-ChildItem HKLM:\System\CurrentControlSet\Services\LanManServer\Shares\
    $DB_PRO = Get-ItemProperty HKLM:\System\CurrentControlSet\Services\LanManServer\Shares\
    $RT_ARR = @()
    foreach($SH in ([array]$DB_SHA.Property)){
        $SH_NAM = $SH
        $DB_PAT = $DB_PRO.$SH_NAM
        foreach($PRO in $DB_PAT){
            if($PRO -like "Path=*"){
                $RT_ARR += $PRO.Replace("Path=","")   
            }
        }
    }
    return $RT_ARR
}
function New-SharedFoldersReport{
    param(
        [array]$SharedFolders
    )
    $GR_OUT = @()
    $I = 1
    foreach($FLD in $SharedFolders){
        Write-Progress -Id 1 -Activity “Obteniendo información” -status “Trabajando en $FLD” -percentComplete ($i / $SharedFolders.count*100)
        $FIL = Get-ChildItem $FLD -Recurse -ErrorAction SilentlyContinue | Select Name,LastAccessTime | Sort-Object LastAccessTime        
        $GR_TMP = New-Object PsObject 
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name Compartido -Value $FLD
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Último archivo accedido" -Value $FIL[0].Name
        Add-Member -InputObject $GR_TMP -MemberType NoteProperty -Name "Fecha de último acceso" -Value $FIL[0].LastAccessTime
        $GR_OUT += $GR_TMP
        $I++
    }
    return $GR_OUT 
}
New-SharedFoldersReport -SharedFolders (Get-SharedFolders)