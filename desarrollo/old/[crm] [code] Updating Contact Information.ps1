<##########################################################################################################################################################################
                                                                         Global Functions
##########################################################################################################################################################################>

function Import-FilePath{
    param(
            [string]$Title = "Find your CSV",
            [string]$Filter = 'CSV (Comma delimited *.csv)|*.csv',
            [string]$Path = $env:USERPROFILE
    )


	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$obj_imp_path = New-Object System.Windows.Forms.OpenFileDialog
	$obj_imp_path.InitialDirectory = $Path
	$obj_imp_path.Filter = $Filter
	$obj_imp_path.Title = $Title
	$Show = $obj_imp_path.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $obj_imp_path.FileName
	}
}

<##########################################################################################################################################################################
                                                                          Global Variables
##########################################################################################################################################################################>

$DB_ACC_CSV_PT = Import-FilePath -Title "Find the CSV of Contacts DB"
$DB_ACC_UP = Import-Csv -Path $DB_ACC_CSV_PT -Delimiter "," -Encoding Default
$GBL_COU = 1

<##########################################################################################################################################################################
                                                                          Creating logs file
##########################################################################################################################################################################>

"ID,Firstname,Lastname" | Out-File $ENV:UserProfile\Desktop\log.csv -Append

<##########################################################################################################################################################################
#                                                                             Process
##########################################################################################################################################################################>

Foreach($ACC_INF in $DB_ACC_UP) {
    $TMP_ENTITY = New-Object Microsoft.Xrm.Sdk.Entity("contact")

    $TMP_ENTITY.Id = [Guid]::Parse($ACC_INF.id)
    $TMP_ENTITY.Attributes["firstname"] = $ACC_INF.firstname
    $TMP_ENTITY.Attributes["middlename"] = $ACC_INF.secondname
    $TMP_ENTITY.Attributes["lastname"] = $ACC_INF.lastname
    $TMP_ENTITY.Attributes["new_segundoapellido"] = $ACC_INF.secondlastname

    $TMP_FULL_NAME = $ACC_INF.firstname + " " + $ACC_INF.lastname

    Write-Output ('{0},{1},{2}' -f $ACC_INF.Id,$ACC_INF.firstname,$ACC_INF.lastname) | Out-File $ENV:UserProfile\Desktop\log.csv -Append
    
    Write-Progress -Activity “Updating Information” -status “Updating $TMP_FULL_NAME” -percentComplete ($GBL_COU / $DB_ACC_UP.Count*100)
    
    $GBL_COU++

    $service.Update($TMP_ENTITY)
}