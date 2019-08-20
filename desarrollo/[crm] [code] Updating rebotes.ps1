<##########################################################################################################################################################################
#                                                                        Global Functions
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
$query = new-object Microsoft.Xrm.Sdk.Query.QueryExpression("contact")
$query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet("emailaddress1","emailaddress2","emailaddress3")
$pageNumber = 1;
$query.PageInfo = New-Object -TypeName Microsoft.Xrm.Sdk.Query.PagingInfo;
$query.PageInfo.PageNumber = $pageNumber;
$query.PageInfo.Count = 5000;
$query.PageInfo.ReturnTotalRecordCount = $true
$query.PageInfo.PagingCookie = $null;
<##########################################################################################################################################################################
                                                                            Loading files
##########################################################################################################################################################################>
$DB_CLI_CSV_PT = Import-FilePath -Title "Tu archivo rebotes con cabecera de 'Rebote'"
$DB_REB = Import-Csv $DB_CLI_CSV_PT -Delimiter ","
$TMP_ENTITY = $null
$Found = 0
<##########################################################################################################################################################################
                                                                               Process
##########################################################################################################################################################################>
do{
    $response = $service.RetrieveMultiple($query)
    if ($response.Entities -ne $null){
        foreach($acct in $response.Entities){
            $ATRIB = $acct.Attributes
            foreach($attribute in $ATRIB){
                if($attribute.Key -eq "emailaddress1"){
                    foreach($REB in $DB_REB){
                        if($REB.Rebote -eq $attribute.Value){
                            $Found++
                            Write-host "The contact who have the mail"  $Attribute.Value " has been updated"
                            $Attribute.Key + ',' + $Attribute.Value | Out-File $env:UserProfile\Desktop\Log.csv -Append
                            $TMP_ENTITY = New-Object Microsoft.Xrm.Sdk.Entity("contact")
                            $contactMethodOption = new-object Microsoft.Xrm.Sdk.OptionSetValue(100000007)
                            $TMP_ENTITY.Attributes["new_tipoenvioemail"] = [Microsoft.Xrm.Sdk.OptionSetValue]$contactMethodOption
                        }
                    }
                }
                elseif($attribute.Key -eq "emailaddress2"){
                    foreach($REB in $DB_REB){
                        if($REB.Rebote -eq $attribute.Value){
                            $Found++
                            Write-host "The contact who have the mail"  $Attribute.Value " has been updated"
                            $Attribute.Key + ',' + $Attribute.Value | Out-File $env:UserProfile\Desktop\Log.csv -Append
                            $TMP_ENTITY = New-Object Microsoft.Xrm.Sdk.Entity("contact")
                            $contactMethodOption = new-object Microsoft.Xrm.Sdk.OptionSetValue(100000007)
                            $TMP_ENTITY.Attributes["new_tipoenvioemail2"] = [Microsoft.Xrm.Sdk.OptionSetValue]$contactMethodOption
                        }
                    }
                }
                elseif($attribute.Key -eq "emailaddress3"){
                    foreach($REB in $DB_REB){
                        if($REB.Rebote -eq $attribute.Value){
                            $Found++
                            Write-host "The contact who have the mail"  $Attribute.Value " has been updated"
                            $Attribute.Key + ',' + $Attribute.Value | Out-File $env:UserProfile\Desktop\Log.csv -Append
                            $TMP_ENTITY = New-Object Microsoft.Xrm.Sdk.Entity("contact")
                            $contactMethodOption = new-object Microsoft.Xrm.Sdk.OptionSetValue(100000007)
                            $TMP_ENTITY.Attributes["new_tipoenvioemail3"] = [Microsoft.Xrm.Sdk.OptionSetValue]$contactMethodOption
                        }
                    }
                }
                if($TMP_ENTITY -ne $null){
                    if($attribute.Key -eq "contactid"){
                        $Attribute.Key + ',' + $Attribute.Value | Out-File $env:UserProfile\Desktop\Log.csv -Append
                        Write-Host "-----------------------------------------------------------------------------------------------"
                        "------------------------" + ',' + "------------------------" | Out-File $env:UserProfile\Desktop\Log.csv -Append
                        $TMP_ENTITY.Id = [Guid]::Parse($attribute.Value)
                        $service.Update($TMP_ENTITY)
                        $TMP_ENTITY = $null
                    }
                }
            }
            
        }
    }
    if($response.MoreRecords){
        $query.PageInfo.PageNumber++
        $query.PageInfo.PagingCookie = $response.PagingCookie
    }
    else{
        break;
    }
    if($Found -eq $DB_REB.Count){
        break;
    }
}
while($true)