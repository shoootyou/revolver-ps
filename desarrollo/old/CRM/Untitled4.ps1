$query = new-object Microsoft.Xrm.Sdk.Query.QueryExpression("contact")
#$query.ColumnSet.AllColumns = $true

$query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet("emailaddress1","emailaddress2","emailaddress3")
#$query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet("contactid")
#$Query.Criteria.AddCondition("gendercode", [Microsoft.Xrm.Sdk.Query.ConditionOperator]::Equal, "2")
#$Query.ExtensionData = [Microsoft.Xrm.Sdk.Query.ColumnSet.ExtensionData]::BirthDate


$pageNumber = 1;
$query.PageInfo = New-Object -TypeName Microsoft.Xrm.Sdk.Query.PagingInfo;
$query.PageInfo.PageNumber = $pageNumber;
$query.PageInfo.Count = 5000;
$query.PageInfo.ReturnTotalRecordCount = $true
$query.PageInfo.PagingCookie = $null;
$DB_REB = Import-Csv $env:UserProfile\Desktop\Rebotes2.csv
$TMP_ENTITY = $null

$Found = 0
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
                            Write-host $Attribute.Key','$Attribute.Value
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
                            Write-host $Attribute.Key','$Attribute.Value
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
                            Write-host $Attribute.Key','$Attribute.Value
                            $Attribute.Key + ',' + $Attribute.Value | Out-File $env:UserProfile\Desktop\Log.csv -Append
                            $TMP_ENTITY = New-Object Microsoft.Xrm.Sdk.Entity("contact")
                            $contactMethodOption = new-object Microsoft.Xrm.Sdk.OptionSetValue(100000007)
                            $TMP_ENTITY.Attributes["new_tipoenvioemail3"] = [Microsoft.Xrm.Sdk.OptionSetValue]$contactMethodOption
                        }
                    }
                }
                if($TMP_ENTITY -ne $null){
                    if($attribute.Key -eq "contactid"){
                        Write-host $Attribute.Key','$Attribute.Value
                        $Attribute.Key + ',' + $Attribute.Value | Out-File $env:UserProfile\Desktop\Log.csv -Append
                        Write-Host "------------------------,------------------------"
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