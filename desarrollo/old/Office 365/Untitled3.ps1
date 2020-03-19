$query = new-object Microsoft.Xrm.Sdk.Query.QueryExpression("contact")
#$query.ColumnSet.AllColumns = $true
#$query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet("firstname","fullname","birthdate")
$query.ColumnSet = new-object Microsoft.Xrm.Sdk.Query.ColumnSet("contactid")
#$Query.Criteria.AddCondition("gendercode", [Microsoft.Xrm.Sdk.Query.ConditionOperator]::Equal, "2")
#$Query.ExtensionData = [Microsoft.Xrm.Sdk.Query.ColumnSet.ExtensionData]::BirthDate


$pageNumber = 1;
$query.PageInfo = New-Object -TypeName Microsoft.Xrm.Sdk.Query.PagingInfo;
$query.PageInfo.PageNumber = $pageNumber;
$query.PageInfo.Count = 1;
$query.PageInfo.PagingCookie = $null;


<#
("address1_stateorprovince", ConditionOperator.Equal, "WA"),
("address1_city", ConditionOperator.In, new String[] {"Redmond", "Bellevue" , "Kirkland", "Seattle"}),
("createdon", ConditionOperator.LastXDays, 30),
("emailaddress1", ConditionOperator.NotNull)
#>

# RetrieveMultiple returns a maximum of 5000 records by default. 
# If you need more, use the response's PagingCookie.
$response = $service.RetrieveMultiple($query)

$i = 0
$MTX_Account = $response.Entities
foreach($IN_MTX_ACC in $MTX_Account){
    $i++
    $IN_MTX_ATTR = $IN_MTX_ACC.Attributes
    foreach($Attribute in $IN_MTX_ATTR){
        $Attribute.Key + ': ' + $Attribute.Value
    }
    $IN_MTX_FVAL = $IN_MTX_ACC.FormattedValues
    foreach($FormatValue in $IN_MTX_FVAL){
        $FormatValue.Key + ': ' + $FormatValue.Value
    }
    Write-Host "--------------------"
    
}
Write-Host $i


$Retrieve = New-Object Microsoft.Xrm.Sdk.Messages.RetrieveRequest
$Retrieve.Target = New-Object Microsoft.Xrm.Sdk.EntityReference("contact",'91a58dc8-7feb-e311-a91c-6c3be5a8a0c8')
$Retrieve.ColumnSet = New-Object Microsoft.Xrm.Sdk.Query.ColumnSet("birthdate")


$TMP_01 = $Conecting.Execute($Retrieve)

$TMP_01.Entity.Attributes
