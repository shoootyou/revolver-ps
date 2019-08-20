$EntityQuery = New-Object Microsoft.Xrm.Sdk.Query.QueryByAttribute

$EntityQuery.EntityName([Microsoft.Xrm.Sdk.Query.QueryByAttribute.EntityName]::'contact')

$Service.Retrieve([Microsoft.Xrm.Sdk.Query.QueryByAttribute]::'contact')



querybyattribute.Attributes.AddRange("address1_city");

                    //  Value of queried attribute to return.
                    querybyattribute.Values.AddRange("Redmond");






<#Create query using QueryByAttribute.
QueryByAttribute querybyattribute = new QueryByAttribute("account") {
   ColumnSet = new ColumnSet("name", "address1_city", "emailaddress1"),
   Attributes.AddRange("address1_city"),
   Values.AddRange("Redmond")
};#>


$QueryByAttrib = New-Object Microsoft.Xrm.Sdk.Query.QueryByAttribute
$EntityQuery.EntityName([Microsoft.Xrm.Sdk.Query.QueryByAttribute.EntityName]::Equals('contact'))