# Create app of type Web app / API in Azure AD, generate a Client Secret, and update the client id and client secret here
$ClientID = "fd97fbfa-10cc-4e31-9b50-3695a92eba00"
$ClientSecret = "kh0iqZkFG7fwPabr9Tbul1MqY+hmoZ0VWL6nUGNmmT0="
$loginURL = "https://login.microsoftonline.com/"
$tenantdomain = "revocorptest01.onmicrosoft.com"
# Get the tenant GUID from Properties | Directory ID under the Azure Active Directory section
$TenantGUID = "e57bf946-a543-421d-9e8a-745f9a25e38d"
$resource = "https://manage.office.com"
# auth
$body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
$oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}