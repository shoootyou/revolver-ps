$REQ = Invoke-WebRequest -Method Post -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/start?contentType=Audit.General"
$TMP = Invoke-WebRequest -Method GET -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/content?contentType=Audit.General"
$TMP.Content

# Get a content blob
$uri = 'https://manage.office.com/api/v1.0/e57bf946-a543-421d-9e8a-745f9a25e38d/activity/feed/audit/20181123153750437144785$20181123211708832042073$audit_general$Audit_General'
$contents = Invoke-WebRequest -Method GET -Headers $headerParams -Uri $uri
$contents.Content | Out-File .\Downloads\4.json