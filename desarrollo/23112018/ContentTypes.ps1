$Subs = Invoke-WebRequest -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/list"
$Subs.Content