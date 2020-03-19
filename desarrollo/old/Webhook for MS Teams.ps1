# ---------------------------------------------------
# Script: C:\Users\stefstr\Microsoft\OneDrive - Microsoft\Scripts\PS\MicrosoftTeams\testwebhook_v2.ps1
# Version: 0.1
# Author: Stefan Stranger
# Date: 11/03/2016 10:48:58
# Description: Call Microsoft Teams Incoming Webhook from PowerShell
# Comments:
# Changes:  
# Disclaimer: 
# This example is provided “AS IS” with no warranty expressed or implied. Run at your own risk. 
# **Always test in your lab first**  Do this at your own risk!! 
# The author will not be held responsible for any damage you incur when making these changes!
# ---------------------------------------------------


$webhook = 'https://outlook.office.com/webhook/f86423b6-fcbe-4643-8d52-3b9ef22855e3@53560dc4-ccc8-496e-a9e7-a60f54bfecd0/IncomingWebhook/2892e246cc164a43aa3c07958708bf4e/b841c37e-618a-4af2-9068-9bf603a4e952'

$Body = @{
        'text'= 'Mensaje de Prueba desde PowerShell'
}

$params = @{
    Headers = @{'accept'='application/json'}
    Body = $Body | convertto-json
    Method = 'Post'
    URI = $webhook 
}

Invoke-RestMethod @params