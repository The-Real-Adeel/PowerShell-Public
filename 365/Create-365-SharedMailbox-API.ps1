###########################################
# Authentication to retrieve Bearer Token #
###########################################

$uriForAuth = 'https://login.windows.net/<tenantID>/oauth2/token'
$HeaderForAuth = @{'accept'='application/json'}
$bodyForAuth = @{
    'grant_type' = 'client_credentials'
    'resource'= 'https://management.core.windows.net/'
    'client_id' = '<enter service principal id>'
    'client_secret' = '<enter client secret>'
}
$token = Invoke-RestMethod -Method post -Uri $uriForAuth -Body $bodyForAuth -Headers $HeaderForAuth

##############################################################################
# Use Bearer Token and Initalize the runbook with parameters to pass through #
##############################################################################
# FullAccess, SendAs & SendBehalfOf can be Null. Shared Mailbox & DisplayName must be mandatory

$JobId = [GUID]::NewGuid().ToString()
$URI = "https://management.azure.com/subscriptions/<subscriptionID>/resourceGroups/<resourceGroupName>/providers/Microsoft.Automation/automationAccounts/<automationAccountName>/jobs/${JobId}?api-version=2019-06-01"
$Headers = @{"Authorization" = "$($Token.token_type) "+ "$($Token.access_token)"}
$Body = @"
        {
           "properties":{
           "runbook":{
               "name":"create_365_shared_mailbox"
           },
           "runOn":"<HybirdWorkerGroup>", 
           "parameters":{
                "SharedMailbox":"<Email>",
                "DisplayName":"<DisplayName>",
                "FullAccess":"<email>", 
                "SendAs":"<email>, <email2>",
                "SendBehalfOf":"<email>"
           }
          }
       }
"@
Invoke-RestMethod -Method PUT -Uri $Uri -body $body -Headers $headers -ContentType 'application/json'