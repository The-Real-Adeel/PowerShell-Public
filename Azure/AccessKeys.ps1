<#

Script: ScanStorageAccountsForAccessKeys.ps1
Author: Adeel Anwar
Date: August 2nd, 2022
Duration: 10 minutes (est)

Description: 
    Scan the entire tenant (every subscription & every storage account) for services accessing the storage accounts via access keys.
    It will output the results into a CSV file.

#>

Connect-AzAccount

[System.Collections.ArrayList]$accessKeysUsage = New-Object -TypeName System.Collections.ArrayList # Create a collection called AccessKeysScan that will hold the array of data
$subscriptionlists = Get-AzSubscription # Scan every subscription. This variable will be used to list every single one
$Last90Days = (get-date).AddDays(-90) # Default for cmd for Az-Log is 7 days unless you add a start time. Store 90 days ago as a variable

# Loop through each subscription
foreach ($subscriptionlist in $subscriptionlists) {
    $context = Set-AzContext $subscriptionlist # Set context to the current subscription in the loop to scan
    $storageAccounts = Get-AzResource -ResourceType 'Microsoft.Storage/storageAccounts' # Scan every storage accounts. This variable will be used to list every single one
    # Loop through each storage account in storage accounts
    foreach ($storageAccount in $storageAccounts) {
        # run Az-log to see if storage account was accessed via Access Keys in the last 90 days
        $accessKeyCheck = Get-AzLog -ResourceId $storageAccount.id -StartTime $Last90Days | Where-Object {$_.authorization.action -eq "Microsoft.Storage/storageAccounts/listKeys/action"}
        #Store the latest one that tried to successfully access as a variable that is not an email (false positive)
        $LatestAccessAttempt = $accessKeyCheck | where-object {$_.caller -notlike "*central1.com*"} | Select-Object -First 1 

        # use if/else to store data in array objects depending on whether something tried to access using Access Keys
        if ($null -ne $LatestAccessAttempt){ #if az-Log is not null
            $StorageAccountDetails = [ordered]@{ #Store data in the following, in this order as variables
                SubscriptionName = $context.Subscription.Name
                StorageAccountName = $storageAccount.Name
                ResourceGroupName = $storageAccount.ResourceGroupName
                AccessUsingKeys = "True"
                CallerID = $LatestAccessAttempt.Caller # The Object ID that is calling the request for the key
                CallerDisplayName = (Get-AzADServicePrincipal -ObjectId $LatestAccessAttempt.Caller).displayName # Using the caller, it will get the display name
                LastAttemptDate = $LatestAccessAttempt.EventTimestamp.DateTime
            }
            $accessKeysUsage.add((New-Object psobject -Property $StorageAccountDetails)) | Out-Null
        }else{ #else if az-log is null still record the entries
            $StorageAccountDetails = [ordered]@{ #Store data in the following, in this order as variables
                SubscriptionName = $context.Subscription.Name
                StorageAccountName = $storageAccount.Name
                ResourceGroupName = $storageAccount.ResourceGroupName 
                AccessUsingKeys = "False"
                CallerID = "" # No calls, leave blank
                CallerDisplayName = "" # No calls, leave blank
                LastAttemptDate = "" # No calls, leave blank
            }
            $accessKeysUsage.add((New-Object psobject -Property $StorageAccountDetails)) | Out-Null # Add the data to the collection
        }
    }
}
$accessKeysUsage | Export-Csv -Path c:\test\test.csv -NoTypeInformation # take the collected data and place it in a csv file