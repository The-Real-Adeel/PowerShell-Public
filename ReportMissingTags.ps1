<#

Script: report_azure_missingtags.ps1
Author: Adeel Anwar

Description: 
    Scan all subscription searching the value stored in $key
    If its missing report it in an excel sheet. Place all resources affected

#>

[System.Collections.ArrayList]$FindOwnerTags = New-Object -TypeName System.Collections.ArrayList

$SubscriptionLists = Get-AzSubscription
[string]$key = '<InsertTagKey' 

foreach ($subscriptionlist in $subscriptionlists) {
    $context = Set-AzContext $subscriptionlist
    $ResourceTags = Get-AzResource
    foreach ($Resourcetag in $Resourcetags) {
        if ($null -eq ($Resourcetag | Where-object {$_.Tags.keys -eq $Key})){ # Does Not Contain Primary Tag
            $TagDetails = [ordered]@{ #Store data in the following, in this order as variables
                SubscriptionName = $context.Subscription.Name
                Resource = $Resourcetag.name
                ResourceGroupName = $Resourcetag.ResourceGroupName
                Tags = (($resourcetag | Select-Object tags | out-string -stream).TrimStart("{, Tags, ----, ") | out-string).trim() # Converts hashtable to a single string line
            }
        }
        $FindOwnerTags.add((New-Object psobject -Property $TagDetails)) | Out-Null
    }
}
# $FindOwnerTags | Format-table -AutoSize
$FileName = '.\MissingTags_' + ($key -replace '[:]','') + '.csv'
$FindOwnerTags | Export-Csv -Path $FileName -NoTypeInformation

