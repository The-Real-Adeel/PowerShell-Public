<#

Script: ReplaceTags.ps1
Author: Adeel Anwar

Description: 
    Scan a subscription for tags, replace them with the new tag and report the results. 
    If the status is blank, it succeeded. If the status has an error, it failed

#>

# Connect-AzAccount
# note you may recieve "Operation returned an invalid status code 'Accepted'" but it works. Issue has been rasised with MS: https://github.com/Azure/azure-powershell/issues/19120

$TagKey = "Cost-Centre"
$OldTagValue = "3334" 
$NewTagValue = "3335"
$subscription = "Contoso"

[System.Collections.ArrayList]$FindCostCentreTags = New-Object -TypeName System.Collections.ArrayList # store data in collection

$context = Set-AzContext $Subscription

$ResourceTags = Get-AzResource -Tag @{$tagKey=$OldTagValue}
    foreach ($Resourcetag in $Resourcetags) {
        $results = $null
        # Use ErrorVariable to store error data in a variable. These kinds of errors are not caught by try/catch. Note the data will be an array
        Update-AzTag -ResourceId $ResourceTag.id -Tag @{$tagKey=$NewTagValue} -Operation Merge -ErrorVariable results -ErrorAction SilentlyContinue  | Out-Null
        if ($null -eq $results) {
            $results = "Successfully changed"
        }
        $ResourceWithTagDetails = [ordered]@{ #Store data in the following, in this order as variables
            SubscriptionName = $context.Subscription.Name
            Resource = $Resourcetag.name
            ResourceGroupName = $Resourcetag.ResourceGroupName
            PreviousTag = "$($tagkey) : $($OldTagValue)"
            NewTag = "$($tagkey) : $($NewTagValue)"
            Status = $results -join ',' # Store array in to a single string
        }
        $FindCostCentreTags.add((New-Object psobject -Property $ResourceWithTagDetails)) | Out-Null # extract data and put it in the collection
    }

$ResourceGroupTags = Get-AzResourceGroup -Tag @{$tagKey=$OldTagValue}

foreach ($ResourceGrouptag in $ResourceGrouptags) {
    $results = $null
    Update-AzTag -ResourceId $ResourceGrouptag.resourceid -Tag @{$tagKey=$NewTagValue} -Operation Merge -ErrorVariable results -ErrorAction SilentlyContinue  | Out-Null
    if ($null -eq $results) {
        $results = "Successfully changed"
    }
    $ResourceWithTagDetails = [ordered]@{ #Store data in the following, in this order as variables
        SubscriptionName = $context.Subscription.Name
        Resource = "NULL (Changing the Resource Group's Tag)"
        ResourceGroupName = $ResourceGrouptag.ResourceGroupName
        PreviousTag = "$($tagkey) : $($OldTagValue)"
        NewTag = "$($tagkey) : $($NewTagValue)"
        Status = $results -join ',' #Store array in to a single line
    }

    $FindCostCentreTags.add((New-Object psobject -Property $ResourceWithTagDetails)) | Out-Null # extract data and put it in the collection
}
$FindCostCentreTags | Format-table -AutoSize

$FindCostCentreTags | Export-Csv -Path .\tagtest.csv -NoTypeInformation # take the collected data and place it in a csv file