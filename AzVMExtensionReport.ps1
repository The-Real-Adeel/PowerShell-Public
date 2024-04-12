# Define the path for the Excel file
$excelFilePath = "C:\test\extensionData.xlsx"

# Prepare to collect VM and extension data
[System.Collections.ArrayList]$extData = New-Object System.Collections.ArrayList

# Get Azure subscriptions
$subs = Get-AzSubscription

foreach ($sub in $subs) {
    $context = Set-AzContext -SubscriptionId $sub.Id
    $vms = Get-AzVM
    foreach ($vm in $vms) {
        $exts = Get-AzVMExtension -ResourceGroupName $vm.ResourceGroupName -VMName $vm.Name
        if($null -eq $exts){ #if they dont have it, add 
            $extDetails = [ordered]@{
                "Subscription" = $sub.Name
                "Resource Group" = $vm.ResourceGroupName
                "Virtual Machine" = $vm.Name
                "OS" = $vm.StorageProfile.OsDisk.OsType
                "Extension Name" = ""
                "Extension Publisher" = ""
                "Extension Type" = ""
                "State" = ""
                "Direct Link" = ""
            }
            $null = $extData.Add((New-Object PSObject -Property $extDetails))            
        }
        else{
            $link = "<AzDomain>/resource$($vm.id)/appsandextensions"
            foreach ($ext in $exts) {
                # Store data, including a hyperlink for the VM name
                $extDetails = [ordered]@{
                    "Subscription" = $sub.Name
                    "Resource Group" = $vm.ResourceGroupName
                    "Virtual Machine" = $vm.Name
                    "OS" = $vm.StorageProfile.OsDisk.OsType
                    "Extension Name" = $ext.Name
                    "Extension Publisher" = $ext.Publisher
                    "Extension Type" = $ext.ExtensionType
                    "State" = $ext.ProvisioningState
                    "Direct Link" = "=HYPERLINK(""$link"", """ + "Click Here" + """)"
                }
                $null = $extData.Add((New-Object PSObject -Property $extDetails))
            }
        }
    }
}

# Export the collected data to an Excel file
$extData | Export-Excel -Path $excelFilePath -AutoSize -TableName "VMExtensions" -Show

"Total Number of VMs = " + ($extData | group-object "virtual machine").count

"--------------"
"Windows Report"
"--------------"
"Total Windows VMs = " + ($extData | where-object { ($_.OS -eq "Windows") } | Group-Object "virtual machine").count
"Total Windows VMs with Extensions = " + ($extData | where-object { ($_.OS -eq "Windows") -and ($_."Extension Name" -ne "")} | Group-Object "virtual machine").count
"Total Windows VMs Without Extensions = " + ($extData | where-object { ($_.OS -eq "Windows") -and ($_."Extension Name" -eq "")} | Group-Object "virtual machine").count

""
"Listing all extensions applied to windows VMs and total count:"
$WExtensionList = $extData | where-object { ($_.OS -eq "Windows") -and ($_."Extension Name" -ne "")} | Group-Object "Extension Name" | select-object count, name | sort-object count -Descending 
foreach($item in $WExtensionList){
    Write-Output "- $($item.name) applied $($item.count) times"
}
""

"-------------"
"Linux Report"
"-------------"
"Total Linux VMs = " + ($extData | where-object { ($_.OS -eq "Linux") } | Group-Object "virtual machine").count
"Total Linux VMs with Extensions = " + ($extData | where-object { ($_.OS -eq "Linux") -and ($_."Extension Name" -ne "")} | Group-Object "virtual machine").count
"Total Linux VMs Without Extensions = " + ($extData | where-object { ($_.OS -eq "Linux") -and ($_."Extension Name" -eq "")} | Group-Object "virtual machine").count

""
"Listing all extensions applied to Linux VMs and total count:"
$WExtensionList = $extData | where-object { ($_.OS -eq "Linux") -and ($_."Extension Name" -ne "")} | Group-Object "Extension Name" | select-object count, name | sort-object count -Descending 
foreach($item in $WExtensionList){
    Write-Output "- $($item.name) applied $($item.count) times"
}
""