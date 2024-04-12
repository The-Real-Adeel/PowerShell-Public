<#
Script: KQL-AzureResourceGraph-OrphanedResources.ps1
Author: Adeel
Date: April 1st, 2024
Description: 
	- Build a report of all orphaned resources by using KQL
    - Orphaned resources are resources that are not being leveraged and possibility costing money
    - We will use functions from Functions.ps1 to send email as well as a function to build HTML tables without filling this script up with HTML code
#>
. .\Functions.ps1

#---------------------
# Inital Variables Set
#---------------------

# Generate Intro Opening HTML (before first table)
$IntroOpeningHTML = '
This monthly report detects all orphaned resources. As in resources in our Tenant that are not being used in some way or another. <br>
- The first table will provide an overview: Containing the <b>Type</b>, the <b>amount</b> of resources that match and the <b>description</b> of why its marked as orphaned. <br>
- The tables that follow will only be for the resources it found (0 = no table). Each will list out all the resources & their details it detected as "orphaned".<br>
- Anything with IGNORE in the description is disabled, its tally would still be collected but it wont be reported as table. <br>
<br>
'
# Generate Intro Closing HTML (after first table)
$IntroClosingHTML = '<br>If you have any questions, please reach out to <ENTER EMAIL HERE>.<br>
'


# Stored to create the intro html for each table + an overview table using this data + query count data collected
$resourcesObject = [PSCustomObject]@{
    appServicePlan = @{Name="appServicePlan"; Description="App Service Plans detected without any hosting Apps."}
    vmPowerState = @{Name="vmPowerState"; Description="Azure VM PowerState detected as not currently running."}
    availabilitySet = @{Name="availabilitySet"; Description="Availability Sets that are not associated to any Virtual Machine (VM) or Virtual Machine Scale Set (VMSS)."}
    managedDisk = @{Name="managedDisk"; Description="Managed Disks with 'Unattached' state and not related to Azure Site Recovery."}
    publicIP = @{Name="publicIP"; Description="Public IPs that are not attached to any resource (VM, NAT Gateway, Load Balancer, Application Gateway, Public IP Prefix, etc.)."}
    networkInterface = @{Name="networkInterface"; Description="Network Interfaces that are not attached to any resource."}
    networkSecurityGroup = @{Name="networkSecurityGroup"; Description="Network Security Groups (NSG) with no Network Interfaces nor subnets connected."}
    routeTable = @{Name="routeTable"; Description="Route Tables that not attached to any subnet."}
    loadBalancer = @{Name="loadBalancer"; Description="Load Balancers with empty backend address pools."}
    frontDoorWAFPolicy = @{Name="frontDoorWAFPolicy"; Description="Front Door WAF Policy without associations. (Frontend Endpoint Links, Security Policy Links)."}
    trafficManager = @{Name="trafficManager"; Description="Traffic Manager Profiles with no endpoints."}
    applicationGateway = @{Name="applicationGateway"; Description="Application Gateways without backend targets (in backend pools)."}
    virtualNetwork = @{Name="virtualNetwork"; Description="IGNORED: Virtual Networks (VNETs) without any subnets."}
    subnet = @{Name="subnet"; Description="IGNORED: Subnets without Connected Devices or Delegation (Empty Subnets)."}
    natGateway = @{Name="natGateway"; Description="NAT Gateways that not attached to any subnet."}
    azureFirewallIPGroup = @{Name="azureFirewallIPGroup"; Description="IP Groups that are not attached to any Azure Firewall."}
    privateDNSZones = @{Name="privateDNSZones"; Description="Private DNS zones without Virtual Network Links."}
    privateEndpoints = @{Name="privateEndpoints"; Description="Private Endpoints that are not connected to any resource."}
    resourceGroup = @{Name="resourceGroup"; Description="Resource Groups without resources (including hidden types resources)."}
    apiConnections = @{Name="apiConnections"; Description="API Connections that not related to any Logic App."}
    expiredCertificates = @{Name="expiredCertificates"; Description="Expired certificates discovered in Azure."}
}
$IntroColumns = "Type", "Detected", "Description"
$IntroRows = New-Object System.Collections.Specialized.OrderedDictionary
$IntroRowCount = 1
$query = @{} # init hash table for queries
$HTML = ""

#----------
# Functions
#----------
function GenerateOpenHTML {
    param (
        [Parameter(Mandatory=$True)]
        [String]$service,
        
        [Parameter(Mandatory=$True)]
        [PSObject]$resourcesObject
    )
    # Using a script block to dynamically access property
    $name = & { $resourcesObject.$service.Name}
    $description = & { $resourcesObject.$service.Description }
    $value = @"
<br>
<hr style="height:3px;border-width:0;color:#FF0000;background-color:#FF0000">
<br>
<p style="font-size:20px; "> <b>$($name):</b> $($description) </p>
<br>
"@
    return $value
}

#-------------
# Collect Data 
#-------------
# From all subscriptions for all the different kinds of resources. Generating unique html tables for each as they all contain different fields

# App Service plans without hosting Apps.
$query.appServicePlan = Search-AzGraph -Query @'
resources
    | where type =~ "microsoft.web/serverfarms"
    | where properties.numberOfSites == 0
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, Sku=sku.name, Tier=sku.tier, id=id, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.appServicePlan)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "SKU", "Tier", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "appServicePlan" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.appServicePlan) {

        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.sku)", "$($item.tier)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML = New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Azure VM PowerState detected as not currently running.
$query.vmPowerState = Search-AzGraph -Query @'
resources
    | where type == "microsoft.compute/virtualmachines"
    | extend details = pack_all()
    | extend powerstate = replace("PowerState/", "", tostring(properties.extended.instanceView.powerState.code)) //Replace value "PowerState/" with nothing in property that we coverted to string
    | where powerstate in ("stopped", "deallocated", "stopping", "deallocating")
    | where isnull(tags['Managed PowerState']) or tolower(tostring(tags['Managed PowerState'])) != "true"
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, powerstate, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.vmPowerState)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "PowerState", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "vmPowerState" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.vmPowerState) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.PowerState)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Availability Sets that not associated to any Virtual Machine (VM) or Virtual Machine Scale Set (VMSS).
$query.availabilitySet = Search-AzGraph -Query @'
resources
    | where type =~ "Microsoft.Compute/availabilitySets"
    | where properties.virtualMachines == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.availabilitySet)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "availabilitySet" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.availabilitySet) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Managed Disks with 'Unattached' state and not related to Azure Site Recovery.
# Note: Azure Site Recovery (aka: ASR) managed disks are excluded from the orphaned resource query.
$query.managedDisk = Search-AzGraph -Query @'
resources
    | where type has "microsoft.compute/disks"
    | extend diskState = tostring(properties.diskState)
    | where managedBy == ""
    | where not(name endswith "-ASRReplica" or name startswith "ms-asr-" or name startswith "asrseeddisk-")
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, diskState, sku=sku.name, size=properties.diskSizeGB, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.managedDisk)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "DiskState", "SKU", "DiskSize(GB)", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "managedDisk" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.managedDisk) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.DiskState)", "$($item.sku)", "$($item.size)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Public IPs that are not attached to any resource (VM, NAT Gateway, Load Balancer, Application Gateway, Public IP Prefix, etc.).
$query.publicIP = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/publicipaddresses"
    | where properties.ipConfiguration == "" and properties.natGateway == "" and properties.publicIPPrefix == ""
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, sku=sku.name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.publicIP)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "SKU", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "publicIP" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.publicIP) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.sku)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}


# Network Interfaces that are not attached to any resource.
$query.networkInterface = Search-AzGraph -Query @'
resources
    | where type has "microsoft.network/networkinterfaces"
    | where isnull(properties.privateEndpoint)
    | where isnull(properties.privateLinkService)
    | where properties.hostedWorkloads == "[]"
    | where properties !has "virtualmachine"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, tags, id=id, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.networkInterface)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "networkInterface" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.networkInterface) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}


# Network Security Groups
$query.networkSecurityGroup = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/networksecuritygroups" and isnull(properties.networkInterfaces) and isnull(properties.subnets)
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.networkSecurityGroup)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "networkSecurityGroup" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.networkSecurityGroup) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Route Tables that not attached to any subnet.
$query.routeTable = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/routetables"
    | where isnull(properties.subnets)
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.routeTable)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "routeTable" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.routeTable) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Load Balancers with empty backend address pools.
$query.loadBalancer = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/loadbalancers"
    | where properties.backendAddressPools == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.loadBalancer)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "loadBalancer" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.loadBalancer) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Front Door WAF Policy without associations. (Frontend Endpoint Links, Security Policy Links)
$query.frontDoorWAFPolicy = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/frontdoorwebapplicationfirewallpolicies"
    | where properties.frontendEndpointLinks== "[]" and properties.securityPolicyLinks == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, sku=sku.name, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.frontDoorWAFPolicy)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "SKU", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "frontDoorWAFPolicy" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.frontDoorWAFPolicy) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.sku)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Traffic Manager Profiles
$query.trafficManager = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/trafficmanagerprofiles"
    | where properties.endpoints == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, sku=sku.name, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.trafficManager)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "SKU", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "trafficManager" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.trafficManager) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.sku)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Application Gateways without backend targets. (in backend pools)
$query.applicationGateway = Search-AzGraph -Query @'
resources
    | where type =~ "microsoft.network/applicationgateways"
    | extend backendPoolsCount = array_length(properties.backendAddressPools),SKUName= tostring(properties.sku.name), SKUTier= tostring(properties.sku.tier),SKUCapacity=properties.sku.capacity,backendPools=properties.backendAddressPools , AppGwId = tostring(id)
    | project AppGwId, resourceGroup, location, subscriptionId, tags, name, SKUName, SKUTier, SKUCapacity
    | join (
        resources
        | where type =~ "microsoft.network/applicationgateways"
        | mvexpand backendPools = properties.backendAddressPools
        | extend backendIPCount = array_length(backendPools.properties.backendIPConfigurations)
        | extend backendAddressesCount = array_length(backendPools.properties.backendAddresses)
        | extend backendPoolName  = backendPools.properties.backendAddressPools.name
        | extend AppGwId = tostring(id)
        | summarize backendIPCount = sum(backendIPCount) ,backendAddressesCount=sum(backendAddressesCount) by AppGwId
    ) on AppGwId
    | project-away AppGwId1
    | where  (backendIPCount == 0 or isempty(backendIPCount)) and (backendAddressesCount==0 or isempty(backendAddressesCount))
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, SKUTier, SKUCapacity, id=AppGwId, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.applicationGateway)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "SKU Tier", "SKU Capacity" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "applicationGateway" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.applicationGateway) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.SKUTier)", "$($item.SKUCapacity)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Virtual Networks (VNETs) without subnets.
$query.virtualNetwork = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/virtualnetworks"
    | where properties.subnets == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1


if ($null -ne $($query.virtualNetwork)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "virtualNetwork" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.virtualNetwork) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Subnets without Connected Devices or Delegation. (Empty Subnets)
$query.subnet = Search-AzGraph -Query @'
resources
    | where type =~ "microsoft.network/virtualnetworks"
    | extend subnet = properties.subnets
    | mv-expand subnet
    | extend ipConfigurations = subnet.properties.ipConfigurations
    | extend delegations = subnet.properties.delegations
    | where isnull(ipConfigurations) and delegations == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, resource=tostring(subnet.name), id=id, VNetName=name, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1


if ($null -ne $($query.subnet)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Virtual Network","Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "subnet" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.subnet) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)/subnets!!TITLE!!$($item.Resource)", $($item.VNetName) ,"$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}


# NAT Gateways that not attached to any subnet.
$query.natGateway = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/natgateways"
    | where isnull(properties.subnets)
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, sku=tostring(sku.name), tier=tostring(sku.tier), id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.natGateway)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "SKU", "Tier" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "natGateway" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.natGateway) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.sku)", "$($item.tier)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# IP Groups that not attached to any Azure Firewall.
$query.azureFirewallIPGroup = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/ipgroups"
    | where properties.firewalls == "[]" and properties.firewallPolicies == "[]"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.azureFirewallIPGroup)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "azureFirewallIPGroup" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.azureFirewallIPGroup) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}


# Private DNS zones without Virtual Network Links.
$query.privateDNSZones = Search-AzGraph -Query @'
resources
    | where type == "microsoft.network/privatednszones"
    | where properties.numberOfVirtualNetworkLinks == 0
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, NumberOfRecordSets=properties.numberOfRecordSets, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.privateDNSZones)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "NumberOfRecordSets" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "privateDNSZones" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.privateDNSZones) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)","$($item.NumberOfRecordSets)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Private Endpoints that are not connected to any resource.
$query.privateEndpoints = Search-AzGraph -Query @'
resources
    | where type =~ "microsoft.network/privateendpoints"
    | extend connection = iff(array_length(properties.manualPrivateLinkServiceConnections) > 0, properties.manualPrivateLinkServiceConnections[0], properties.privateLinkServiceConnections[0])
    | extend subnetId = properties.subnet.id
    | extend subnetIdSplit = split(subnetId, "/")
    | extend vnetId = strcat_array(array_slice(subnetIdSplit,0,8), "/")
    | extend vnetIdSplit = split(vnetId, "/")
    | extend serviceId = tostring(connection.properties.privateLinkServiceId)
    | extend serviceIdSplit = split(serviceId, "/")
    | extend serviceName = tostring(serviceIdSplit[8])
    | extend serviceTypeEnum = iff(isnotnull(serviceIdSplit[6]), tolower(strcat(serviceIdSplit[6], "/", serviceIdSplit[7])), "microsoft.network/privatelinkservices")
    | extend stateEnum = tostring(connection.properties.privateLinkServiceConnectionState.status)
    | extend groupIds = tostring(connection.properties.groupIds[0])
    | where stateEnum == "Disconnected"
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=name, serviceName, serviceTypeEnum, groupIds, vnetId, subnetId, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.privateEndpoints)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource", "Service Name", "Service Type", "GroupIds", "VNet", "Subnet", "Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "privateEndpoints" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.privateEndpoints) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)","$($item.serviceName)", "$($item.serviceTypeEnum)", "$($item.groupIds)", "$($item.details.vnetidsplit | Select-Object -Last 1)", "$($item.details.subnetidsplit | Select-Object -Last 1)",  "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}


# Resource Groups without resources (including hidden types resources).
$query.resourceGroup = Search-AzGraph -Query @'
ResourceContainers
    | where type == "microsoft.resources/subscriptions/resourcegroups"
    | extend rgAndSub = strcat(resourceGroup, "--", subscriptionId)
    | join kind=leftouter (
        Resources
        | extend rgAndSub = strcat(resourceGroup, "--", subscriptionId)
        | summarize count() by rgAndSub
    ) on rgAndSub
    | where isnull(count_)
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | project subscription, ResourceGroup=name, location, id=id, tags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.resourceGroup)){
    $Columns = "Subscription", "ResourceGroup", "Location" ,"Tag:Creator", "Tag:PrimaryOwner", "Tag:SecondaryOnwer", "Tag:CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "resourceGroup" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.resourceGroup) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.resourceGroup)", "$($item.Location)", "$($item.tags."Creator")", "$($item.tags."Owner Primary")", "$($item.tags."Owner Secondary")", "$($item.tags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}


# API Connections that not related to any Logic App.
$query.apiConnections = Search-AzGraph -Query @'
resources
    | where type =~ "Microsoft.Web/connections"
    | project resourceId = id , apiName = name, subscriptionId, resourceGroup, tags, location
    | join kind = leftouter (
        resources
        | where type == "microsoft.logic/workflows"
        | extend resourceGroup, location, subscriptionId, properties
        | extend var_json = properties["parameters"]["$connections"]["value"]
        | mvexpand var_connection = var_json
        | where notnull(var_connection)
        | extend connectionId = extract("connectionId\":\"(.*?)\"", 1, tostring(var_connection))
        | project connectionId, name
        )
        on $left.resourceId == $right.connectionId
    | where connectionId == ""
    | extend details = pack_all()
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions"
        | project subscriptionId, subscription=name
    ) on subscriptionId
    | join kind=leftouter (
        resourcecontainers
        | where type == "microsoft.resources/subscriptions/resourcegroups"
        | project resourceGroup, rgtags=tags
    ) on resourceGroup
    | project subscription, resourceGroup, location, Resource=apiName, id=resourceId, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.apiConnections)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "apiConnections" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.apiConnections) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Blue"
}

# Expired Certificates
$query.expiredCertificates = Search-AzGraph -Query @'
resources
| where type == "microsoft.web/certificates"
| extend expiresOn = todatetime(properties.expirationDate)
| where expiresOn <= now()
| extend details = pack_all()
| join kind=leftouter (
    resourcecontainers
    | where type == "microsoft.resources/subscriptions"
    | project subscriptionId, subscription=name
) on subscriptionId
| join kind=leftouter (
    resourcecontainers
    | where type == "microsoft.resources/subscriptions/resourcegroups"
    | project resourceGroup, rgtags=tags
) on resourceGroup
| project subscription, resourceGroup, location, Resource=name, id=id, tags, rgtags, details
'@ -First 1000

Start-Sleep -Seconds 1

if ($null -ne $($query.expiredCertificates)){
    $Columns = "Subscription", "ResourceGroup", "Location", "Resource" ,"Tag:Creator", "Tag:RG-PrimaryOwner", "Tag:RG-SecondaryOnwer", "Tag:RG-CostCentre"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary
    $RowCount = 1
    $OpeningHTML = generateOpenHTML -service "expiredCertificates" -resourcesObject $resourcesObject
    $ClosingHTML = ''
    foreach ($item in $query.expiredCertificates) {
        $RowCount += 1
        $Row = "$($item.Subscription)", "$($item.ResourceGroup)", "$($item.Location)", "!!LINK!!<AzureDomain>/resource$($item.id)!!TITLE!!$($item.Resource)", "$($item.tags."Creator")", "$($item.rgtags."Owner Primary")", "$($item.rgtags."Owner Secondary")", "$($item.rgtags."Cost Centre")"
        $Rows += @{$RowCount = $Row}
    }
    $HTML += New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Green"
}

#---------------
# Generate Email
#---------------

# Finalize data for first table and generate it. Add it to the front of the HTML
foreach($resource in $resourcesObject.psobject.Properties){
    $IntroRowCount += 1
    $tally = $query.$($resource.value.name).count
    $IntroRow = "$($resource.value.name)", "<b>$($tally)</b>", "$($resource.value.Description)"
    $IntroRows += @{$IntroRowCount = $IntroRow}
}
$IntroHTML = New-HTMLTable -Rows $IntroRows -Columns $IntroColumns -OpeningHTML $IntroOpeningHTML -ClosingHTML $IntroClosingHTML -Theme "Blue"

# Combine the Intro HTML with all the tables collected to send as body of the email\
$HTML = $IntroHTML + $HTML

## Test html locally
# $HTML | Out-File "C:\test\orphanedResources.html" 

Send-EmailFunction -Subject "Orphaned Resources"  -To "<email>" -Body $HTML -HTML $True -Importance $True
