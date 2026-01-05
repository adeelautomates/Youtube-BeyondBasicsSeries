<#
    NAME: manual_azure_subscriptionPostDeployConfig
    CREATOR: Adeel A.
    DESCRIPTION: 
        - Tasks to do after a subscription has been created:
          - Build an Event Grid system to capture when resources in the tenant are created
          - <toAdd>
#>

#----------------
# Parameters
#----------------

param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$subscriptionName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$region
)

#----------------------
# Authenticate to Azure
#----------------------
try {
    if (-not (Get-AzContext)) {
        # check if there is already an account signed in, If Not... 
        connect-azaccount -identity | out-null # Connect as Self
        write-output "Connected as the Managed Identity of the Automation Account."
    } # Otherwise use the signed in account
    else {
        Write-Output "Using existing Azure Connection: $((Get-AzContext).Account.Id)"
    }
}
catch {
    write-error "Failed to authenticate: $($_.Exception.Message)"
    throw
}
Write-Output "-------------------------------------------"

#----------
# Variables
#----------
$retrieveToken = Get-AzAccessToken -ResourceUrl "https://management.azure.com"
$token = $retrieveToken.Token -is [securestring] ? ($retrieveToken.token | ConvertFrom-SecureString -AsPlainText) : $retrieveToken.Token
$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}
$region = $region.trim().ToLower().Replace(" ", "")
$acceptedRegions = @{
    canadacentral = "cc"
    canadaeast = "ce"
}

#---------
# Validate
#---------
write-output "Running Validation...."
try {
    $uri = "https://management.azure.com/subscriptions?api-version=2020-01-01"
    $subscription = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).value | Where-Object displayName -eq $subscriptionName
    if(-not $subscription){
        write-error "Subscription '$subscriptionName' was not found" -ErrorAction Stop
    }
    if(-not $acceptedRegions.ContainsKey($region)){
        write-error "Region '$region' is not supported. Allowed $($acceptedRegions.keys -join ", ")" -ErrorAction Stop
    }
    $regionShort = $acceptedRegions[$region]
    $envShort    = $subscription.tags.'env_short'
    $uriPrefix   = "https://management.azure.com/subscriptions/$($subscription.subscriptionId)/"
    Write-Output "- Parameters configured correctly. Proceeding"

}
catch {
    write-error "Could not validate parameter(s). Failing script: $($_.Exception.Message)"
    throw
}

#-----------------------------------------
# Configure Management RG for Subscription
#-----------------------------------------
write-output "Configuring Resource Group (If needed)"
$tags = @{
    "cost-centre" = "1001"
    env_short = $envShort
    owner1 = "owner1@example.com"
    owner2 = "owner2@example.com"
}
$rgName = "rg-$($regionShort)-$($envShort)-management-001"
$uri = "$uriPrefix/resourceGroups/$($rgName)?api-version=2021-04-01"

try {
    $rg = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -erroraction Stop
    write-output "- Resource Group already exists, Using: $($rg.name)"
}
catch {
    $body = @{
        location = $region
        tags = $tags
    } | ConvertTo-Json -Depth 20
    try {
        $rg = Invoke-RestMethod -Method PUT -body $body -Uri $uri -Headers $headers -erroraction Stop
        write-output "- Resource Group did not exist, Created: $rgName"
    }
    catch {
        $errorStatus = $_
        write-error "Failed to Create Resource Group: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | ConvertFrom-Json).error.Message)"
        throw
    }
}

#-------------------------------------
# Register Providers for Subscriptions
#-------------------------------------
$providers = @(
    "Microsoft.EventGrid"
    "Microsoft.EventHub"
)
write-output "Registering Providers for Subscriptions (if needed)"
foreach($provider in $providers){
    $uri = "$uriPrefix/providers/$($provider)?api-version=2021-04-01"
    if((Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers).RegistrationState -eq "Registered"){
        "- Already Registered $provider"
        continue
    }
    $uri = "$uriPrefix/providers/$($provider)/register?api-version=2021-04-01"
    Invoke-RestMethod -Method POST -Uri $Uri -Headers $headers | out-null
    write-output "- Registered Provider: $provider"
}
start-sleep 10


#--------------------------
# Configure EventGrid Topic
#--------------------------
Write-Output "Configuring EventGrid Topic (If needed)."
$eventTopicName = "example-eg-topic-subscription-$($envShort)"
$uri = "$uriPrefix/resourceGroups/$($rg.name)/providers/Microsoft.EventGrid/systemTopics/$($eventTopicName)?api-version=2022-06-15"
try{
    $eventTopic = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop
    write-output "- EventGridTopic already exists, Using: $eventTopicName..."
}
catch{
    $body = @{
        location = "global" # global because subscriptions are not tied to a region, neither are their event grids
        properties = @{
            source = $subscription.id #resourceID
            TopicType = "Microsoft.Resources.Subscriptions"
        }
        tags = $tags
    } | convertTo-Json -Depth 20
    try {
        $eventTopic = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $body -ErrorAction Stop
        write-output "- Successfully Created EventGridTopic: $($eventTopic.Name)"
    }
    catch {
        $errorStatus = $_
        write-error "Failed to create EventGrid Topic: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | ConvertFrom-Json).error.Message)"
        throw
    }
}

#---------------------------------
# Configure EventGrid Subscription
#---------------------------------
Write-Output "Configuring EventGrid Subscription"
$eventSubName     = "new-resourceTagging"
$functionName     = "eventGrid-resourceOwnerTimeTagger"
$funcResourceId   = "/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-cc-mgmt-funcapp-001/providers/Microsoft.Web/sites/example-cc-mgmt-funcapp-001/functions/$functionName"
$uri              = "$uriPrefix/resourceGroups/$($rg.name)/providers/Microsoft.EventGrid/systemTopics/$($eventTopic.name)/eventSubscriptions/$($eventSubName)?api-version=2022-06-15"
$advFilters = @(
    @{
        operatorType    = "StringNotContains"
        key             = "data.operationName"
        values          = @(
            "Microsoft.Authorization/",
            "Microsoft.Resources/",
            "Microsoft.Insights/",
            "Microsoft.OperationalInsights/",
            "Microsoft.OperationsManagement/",
            "Microsoft.RecoveryServices/",
            "Microsoft.Support/",
            "Microsoft.Blueprint/",
            "Microsoft.Compute/snapshots/write",
            "Microsoft.Compute/virtualMachines/runCommands/write",
            "Microsoft.Network/privateDnsZones/A/write",
            "Microsoft.EventGrid/systemTopics/eventSubscriptions/write",
            "Microsoft.ServiceBus/namespaces/queues/write",
            "Microsoft.Compute/virtualMachines/extensions/write",
            "Microsoft.HybridCompute/machines/extensions/write",
            "Microsoft.Compute/virtualMachines/redeploy/action",
            "Microsoft.Compute/virtualMachines/performMaintenance/action"
        )
    }
)
$body = @{
    properties = @{
        destination = @{
            endpointType = "AzureFunction"
            properties   = @{
                resourceId = $funcResourceId
            }
        }
        filter = @{
            includedEventTypes = @(
                "Microsoft.Resources.ResourceWriteSuccess"
            )
            advancedFilters    = $advFilters
        }
    }
} | ConvertTo-Json -Depth 20

try {
    $eventSub = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $body -ErrorAction Stop
    Write-Output "- Success Event Grid Subscription: $($eventSub.name) with Destination $($functionName)"
}
catch {
    $errorStatus = $_
    write-error "Failed to Create/Update Event Subscription: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
    throw
}

#-----------------------------------------
# Configure Log Analytics for Subscription
#-----------------------------------------
Write-Output "Configuring Log Analytics"
$diagSettingName = "send-activitylog-to-logAnalytics"
$workspaceResourceId = "/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/rg-cc-mgmt-monitor-001/providers/Microsoft.OperationalInsights/workspaces/example-cc-loganalytics-all-logs-001"
$uri = "$uriPrefix/providers/Microsoft.Insights/diagnosticSettings/$($diagSettingName)?api-version=2021-05-01-preview"
$body = @{
    properties = @{
        workspaceId = $workspaceResourceId
        logs = @(
            @{ category = "Administrative"; enabled = $true}
            # @{ category = "Policy"; enabled = $true } # if you wanted to add more do it like this
        )
    }
} | ConvertTo-Json -Depth 20
try {
    Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $body -ErrorAction Stop | out-null
    Write-Output "- Success Diagnostic Setting '$diagSettingName' configured: ActivityLog -> Log Analytics"   
}
catch {
    $errorStatus = $_
    write-error "Failed to Create/Update Log Analytics: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
    throw
}

#----------------------
# Default Groups & RBAC
#---------------------- 
Write-Output "Configuring Groups and RBAC for them"
$RetrieveToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/"
$mgtoken = $RetrieveToken.Token -is [securestring] ? ($RetrieveToken.Token | ConvertFrom-SecureString -AsPlainText) : $RetrieveToken.Token
$mgheaders = @{
    Authorization = "Bearer $mgtoken"
    "Content-Type" = "application/json"
}
$groupUri = "https://graph.microsoft.com/v1.0/groups"
$entraGroups = @(
    @{
        groupName = "az-sg-sub-$($envShort)-reader"
        role      = "Reader"
    }
    @{
        groupName = "az-sg-sub-$($envShort)-contributor"
        role      = "Contributor"
    }
)
$uri = "$uriPrefix/providers/Microsoft.Authorization/roleDefinitions?api-version=2022-04-01"
$roleDefinitions = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop).value


foreach($entraGroup in $entraGroups){
    $uri = $groupUri + "?`$filter=displayName eq '$($entraGroup.groupName)'"
    $TargetGroup = (Invoke-RestMethod -Method GET -Uri $uri -Headers $mgheaders).value
    if(-not $TargetGroup) {
        $body = @{
            displayName  = $entraGroup.groupName
            mailEnabled = $false
            mailNickname = $entraGroup.groupName.ToLower()
            securityEnabled = $true
            description  = "Subscription access group for $subscriptionName with ($($entraGroup.role)) access"
        } | ConvertTo-Json -Depth 10
        try {
            $TargetGroup = Invoke-RestMethod -Method POST -Uri $groupUri -Headers $mgheaders -Body $body -errorAction Stop
            Write-Output "- Created group: $($TargetGroup.displayName) [$($TargetGroup.id)]"
        }
        catch {
            $errorStatus = $_
            Write-Error "Failed to create group '$($entraGroup.groupName)': [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
            throw
        }
    }
        else{
        Write-Output "- Found Group: $($TargetGroup.displayName) [$($TargetGroup.id)] "
    }
    start-sleep 30
    try{
        # Get Role Definition for Group
        $roleDefinition = $roleDefinitions | where-object {$_.properties.roleName -eq $entraGroup.role}

        # Generate Deterministic GUID (so its always the same for every unique group)
        $seed = "$($subscription.subscriptionId)|$($TargetGroup.id)|$($entraGroup.role)"
        $md5 = [System.Security.Cryptography.MD5CryptoServiceProvider]::Create()
        $inputBytes = [System.Text.Encoding]::Default.GetBytes($seed)
        $hashBytes = $md5.ComputeHash($inputBytes)
        $roleAssignmentId = [Guid]::New($hashBytes)

        # Assigning Role Permission
        $uri = "$uriPrefix/providers/Microsoft.Authorization/roleAssignments/$($roleAssignmentId)?api-version=2022-04-01"
        $body = @{
            properties = @{
                roleDefinitionId = $roleDefinition.id
                principalId = $TargetGroup.id
                principalType = "Group"
            }
        } | ConvertTo-Json
        $roleAssign = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $body -ErrorAction Stop
        Write-Output "- Success RBAC Perms for : $($TargetGroup.displayName) Granted "
    }
    catch {
        $errorStatus = $_
        Write-Error "Failed to assign role: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
        throw
    }
}

Write-Output "-------------------------------------------"
