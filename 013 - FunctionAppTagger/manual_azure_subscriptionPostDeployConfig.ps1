<#
    NAME: manual_azure_subscriptionPostDeployConfig
    CREATOR: Adeel A.
    DESCRIPTION: 
        - Tasks to do after a subscription has been created:
          - Build an Event Grid system to capture when resources in the tenant are created
    NOTES:
        - LOOK FOR ANY <> to modify for your script
        - I will modify this in the future so stay tuned for a more refined version of it
#>


#----------------------
# Parameters
#----------------------
param (
    [Parameter(Mandatory = $true)]
        $subscriptionName,
    [Parameter(Mandatory = $true)]
        $region
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
#---------
# Validate
#---------
try {
    $subscription = get-azsubscription -SubscriptionName $subscriptionName -ErrorAction stop
    if (!(get-azlocation | Where-Object location -eq $region)){ 
        Write-Error "Invalid Azure Region: $Region" -ErrorAction stop 
    }
}
catch {
    write-error "Could not validate parameter(s). Failing script: $($_.Exception.Message)"
    throw
}
Write-Output "Parameters configured correctly... proceeding"
set-azcontext -Subscription $subscription.id | Out-Null

#-------------------------------------
# Configure EventGrid for Subscription
#-------------------------------------
# Create RG in subscription (if it doesn't exist already)
$tags = @{
    "cost-centre" = "1001"
    env_short = $subscription.tags.'env_short' # <> this tag contains values like prod, dev, etc (short form for sub), i use it for naming
    owner1 = ""
    owner2 = ""
}
$rgName = "rg-cc-$($subscription.tags.'env_short')-management-001"
try {
    $rg = get-azresourcegroup -name $rgname -erroraction Stop
    write-output "RG already exists, Using: $rgName"
}
catch{
    write-output "RG does not exist, Creating: $rgName"
    $rg = New-AzResourceGroup -name $rgname -Location $region -Tag $tags
}

# Register EventGrid (if it doesn't exist already)
if ((get-AzResourceProvider | where-object { $_.ProviderNamespace -eq "Microsoft.EventGrid" }).RegistrationState -ne "Registered"){
    write-warning "Microsoft.EventGrid not registered on $subscriptionName.. Registering..."
    Register-AzResourceProvider -ProviderNamespace "Microsoft.EventGrid" | out-null
}
else{
    write-output "Microsoft.EventGrid already registered on $subscriptionName. Proceeding..."
}

# Create EventGrid Topic in RG
$eventTopicName = "corpo-eg-topic-subscription-$($subscription.tags.'env_short')"

try {
    $eventTopic = Get-AzEventGridSystemTopic -name $eventTopicName -ResourceGroupName $rg.ResourceGroupName -ErrorAction Stop
    write-output "EventGridTopic already exists, Using: $eventTopicName..."
}
catch {
    write-output "Topic not found for $($subscription.Name), Creating..."
    $SysTopicparams = @{
        NAme              = $eventTopicName
        ResourceGroupName = $rg.ResourceGroupName
        Location          = "global"
        Source            = "/subscriptions/$($subscription.id)"
        TopicType         = "Microsoft.Resources.Subscriptions"
    }
    try {
        $eventTopic = New-AzEventGridSystemTopic @SysTopicparams -ErrorAction Stop
        write-output "Topic Created: $($eventTopic.Name)"
    }
    catch {
        write-error "Failed to create EventGrid Topic: $($_.Exception.Message)"
        throw
    }
}

# Create EventGrid Subscription in EventGrid Topic and assign Endpoint to Function App
$eventSubName = "new-resourceTagging"
if ($eventSub = Get-AzEventGridSystemTopicEventSubscription -ResourceGroupName $rg.ResourceGroupName -SystemTopicName $eventTopicName | Where-Object name -eq $eventSubName){
    Write-Output "EventGrid Subscription already exists Using $eventSubName"
}
else{
    $functionName   = "eventGrid-resourceOwnerTimeTagger"
    $funcResourceId = "<FunctionAppResourceID>/functions/$functionName" # <> enter resourceID of your functionApp in the angled brakcets
    $destination    = New-AzEventGridAzureFunctionEventSubscriptionDestinationObject -ResourceId $funcResourceId

    $eventSubParams = @{
        EventSubscriptionName   = $eventSubName
        SystemTopicName         = $eventTopicName
        ResourceGroupName       = $rg.ResourceGroupName
        Destination             = $destination
        FilterIncludedEventType = @("Microsoft.Resources.ResourceWriteSuccess")
    }
    try {
        $eventSub = New-AzEventGridSystemTopicEventSubscription @eventSubParams -ErrorAction Stop
        Write-Output "Event subscription created: $($eventSub.Name)"
    }
    catch {
        Write-Error "Failed to create event subscription: $($_.Exception.Message)"
        throw
    }
}

# Update filters on the Event Grid Subscription.
$advFilter = @()
$advFilter += New-AzEventGridStringNotContainsAdvancedFilterObject -key "data.operationName" -value @(
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
# $advFilter += New-AzEventGridStringNotContainsAdvancedFilterObject -Key "data.operationName" -Value @("Some.Other/Provider")
$updateEventSubParams = @{
    EventSubscriptionName   = $eventSubName
    SystemTopicName         = $eventTopic.name
    ResourceGroupName       = $rg.ResourceGroupName
    FilterAdvancedFilter    = $advFilter
}
try {
    $eventSub = Update-AzEventGridSystemTopicEventSubscription @updateEventSubParams -ErrorAction Stop
    Write-Output "Advanced filters updated for subscription: $eventSubName"
}
catch {
    Write-Error "Failed to update advanced filters: $($_.Exception.Message)"
    throw
}
