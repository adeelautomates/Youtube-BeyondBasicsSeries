# code for function App
param($eventGridEvent, $TriggerMetadata)

write-output "------------------------------"
$eventGridEvent.subject
$eventGridEvent.action
$eventGridEvent.data.authorization.evidence.principalType
$eventGridEvent.data.authorization.evidence.principalId
$eventGridEvent.eventtime
write-output "------------------------------"
$eventGridEvent | ConvertTo-Json -Depth 10 

# Get Graph token using managed identity
$token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
$headers = @{
    Authorization = "Bearer $token"
}

# Get Tag Values
$resourceID = $eventGridEvent.subject
$principalID = $eventGridEvent.data.authorization.evidence.principalId
$principalType = $eventGridEvent.data.authorization.evidence.principalType

if($principalType -eq "User"){
    $identity = invoke-restmethod -method get -uri "https://graph.microsoft.com/v1.0/users/$principalID" -Headers $headers | select-object userPrincipalName
    $creatorTag = $identity.userPrincipalName
    $creatorTypeTag = $principalType
}
elseif($principalType -eq "ServicePrincipal"){
    $identity = invoke-restmethod -method get -uri "https://graph.microsoft.com/v1.0/serviceprincipals/$principalID" -Headers $headers | select-object displayName, servicePrincipalType
    $creatorTag = $identity.displayName
    $creatorTypeTag =  $identity.servicePrincipalType
}
else{
    $creatorTag = $principalID
    $creatorTypeTag = "Unknown"
}
$creationTimeTag = ($eventGridEvent.eventtime).ToString("yyyy-MM-dd HH:mm")


# Apply Tags
try{
    $tagInfo = Get-AzTag -ResourceID $resourceID
    if(!$taginfo.properties.TagsProperty.creator){
        Update-AzTag -ResourceId $resourceID -Tag @{"creator" = $creatorTag; "creatorType" = $creatorTypeTag} -Operation Merge -ErrorAction Stop | Out-Null
        write-output "Tag Applied: creator : $creatorTag"
        write-output "Tag Applied: creatorType : $creatorTypeTag"
    }
    if (!$tagInfo.properties.TagsProperty.creationTime){
        Update-AzTag -ResourceId $resourceID -Tag @{"creationTime"="$creationTimeTag"} -Operation Merge -ErrorAction Stop | Out-Null
        write-output "Tag Applied: creationTime : $creationTimeTag"
    }
}
catch {
    write-error "Failed to Update Tag(s) for resource: $($_.exception.message)"
}