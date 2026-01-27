# Connect-AzAccount

#--------------------------------------
# Create Resource and Assign RBAC
#-------------------------------------- 
$subscriptionID = "00000-00000-0000-0000"
$resourceGroupName = "rg-placeholder"
$storageAccountName = "storageaccountplaceholder"
$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com/").Token | ConvertFrom-SecureString -AsPlainText
$headers = @{
    Authorization  = "Bearer $token"
    "Content-Type" = "application/json"
}

# Create Storage Account
$body = @{
    location   = "canadacentral"
    sku        = @{ name = "Standard_LRS" }
    kind       = "StorageV2"
    properties = @{
        accessTier                   = "Hot"
        minimumTlsVersion            = "TLS1_2"
        allowSharedKeyAccess         = $false
        supportsHttpsTrafficOnly     = $true
        defaultToOAuthAuthentication = $true
    }
} | ConvertTo-Json -Depth 5
$uri = "https://management.azure.com/subscriptions/$($subscriptionID)/resourceGroups/$($resourceGroupName)/providers/Microsoft.Storage/storageAccounts/$($storageAccountName)?api-version=2025-06-01"
invoke-restmethod -Method "PUT" -Headers $headers -Uri $uri -Body $body

# SA Resource ID (as put does not output anything to capture from api above)
$storageAccountId = (invoke-restmethod -Method "GET" -Headers $headers -Uri $uri).id

# Assign RBAC
$roles = @(
    "00000-00000-0000-0000" # "Storage Blob Data Contributor"
    "00000-00000-0000-0000" # "Storage Table Data Contributor" 
    "00000-00000-0000-0000" # "Storage Queue Data Contributor" 
)
foreach ($role in $roles) {
    $assignmentId = [guid]::NewGuid().Guid
    $body = @{
        properties = @{
            roleDefinitionId = "/subscriptions/$subscriptionID/providers/Microsoft.Authorization/roleDefinitions/$($role)"
            principalId      = "00000-00000-0000-0000" # me
        }
    } | ConvertTo-Json -Depth 3
    $uri = "https://management.azure.com$($storageAccountId)/providers/Microsoft.Authorization/roleAssignments/$($assignmentId)?api-version=2022-04-01"
    Invoke-RestMethod -Method PUT -Headers $headers -Uri $uri -Body $body
}


#---------------------------------
# Create Container in Blob Storage
#---------------------------------
$SAtoken = (Get-AzAccessToken -ResourceUrl "https://storage.azure.com/").Token | ConvertFrom-SecureString -AsPlainText
$headers = @{
    Authorization = "Bearer $SAToken"
    "x-ms-version" = "2026-02-06"
    # "x-ms-date" = [DateTime]::UtcNow.ToString("R")
}
$containerName = "container-placeholder"
$uri = "https://$($storageAccountName).blob.core.windows.net/$($containerName)?restype=container"
invoke-restmethod -Method PUT -Uri $uri -Headers $headers

#----------------------------
# Upload File to Blob Storage
#----------------------------
$headers = @{
    Authorization = "Bearer $SAToken"
    "x-ms-version" = "2026-02-06"
    "x-ms-blob-type" = "BlockBlob"
    "Content-Type"  = "application/octet-stream"
}
$filePath = "C:\test\MyFiles\firstBlob.txt"
$blobName = [System.IO.Path]::GetFileName($filePath)
$uri = "https://$($storageAccountName).blob.core.windows.net/$containerName/$blobName"
Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -InFile $filePath

#--------------------------------------
# Upload Multiple Files to Blob Storage
#-------------------------------------- 
$rootPath = "C:\test\MyFiles"
(Get-ChildItem -Path $rootPath -File -Recurse) | ForEach-Object {
    $blobName = $_.FullName.SubString($rootPath.Length +1)
    $uri = "https://$($storageAccountName).blob.core.windows.net/$containerName/$blobName"
    Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -InFile $_.FullName
}

#--------------------------------------
# Get Files + Metadata
#-------------------------------------- 
$headers = @{
    Authorization  = "Bearer $SAtoken"
    "x-ms-version" = "2026-02-06"
}
$uri = "https://$($storageAccountName).blob.core.windows.net/$($containerName)?restype=container&comp=list"
[xml]$xml = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).TrimStart([char]0xFEFF) # Trim the BOM character
$xml
$xml.EnumerationResults.blobs.blob[0].name
$xml.EnumerationResults.blobs.blob[0].properties

$xml.EnumerationResults.Blobs.Blob | ForEach-Object {
    [pscustomobject]@{
        Name = $_.name
        CreationTime  = $_.Properties.'Creation-Time'
        LastModified  = $_.Properties.'Last-Modified'
        ContentLength = $_.Properties.'Content-Length'
        ContentType   = $_.Properties.'Content-Type'
        AccessTier    = $_.Properties.AccessTier
    }
}
#-----------------------------------------------
# view data of a blob file (if its simple files)
#-----------------------------------------------
$blobName = "movieDetails.txt"
$uri = "https://$($storageAccountName).blob.core.windows.net/$($containerName)/$($blobName)"
Invoke-RestMethod -Method GET -Uri $uri -Headers $headers


$blobName = "movieDetails.csv"
$uri = "https://$($storageAccountName).blob.core.windows.net/$($containerName)/$($blobName)"
Invoke-RestMethod -Method GET -Uri $uri -Headers $headers | ConvertFrom-Csv # for them you need to convertFrom-CSV


#---------------------
# Download single blob
#---------------------
$blobName  = "testWordDoc.docx"
$directory = "C:\test\Downloads"
$destPath  = Join-Path $directory $blobName
$uri = "https://$($storageAccountName).blob.core.windows.net/$($containerName)/$($blobName)"
Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -OutFile $destPath

#---------------------
# Download multi blob
#---------------------
$directory = "C:\test\Downloads"
$uri  = "https://$($storageAccountName).blob.core.windows.net/$($containerName)?restype=container&comp=list"
[xml]$blobs = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).TrimStart([char]0xFEFF)

foreach ($blob in $blobs.EnumerationResults.Blobs.Blob.name){
    $destPath  = Join-Path $directory $blob
    $destDir = Split-Path $destPath -Parent # in case we have virtual folders, lets make them as sub folders
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    $blobUri = "https://$($storageAccountName).blob.core.windows.net/$($containerName)/$($blob)"
    Invoke-RestMethod -Method GET -Uri $blobUri -Headers $headers -OutFile $destPath
}

#-----------
# Make Table
#-----------
$SAtoken = (Get-AzAccessToken -ResourceUrl "https://storage.azure.com").Token | ConvertFrom-SecureString -AsPlainText
$headers = @{
    Authorization  = "Bearer $SAtoken"
    "x-ms-version" = "2026-02-06"
    Accept         = 'application/json;odata=nometadata'
    'Content-type' = 'application/json'
}
$tableName = "table-placeholder"
$uri  = "https://$($storageAccountName).table.core.windows.net/Tables"
$body = @{ 
    TableName = $tableName  # must be alphanumeric
} | ConvertTo-Json 
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


#-----------------
# ADD Row to Table
#-----------------
$body = @{
    PartitionKey = "Test"
    RowKey = "1"
    Name  = "DOES IT WORK!?"
} | ConvertTo-Json
$uri  = "https://$($storageAccountName).table.core.windows.net/$($tableName)"
Invoke-RestMethod -method POST -uri $uri -headers $headers -body $body

#----------------------------
# ADD Multiple Lines to Table
#----------------------------
$appSecretLifecycleMgmt = @(
    [pscustomobject]@{
        tenantId     = "00000-00000-0000-0000"
        displayName   = "app-display-name-01"
        appId        = "00000-00000-0000-0000"
        keyVaultName = "keyvault-placeholder"
    }
    [pscustomobject]@{
        tenantId     = "00000-00000-0000-0000"
        displayName   = "app-display-name-02"
        appId        = "00000-00000-0000-0000"
        keyVaultName = "keyvault-placeholder"
    }
    [pscustomobject]@{
        tenantId     = "00000-00000-0000-0000"
        displayName   = "app-display-name-03"
        appId        = "00000-00000-0000-0000"
        keyVaultName = "keyvault-placeholder"
    }
    [pscustomobject]@{
        tenantId     = "00000-00000-0000-0000"
        displayName  = "app-display-name-04"
        appId        = "00000-00000-0000-0000"
        keyVaultName = "keyvault-placeholder"
    }
)
$uri = "https://$($storageAccountName).table.core.windows.net/$($tableName)"
$partitionKey = "appSecretLifecycleMgmt"

foreach ($app in $appSecretLifecycleMgmt) {
    $body = @{
        PartitionKey = $partitionKey
        RowKey       = $app.appId # just make it unique
        appId        = $app.appId
        displayName  = $app.displayName
        tenantId     = $app.tenantId
        keyVaultName = $app.keyVaultName
    } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -method POST -uri $uri -headers $headers -body $body
}

#---------------
# Get Table Rows
#---------------
$uri = "https://$($storageAccountName).table.core.windows.net/$($tableName)"
$data = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).value
$data

#----------------------------------
# Get Table Rows specific partition
#----------------------------------
$partitionKey = "appSecretLifecycleMgmt"
$uri = "https://$($storageAccountName).table.core.windows.net/$($tableName)" + "?`$filter=PartitionKey eq '$partitionKey'"
$data = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).value
$data

# search by tenantID
 "?`$filter=tenantID eq '$($app.tenantId)'"
# search for partition and row key
 "?`$filter=PartitionKey eq '$partitionKey' and RowKey eq '$rowKey'"

#-----------------------------------------------------------------
# modify existing key using Merge (add without replacing existing)
#-----------------------------------------------------------------
$headers = @{
    Authorization  = "Bearer $SAtoken"
    Accept         = "application/json;odata=nometadata"
    "Content-Type" = "application/json"
    "x-ms-version" = "2026-02-06"
    "If-Match"     = "*"   # add this for updates
}
$partitionKey = "Test"
$rowKey = "1"
$body = @{
    MyNewMethod = "Merge"
    name        = "DOES IT WORK WITH MERGE?"
} | ConvertTo-Json -Depth 5
$uri = "https://$($storageAccountName).table.core.windows.net/$($tableName)(PartitionKey='$($partitionKey)',RowKey='$($rowKey)')"
Invoke-RestMethod -Method MERGE -Uri $Uri -Headers $headers -Body $body

#-----------------------------------------------------------------
# modify existing key using Replace
#-----------------------------------------------------------------
$partitionKey = "Test"
$rowKey = "1"
$body = @{
    MyNewMethod = "REPLACE"
} | ConvertTo-Json -Depth 5
$uri = "https://$($storageAccountName).table.core.windows.net/$($tableName)(PartitionKey='$partitionKey',RowKey='$rowKey')"
Invoke-RestMethod -Method Put -Uri $uri -Headers $headers -Body $body

#--------------
# Delete Entity
#--------------
$partitionKey = "Test"
$rowKey = "1"
$uri = "https://$($storageAccountName).table.core.windows.net/$($tableName)(PartitionKey='$partitionKey',RowKey='$rowKey')"
Invoke-RestMethod -Method Delete -Uri $uri -Headers $headers


#----------------
# Make Queue
#----------------
$SAtoken = (Get-AzAccessToken -ResourceUrl "https://storage.azure.com").Token | ConvertFrom-SecureString -AsPlainText
$queueName = "queue-placeholder"
$headers = @{
    Authorization  = "Bearer $SAtoken" 
    "x-ms-version" = "2026-02-06"
}
$uri = "https://$($storageAccountName).queue.core.windows.net/$($queueName)"
Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers

#---------------------
# Adding Data to Queues (usually done by some service or app, like event grids)
#---------------------
$messages = @(
    @{
        eventType = "RBACChange"
        severity  = "High"
        summary   = "Role assignment added"
        resource  = "/subscriptions/00000-00000-0000-0000/resourceGroups/rg-placeholder"
    }
    @{
        eventType = "TagMissing"
        severity  = "Medium"
        summary   = "Missing owner tag"
        resource  = "/subscriptions/00000-00000-0000-0000/resourceGroups/rg-placeholder/providers/Microsoft.Storage/storageAccounts/storageaccountplaceholder"
    }
    @{
        eventType = "PolicyNonCompliant"
        severity  = "High"
        summary   = "Policy denied a deployment"
        resource  = "/subscriptions/00000-00000-0000-0000"
    }
    @{
        eventType = "SecretExpiring"
        severity  = "High"
        summary   = "Secret expiring in 7 days"
        resource  = "KeyVault: keyvault-placeholder"
    }
    @{
        eventType = "ResourceCreated"
        severity  = "Low"
        summary   = "New blob uploaded"
        resource  = "Container: container-placeholder"
    }
    @{
        eventType = "RBACChange"
        severity  = "Medium"
        summary   = "Role assignment removed"
        resource  = "/subscriptions/00000-00000-0000-0000"
    }
    @{
        eventType = "TagMissing"
        severity  = "Low"
        summary   = "Missing cost-centre tag"
        resource  = "/subscriptions/00000-00000-0000-0000/resourceGroups/rg-placeholder"
    }
    @{
        eventType = "SecretExpiring"
        severity  = "Medium"
        summary   = "Certificate expiring in 14 days"
        resource  = "KeyVault: keyvault-placeholder"
    }
    @{
        eventType = "ResourceCreated"
        severity  = "Low"
        summary   = "New queue message received"
        resource  = "Queue: queue-placeholder"
    }
)

$uri = "https://$($storageAccountName).queue.core.windows.net/$($queueName)/messages"
foreach($m in $messages){
    $payload = @{
        timeUtc = [DateTime]::UtcNow.ToString("o")
        eventType = $m.eventType
        severity = $m.severity
        summary  = $m.summary
        resource = $m.resource
    } | ConvertTo-Json -Depth 5 -Compress
    $Body = @"
<QueueMessage>  
    <MessageText>$payload</MessageText>  
</QueueMessage>  
"@
    Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/xml" | Out-Null
}


#-----------------------
# Reading Single Message
#-----------------------
$messageCount = 1 # Limit 32 
$timeOut = 30 # how long does it dissapear for (30 seconds)
$uri = "https://$($storageAccountName).queue.core.windows.net/$($queueName)/messages?numofmessages=$($messageCount)&visibilitytimeout=$timeOut"
[xml]$msg = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).TrimStart([char]0xFEFF)
$msg.QueueMessagesList.QueueMessage.MessageText | ConvertFrom-Json

# then delete it once read
$messageId  = $msg.QueueMessagesList.QueueMessage.MessageId
$popReceipt = $msg.QueueMessagesList.QueueMessage.PopReceipt
$uri = "https://$($storageAccountName).queue.core.windows.net/$($queueName)/messages/$($messageId)?popreceipt=$($popReceipt)"
Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $headers


#----------------------------------------------------------------------
# Reading multiple Messages, taking aciton on them and deleting message
#----------------------------------------------------------------------

$messageCount = 3 # Limit 32 
$timeOut = 30 # how long does it dissapear for (30 seconds)
$uri = "https://$($storageAccountName).queue.core.windows.net/$($queueName)/messages?numofmessages=$($messageCount)&visibilitytimeout=$timeOut"
[xml]$msg = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).TrimStart([char]0xFEFF)

write-output "--------------------------"
foreach ($item in $msg.QueueMessagesList.QueueMessage){
    write-output "Doing stuff on... $(($item.MessageText | ConvertFrom-Json).resource)"
    write-output "Task Completed. removing from queue"

    $messageId  = $item.MessageId
    $popReceipt = $item.PopReceipt

    $uri = "https://$($storageAccountName).queue.core.windows.net/$($queueName)/messages/$($messageId)?popreceipt=$($popReceipt)"
    Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $headers | out-null

    write-output "Deleted Message"
    write-output "--------------------------"
}
