<#
    NAME: report_azure_getResourceData
    CREATOR: Adeel A.
    DESCRIPTION: 
        - Collects information on resources in our tenant
        - Creates Excel File
        - Uploads to Storage Account
#>

#--------------------------
# authN as Managed Identity
#--------------------------
try{
    if (-not (Get-AzContext)){ # check if there is already an account signed in
        connect-azaccount -identity | out-null # Connect as Self
        write-output "Connected as the Managed Identity of the Automation Account."
    }
    else{
        Write-Output "Using existing Azure Connection: $((Get-AzContext).Account.Id)"
    }
}
catch{
    throw "Failed to authenticate: $($_.Exception.Message)"
}

#----------
# Variables
#----------
# Fill this data out with your services
$storageAccountName = '<storageAccountName>'
$containerName = "<storageAccountContainerName>"

#--------------------------------
# Collect Data from subscriptions
#--------------------------------
write-output "Collecting Data of resources in all subscriptions"
$subs = (Get-AzSubscription).name # collect all subs
$dataCollected = New-Object System.Collections.Generic.List[pscustomobject]
foreach($sub in $subs){ # Loop through each sub
    set-azcontext $sub | Out-Null
    $resources = get-azresource | select-object Name, ResourceGroupName, Location, Tags # collect all of its resources
    foreach($resource in $resources){ # Loop through each resource
        $dataCollected.Add([pscustomobject]@{ # store data in list as it loops
            Subscription = $sub
            ResourceGroup = $resource.ResourceGroupName
            Name = $resource.Name
            Region = $resource.Location
            Tags = (($resource.tags.GetEnumerator() | ForEach-Object {"'$($_.Key) : $($_.Value)'"}) -join ', ')
        })
    }
}
if ($dataCollected.count -eq 0){
    write-output "No resource found. Ending Script"
    exit 1 # doesn't fail a job but marks it as completed
}
write-output "Total Resources Found: $($dataCollected.count). Proceeding with script"

#---------------------
# Upload Data to Excel
#---------------------
# Try to create file and upload it to temp Folder
try {
    $fileName = "ResourceInventory-$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
    $filePath = Join-Path $env:TEMP $fileName
    $dataCollected | export-excel -path $filePath -worksheetname 'Resources' -tableName 'Resources' -AutoSize -tableStyle Light10
}
catch {
    throw "Failed to create file in temp folder: $($_.Exception.Message)"
}

try{ # try to upload file to blob
    $context = New-AzStorageContext -StorageAccountName $storageAccountName -UseConnectedAccount
    Set-AzStorageBlobContent -File $filePath -Container $containerName -Blob $fileName -context $context -force | out-null 
    write-output "Blob[$fileName] uploaded successfully to storage account[$storageAccountName]"
}
catch{ # End job as a fail
    throw "Upload Failed. $($_Exception.Message)"
}
finally{ # means it will happen regardless of if try succeeds or fails
    if ( Test-Path $filePath ){ 
        write-output "Cleaning up... Deleted temp file from Automation Account."
        Remove-Item $filePath -force # delete file we created earlier from the temp folder
    }
}
write-output "Task Completed."
