# Token
$clientID     = "<YOUR_APP_CLIENT_ID>"
$clientSecret = Get-AzKeyVaultSecret -VaultName "<YOUR_KEY_VAULT_NAME>" -Name "<YOUR_KV_SECRET_NAME_FOR_APP_SECRET>" -AsPlainText
$tenantID     = Get-AzKeyVaultSecret -VaultName "<YOUR_KEY_VAULT_NAME>" -Name "<YOUR_KV_SECRET_NAME_FOR_TENANT_ID>" -AsPlainText
$resourceURL  = "https://graph.microsoft.com/.default"

$body = @{
    client_id     = $clientId
    scope         = $resourceURL
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

$uri   = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
$token = (Invoke-RestMethod -Method POST -Uri $uri -Body $body).access_token

# Headers
$headers = @{
    Authorization  = "Bearer $token"
    "Content-Type" = "application/json"
}

# Site ID
$siteName = "<YOUR_SHAREPOINT_SITE_NAME>"
$uri      = "https://graph.microsoft.com/v1.0/sites/<YOUR_TENANT_SHAREPOINT_HOSTNAME>:/sites/$siteName"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$siteId   = $response.id

# Create File + Get its ID
$folderPath = "root:/automation/demo"
$fileName   = "MyNewExcelDoc.xlsx"
# $fileName = "MyNewWordDoc.docx" # can create other types

$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/$($folderPath):/children"
$body = @{
    name = $fileName
    file = @{}
} | ConvertTo-Json -Depth 20

$response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body
$fileId   = $response.id

# Get WorkSheets
$uri      = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response

# Get particular worksheet
$existingWorkSheet = "{00000000-0001-0000-0000-000000000000}"
$uri               = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets('$existingWorkSheet')"
$response          = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response

# Change WorkSheet Name
$existingWorkSheet = "Sheet1"
$newWorkSheet      = "MyTestWorkSheet"
$uri               = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets('$existingWorkSheet')"

$body = @{
    name = $newWorkSheet
} | ConvertTo-Json -Depth 20

$response = Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body


# Get Table
$workSheetId = "{00000000-0001-0000-0000-000000000000}"
$uri         = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets/$worksheetId/tables"
$response    = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers


# Get Table Data
$tableId = "Table2"
$uri     = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables/$tableId/range"
$values  = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).values

# convert 2D Array to Object
$columns = @($values[0])
$rows    = $values[1..($values.count-1)]

$object = foreach($row in $rows){
    $hashtable = [ordered]@{}
    for($i=0; $i -lt $columns.count; $i++){
        $hashtable[$columns[$i]] = $row[$i]
    }
    [pscustomobject]$hashtable
}
$object


# Add a row to a table / appending data
$row = @(
    "Orange"
    "Orange"
    "Medium"
)

$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables/$tableId/rows/add"
$body = @{
    values = @(
        ,$row
    )
} | ConvertTo-Json -Depth 20

$response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


#-------------
# Make a Table
#-------------
# New WorkSheet
$newWorkSheet = "Resources"
$uri          = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets/add"
$body         = @{ name = $newWorkSheet } | ConvertTo-Json -Depth 5
$response     = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body

Set-AzContext "<YOUR_AZURE_SUBSCRIPTION_CONTEXT_NAME>" | Out-Null
$object = Get-AzResource | Select-Object Name, ResourceGroupName, @{ n = "Subscription"; e = { (Get-AzContext).Subscription.Name } }, Location, ResourceId

# make 2D Array from Object
# Get columns
$columns = $object[0].PSObject.Properties.Name

# Rows
$rowsList = [System.Collections.Generic.List[object]]::new()
$rowsList.Add(@($columns)) # add headers

foreach($item in $object){ # add rows
    $row = foreach($column in $columns){
        $item.$column
    }
    $rowsList.Add(@($row))
}

# GetRangeInExcel
$colNum    = $columns.count
$endColumn = ""

while($colNum -gt 0){
    $colNum--
    $endColumn = ([char][int](65 + ($colNum % 26))) + $endColumn
    $colNum    = [math]::Floor($colNum / 26)
}

$endRow     = $rowsList.Count
$tableRange = "A1:$EndColumn$EndRow" # A1:E24

# Add data to Excel (without Table)
$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets('$newWorkSheet')/range(address='$tableRange')"
$body = @{
    values = $rowsList
} | ConvertTo-Json -Depth 10

$response = Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body

# Convert Data to a Table
$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets('$newWorkSheet')/tables/add"
$body = @{
    address    = $tableRange
    hasHeaders = $true
} | ConvertTo-Json -Depth 10

$table         = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body
$previousTable  = $table.Name

# Rename Table
$tableName = "NewResources"
$uri       = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$previousTable')"
$body      = @{ name = $tableName } | ConvertTo-Json

Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body | Out-Null

# Table Style
$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')"
$body = @{ style = "TableStyleLight12" } | ConvertTo-Json

Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body | Out-Null

# Change Font/Color
$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/headerRowRange/format/font"
$body = @{
    size  = 14
    color = "#000000"
} | ConvertTo-Json

Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body | Out-Null

# AutoFit columns for the table
$uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/range/format/autofitColumns"
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers | Out-Null

# AutoFit rows for the table
$uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/range/format/autofitRows"
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers | Out-Null


#-------------------------
# APPENDING OBJECTS TO EXCEL
#-------------------------
Set-AzContext "<YOUR_OTHER_AZURE_SUBSCRIPTION_CONTEXT_NAME>" | Out-Null

$rowsToAdd = Get-AzResource | Select-Object Name, @{ n = "Subscription"; e = { (Get-AzContext).Subscription.Name } }, Location, ResourceId

# Get columns from table header
$uri     = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/headerRowRange"
$columns = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).values[0]

# Build 2D array
$rows = [System.Collections.Generic.List[object]]::new()

foreach($obj in $rowsToAdd){
    $cells = [System.Collections.Generic.List[object]]::new()

    foreach($column in $columns){
        $prop = $obj.PSObject.Properties[$column]
        if($prop){
            $cells.Add($prop.Value)
        }
        else{
            $cells.Add("")
        }
    }

    $rows.Add($cells)
}

$uri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/rows/add"
$body = @{
    values = $rows
} | ConvertTo-Json -Depth 20

$response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body

# AutoFit columns for the table
$uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/range/format/autofitColumns"
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers | Out-Null

# AutoFit rows for the table
$uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/range/format/autofitRows"
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers | Out-Null
