<#
    NAME: manual_identity_permissionReport
    CREATOR: Adeel A.
    DESCRIPTION: 
        - Generate report on all identities accessing Azure, EntraID, M365, Graph
        - Place the report on SharePoint Excel (MyTeamSite > Automation > Demo > )
#>

#-------------
# Authenticate
#-------------
# To Azure
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

# To Graph (Enter your data needed to authN to graph as SP)
set-azcontext "" | out-null
$clientID = ""
$clientSecret = Get-AzKeyVaultSecret -VaultName "" -Name "" -AsPlainText  | ConvertTo-SecureString -AsPlainText
$tenantID = Get-AzKeyVaultSecret -VaultName "" -Name "" -AsPlainText 
$cred = [System.Management.Automation.PSCredential]::new($ClientID, $clientSecret)
Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $cred | out-null
Write-Output "For any Graph Queries, Using: $((Get-MgContext).AppName)"


#------------------------------------------------------
# Active RBAC Assignments WorkSheet - No  JIT/PIM based
#------------------------------------------------------
function New-RbacRow{
    param ($Role, $ScopeName, $ScopeType)
    [pscustomobject]@{
        DisplayName     = $Role.DisplayName
        ID              = $Role.ObjectID
        IdentityType    = $Role.ObjectType
        RBAC            = $Role.RoleDefinitionName
        Scope           = $ScopeName
        ScopeType       = $ScopeType
    }
}
$worksheetRBAC = [System.Collections.Generic.List[pscustomobject]]::new() # store data for excel

# Management Groups
$managementGroups = Get-AzManagementGroup
foreach($mg in $managementGroups){
    $scope = $mg.id
    $roles = Get-AzRoleAssignment -scope $scope
    foreach ($role in $roles){
        if ($role.scope -eq $scope){
            $worksheetRBAC.add( (New-RbacRow -Role $role -ScopeName $mg.DisplayName -ScopeType "ManagementGroup") )
        }
    }
}
# Subscriptions / RGs / (Maybe Resources)
$subscriptions = Get-AzSubscription
foreach($sub in $subscriptions){
    Set-AzContext -Subscription $sub.Id | Out-Null
    $scope = "/subscriptions/$($sub.id)"
    $roles = Get-AzRoleAssignment -scope $scope
    foreach ($role in $roles){
        if ($role.scope -eq $scope){
            $worksheetRBAC.add( (New-RbacRow -Role $role -ScopeName $sub.name -ScopeType "Subscription") )
        }
    }
    $resourceGroups = Get-AzResourceGroup
    foreach ($rg in $resourceGroups){
        $scope = $rg.ResourceId
        $roles = Get-AzRoleAssignment -scope $scope
        foreach ($role in $roles){
            if ($role.scope -eq $scope){
                $worksheetRBAC.add( (New-RbacRow -Role $role -ScopeName $rg.ResourceGroupName -ScopeType "ResourceGroup") )
            }
        }
    }
}
Write-Output "Collected RBAC data"


#------------------------
# EntraID Roles WorkSheet
#------------------------
$worksheetRoles = [System.Collections.Generic.List[pscustomobject]]::new() # store data for excel
$rolesAssigned = Get-MgRoleManagementDirectoryRoleAssignment -All
$definitions = Get-MgRoleManagementDirectoryRoleDefinition -all
foreach($role in $rolesAssigned){
    $principalDetails = (Get-MgDirectoryObject -DirectoryObjectId $role.PrincipalId).AdditionalProperties
    $worksheetRoles.add([PSCustomObject]@{
        DisplayName = $principalDetails.displayName
        ID = $role.PrincipalId
        IdentityType = ($principalDetails.'@odata.type' -replace '#microsoft.graph.', '') # remove what it starts with
        RoleName = ($definitions | where-object id -eq $role.RoleDefinitionId).DisplayName
        RoleID   = $role.RoleDefinitionId
    })
}
Write-Output "Collected EntraID Data"

#---------------------------------------------------------------
# Graph Roles WorkSheet - App Roles & User Delegated Permissions
#---------------------------------------------------------------
$worksheetGraphPerms = [System.Collections.Generic.List[pscustomobject]]::new()
# Microsoft Graph Roles for Service Principals (Not Delegated)
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$ignoreList = @(
    "Microsoft-Developer Program-Sample Data Packs"
)
$servicePrincipals = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $graphSp.id -all | where-object PrincipalDisplayName -NotIn $ignoreList | group-object PrincipalId
foreach ($sp in $servicePrincipals) {
    $roleList = foreach ($role in $sp.group){
        $role | ForEach-Object {
            ($graphSp.AppRoles | where-object Id -eq $_.appRoleId).Value
        }
    }
    $worksheetGraphPerms.Add([PSCustomObject]@{
        DisplayName = $sp.group[0].PrincipalDisplayName
        ID = $sp.name
        IdentityType = $sp.group[0].PrincipalType
        Application = $graphSp.DisplayName
        Permissions = $roleList -join ", "
    })
}

# Get delegated permissions from other Graph Apps that are used by users in automation.
$GraphSPs = @{ # NAME = Object ID (not app ID)
    "Graph Explorer"                     = "5b9295ca-0e9d-4f00-9264-18efe64047ed"
    "Microsoft Graph Command Line Tools" = "7417c131-11e1-4677-95a8-1e8eae574b78"
}

$grant = foreach ($id in $GraphSPs.values){
    Get-MgOauth2PermissionGrant -Filter "ClientId eq '$id'" | Where-Object ConsentType -eq Principal
}
$excludeScopes = @("openid", "profile", "offline_access", "email")
foreach ($grant in $grants){
    $principalDetails = (Get-MgDirectoryObject -DirectoryObjectId $grant.PrincipalId).AdditionalProperties
    $roleList = ($grants.scope).trim() -split ' ' | where-object { $_ -notin $excludeScopes }
    $worksheetGraphPerms.Add([pscustomobject]@{
        DisplayName  = $principalDetails.displayName
        ID           = $grant.principalID
        IdentityType = ($principalDetails.'@odata.type' -replace '#microsoft.graph.', '')
        Application  = ($GraphSPs.GetEnumerator() | Where-Object Value -eq $grant.ClientId).Key
        Permissions  = $roleList -join ", "
    })
}
Write-Output "Collected Graph Scopes for Users"

#------------------------------------------------------------------------------------
# Group Reference Worksheet - Collect groups from other sheets and get members/owners
#------------------------------------------------------------------------------------
$allGroups = [System.Collections.Generic.List[object]]::new()
$allGroups.AddRange(@(
    $worksheetRBAC  | Where-object IdentityType -eq "Group" 
                    | Sort-Object Id -Unique 
                    | Select-Object Id, DisplayName, @{n = "Assignment"; e = {"Azure RBAC"}}
    $worksheetRoles  | Where-object IdentityType -eq "Group" 
                    | Sort-Object Id -Unique 
                    | Select-Object Id, DisplayName, @{n = "Assignment"; e = {"Entra Role"}}
))
$worksheetGroups = [System.Collections.Generic.List[pscustomobject]]::new()
foreach($group in $allGroups){
    $members = (Get-MgGroupMember -All -GroupId $group.id).AdditionalProperties.displayName -join ", "
    $owners = (Get-MgGroupOwner -All -GroupId $group.id).AdditionalProperties.displayName -join ", "
    $worksheetGroups.Add([PSCustomObject]@{
        GroupName = $group.DisplayName
        GroupID   = $group.id
        Assignment = $group.Assignment
        Members = $members
        Owners = $owners
    })
}
Write-Output "Created Group Reference Data"

#------------------------------------------------
# Create an Excel Doc in Sharepoint with our data
#------------------------------------------------
function New-ExcelOnlineTable {
    param(
        [Parameter(Mandatory)] $SiteID,
        [Parameter(Mandatory)] $FileID, # provide File ID of existing File
        [Parameter(Mandatory)] $Headers,
        [Parameter(Mandatory)] $WorksheetName,
        [Parameter(Mandatory)] $Object
    )

    # Set Properties as Object
    $columns = $object[0].PSObject.Properties.Name

    # Set Object data to 2D Array 
    $rowsList = [System.Collections.Generic.List[object]]::new()
    $rowsList.Add(@($columns)) # header row (as one row)
    foreach ($item in $object) {
        # data rows
        $row = foreach ($column in $columns) {
            $item.$column 
        }
        $rowsList.Add(@($row))
    }

    # Get Table Range to data in
    $colNum = $columns.Count
    $endColumn = ""
    while ($colNum -gt 0) {
        # this method so it can go beyond A-Z for width (ie AAA)
        $colNum--
        $endColumn = ([char][int](65 + ($colNum % 26))) + $endColumn
        $colNum = [math]::Floor($colNum / 26)
    }
    $endRow = $rowsList.Count
    $tableRange = "A1:$endColumn$endRow"

    # Add Data to Excel In WorkSheet
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets('$WorksheetName')/range(address='$tableRange')"
    $body = @{ 
        values = $rowsList
    } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body | Out-Null

    # Create Table In WorkSheet
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/worksheets('$WorksheetName')/tables/add"
    $body = @{ 
        address    = $tableRange; 
        hasHeaders = $true
    } | ConvertTo-Json
    $tableName = (Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body).name

    # Change Style of Table
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')"
    $body = @{ style = "TableStyleLight12" } | ConvertTo-Json
    Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body | Out-Null
    
    # Change Font Color & Size of Headers in Table
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$fileId/workbook/tables('$tableName')/headerRowRange/format/font"
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
}

# Get Token, Enter your Creds here
set-azcontext "Corpo-Management" | out-null
$clientID = ""
$clientSecret = Get-AzKeyVaultSecret -VaultName "" -Name "" -AsPlainText 
$tenantID = Get-AzKeyVaultSecret -VaultName "" -Name "" -AsPlainText 
$resourceURL = "https://graph.microsoft.com/.default"
$body = @{
    client_id     = $clientId
    scope         = $resourceURL
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$uri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
$token = (Invoke-RestMethod -Method POST -Uri $uri -Body $body).access_token
$headers = @{
    Authorization  = "Bearer $token"
    "Content-Type" = "application/json"
}

# Next lets get our site ID so we can work with this site
$siteName = 'MyTeamSite'
$uri = "https://graph.microsoft.com/v1.0/sites/lb4s.sharepoint.com:/sites/$siteName"
$siteId = (invoke-RestMethod -Method GET -Uri $uri -Headers $headers).id


# Lets make our file
$fileName = "IdentityFullReport-$((Get-Date).ToString("yyyy-MM-dd")).xlsx"
$folderPath = "root:/automation/demo"
$uri = "https://graph.microsoft.com/v1.0/sites/$($siteId)/drive/$($folderPath):/children"
$body = @{ 
    name = "$fileName"
    file = @{}
} | ConvertTo-Json -Depth 20
$fileCreated = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -body $Body


$ExcelFileData = @(
    [pscustomobject]@{
        WorksheetName = "RBAC Permissions"
        WorksheetData = $worksheetRBAC
    }
    [pscustomobject]@{
        WorksheetName = "EntraID Roles"
        WorksheetData = $worksheetRoles
    }
    [pscustomobject]@{
        WorksheetName = "Graph Roles"
        WorksheetData = $worksheetGraphPerms
    }
    [pscustomobject]@{
        WorksheetName = "Group Member Reference"
        WorksheetData = $worksheetGroups
    }
)

foreach($item in $ExcelFileData){
    # Create WorkSheet
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$($fileCreated.id)/workbook/worksheets/add"
    $body = @{ name = $item.WorksheetName } | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body | Out-Null

    # Create Table
    New-ExcelOnlineTable -SiteID $SiteID -FileID $fileCreated.id -Headers $headers -WorksheetName $item.WorksheetName -Object $item.WorksheetData
}

# More simple to just delete this sheet then to rename it as the first one
$deleteDefaultWorkSheet = "Sheet1"
$uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$($fileCreated.id)/workbook/worksheets/$deleteDefaultWorkSheet"
Invoke-RestMethod -Method DELETE -Uri $uri -Headers $headers | Out-Null

# final output with url link
Write-Output "Created Excel File: $($fileCreated.webUrl)"
