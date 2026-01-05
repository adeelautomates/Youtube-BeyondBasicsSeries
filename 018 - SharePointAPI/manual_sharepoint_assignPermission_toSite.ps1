Param (
	[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()] 
		[string]$siteName,
	[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$appID,
	[Parameter(Mandatory = $true)] 
		[ValidateSet('read', 'write')]
		[string]$role
)

#----------------------
# AuthN to Azure
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

#----------------------------------------------------------------------------------------------
# AuthN to an Identity with ability to manage Service Principals and assign role "Site.Selected"
#-----------------------------------------------------------------------------------------------
$clientID = "00000-000-000-00000"
$clientSecret = Get-AzKeyVaultSecret -VaultName "example-kv-001" -Name "example-graph-identity-rw-all" -AsPlainText
$tenantID = Get-AzKeyVaultSecret -VaultName "example-kv-001" -Name "example-tenant-id" -AsPlainText 
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
    Authorization = "Bearer $Token"
    "Content-Type" = "application/json"
}

# get SP's principal ID
$uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AppId'&`$select=id,appId,displayName"
try {
	$servicePrincipal = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers).value
	write-output "AppID matched a Service Principal. Proceeding..."
}
catch {
	$errorStatus = $_
    write-error "Fail: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
	throw
}

# Assign Permission
$uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($servicePrincipal.id)/appRoleAssignments"
$body = @{
	principalId = $servicePrincipal.id
	resourceId  = "00000-000-000-00000" # Microsoft Graph service principal ID in your tenant
	appRoleId   = "883ea226-0bf2-4a8f-9f9d-92c9162a727d" # Sites.Selected
} | ConvertTo-Json -Depth 5
try {
  $result = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ErrorAction Stop
  Write-Output "Assigned role: Sites.Selected to $($servicePrincipal.displayName)"
}
catch {
	$errorStatus = $_
  	if ($errorStatus -like "*Permission being assigned already exists on the object*") {
    	Write-Output "Already has role Sites.Selected. Proceeding."
  	} 
	else {
    	write-error "Fail: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
		throw
	}
}

#-----------------------------------------------------------------------------------------------
# AuthN to an Identity with ability to grant access to sharepoint sites & grant site permissions
#-----------------------------------------------------------------------------------------------
$clientID = "00000-000-000-00000"
$clientSecret = Get-AzKeyVaultSecret -VaultName "example-kv-001" -Name "example-graph-sharepoint-rw-all" -AsPlainText
$tenantID = Get-AzKeyVaultSecret -VaultName "example-kv-001" -Name "example-tenant-id" -AsPlainText 
$resourceURL = "https://graph.microsoft.com/.default"
$body = @{
    client_id     = $clientId
    scope         = $resourceURL
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$uri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
$token = (Invoke-RestMethod -Method POST -Uri $uri -Body $body -ErrorAction Stop).access_token

$headers = @{
 	 Authorization = "Bearer $Token"
	"Content-Type" = "application/json"
}

# Confirm Site and collect its ID
try {
    $siteNameURI = [uri]::EscapeDataString("$siteName")
    $uri = "https://graph.microsoft.com/v1.0/sites?search=$siteNameURI"
    $site = (Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop).value
    $siteID, $webURL = $site.id, $site.webUrl
    if([string]::IsNullOrWhiteSpace($siteID)){
        Write-Error "No site with the name found. Exiting..."
        throw
    }
    write-output "Found Site. Collected ID. Proceeding..."
}
catch{
	$errorStatus = $_
	Write-Error "Unable to find site: [$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
	throw
}
# Grant Permissions
$grantUri = "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"
$body = @{
    roles = @($role)
    grantedToIdentities = @(
        @{
            application = @{
                id = $servicePrincipal.appId
                displayName = $servicePrincipal.displayName
            }
        }
    )
} | ConvertTo-Json -Depth 10

try {
    Invoke-RestMethod -Method POST -Uri $grantUri -Headers $headers -Body $body -ErrorAction Stop | out-null
    write-output "Assigned Permission:"
    $selectProps = @(
        @{ n = "displayName" ;  e = { $_.grantedToIdentitiesV2.application.displayName } }
        @{ n = "id" ;           e = { $_.grantedToIdentitiesV2.application.id } }
		@{ n = "spSite";        e = { $webURL} }
        @{ n = "spRole";        e = { ($_.roles -join ",") } }
        @{ n = "graphRole";   e = { "Sites.Selected" } }
    )
    (Invoke-RestMethod -Method GET -Uri $grantUri -Headers $headers).value | select-object $selectProps | convertTo-Json
}
catch{
    $errorStatus = $_
    write-error "[$($errorStatus.Exception.statuscode.value__) $($errorStatus.Exception.statuscode)] $(($errorStatus.ErrorDetails.Message | convertFrom-Json).error.message)"
	throw
}
