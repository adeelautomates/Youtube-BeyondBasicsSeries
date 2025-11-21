<#
.SYNOPSIS
Retrieves an OAuth 2.0 access token from Microsoft Entra ID (Azure AD) for Microsoft Graph.

.DESCRIPTION
This function requests an application (client credentials) token from the Microsoft identity platform.
It uses the provided client ID, client secret, and tenant ID to authenticate and returns the token object.
The token can then be used in REST API calls to Microsoft Graph.

.PARAMETER ClientId
The Application (Client) ID of the registered app in Microsoft Entra ID.

.PARAMETER ClientSecret
The client secret associated with the application.

.PARAMETER TenantID
The Directory (Tenant) ID of the Azure AD tenant to authenticate against.

.EXAMPLE
$token = New-CorpoGraphOAuthToken -clientId 'xxxx' -clientSecret 'xxxx' -tenantID 'xxxx'
$token.access_token

This example retrieves an OAuth token and displays the access token value.

.OUTPUTS
[pscustomobject]
Includes access_token, token_type, expires_in, and scope.

.NOTES
Author: Adeel Anwar
Version: 1.0  
Requires: Microsoft Graph application with client credentials permission.
#>
function New-CorpoGraphOAuthToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
            $ClientId,
        [Parameter(Mandatory = $true)] 
            $ClientSecret,
        [Parameter(Mandatory = $true)] 
            $TenantID
    )
    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }
    $uri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
    try {
        $token = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ErrorAction Stop
        return $token
    }
    catch {
        write-error  "FAILED: $(($_ | convertFrom-json).error_description)"
        return
    }
}