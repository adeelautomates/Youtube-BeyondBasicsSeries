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
$token = New-MsGraphOauthToken -clientId 'xxxx' -clientSecret 'xxxx' -tenantID 'xxxx'
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
function New-MsGraphOauthToken {
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


<#
.SYNOPSIS
Generates a simple HTML table from a PowerShell object.

.DESCRIPTION
The New-SimpleHtmlTable function takes one or more PSCustomObjects
and converts them into a formatted HTML <table> string.
Each property of the input object becomes a column, and each
object instance becomes a row.  
Useful for creating clean, readable HTML Tables

.PARAMETER Object
One or more PowerShell objects (PSCustomObject) whose properties
will be used as table columns.

.EXAMPLE
$data = Get-Process | Select-Object Name, CPU, Id -First 5
$html = New-SimpleHtmlTable -Object $data
$html | Out-File "C:\temp\process.html"

Generates an HTML table of the first five running processes.

.EXAMPLE
$resources = Get-AzResource | Select-Object Name, ResourceGroupName, Location
$table = New-SimpleHtmlTable -Object $resources
$body  = "<b>Azure Resources:</b><br><br>$table"

Creates an HTML fragment for embedding resource data in an email.

.OUTPUTS
[String]  
Returns a string containing the HTML markup for the table.

.NOTES
Author: Adeel Anwar  
Version: 1.0  
Intended for use in scripts or runbooks that send HTML reports or email notifications.
#>
function New-SimpleHtmlTable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [pscustomobject]$object # only accepts pscustomobjects to make the table
    )
    # Form column using the object's properties
    $columns = $object[0].PSObject.Properties.Name 
    $htmlTHData = "" 
    $htmlTHData = ($columns | ForEach-Object {"        <th>$_</th>"}) -join "`n"
$htmlTableColumns = @"
    <tr align="left" style="background-color: #e7e7e7; color: #000000; padding: 10px; font-size: 16px">
$htmlTHData
    </tr>
"@
    # Form rows using the properties data
    $htmlTableRows = ""
    foreach($item in $object){
        $htmlTableRow = "    <tr style='border-top: 1px solid #e7e7e7; border-bottom: 1px solid #e7e7e7; padding: 10px;'>`n"
        foreach($column in $columns){
            $value = $item.$column
            $htmlTableRow += "      <td>$value</td>`n"
        }
        $htmlTableRow += "    </tr>`n"
        $htmlTableRows += $htmlTableRow
    }
    $htmlTableRows = $htmlTableRows.TrimEnd("`n")
    # Put it all in a table
$htmlTable = @"
<table cellpadding="10" cellspacing="0" border="0" style="border-collapse: collapse; width:auto; display:inline-table; font-size: 14px; font-family: Calibri, Arial, sans-serif;">
$htmlTableColumns
$htmlTableRows
</table>
"@
    # Return the HTML table
    return $htmlTable
}


<#
.SYNOPSIS
Sends an email using Microsoft Graph API with optional CC, BCC, Attachments and HTML Body.

.DESCRIPTION
The Send-MsGraphMail function sends mail through Microsoft Graph using an authenticated token.
It supports plain text or HTML Body, high importance flag and one or more attachments.
Designed for use in Azure Automation or PowerShell Environments with Service Principal Authentication.

.PARAMETER Token
Paramter for Microsoft Graph Access Token Object containing 'access_token'. Required for Authentication

.PARAMETER Subject
Subject line of the email message.

.PARAMETER Body
Main body content of the email message. Supports plain text by default or HTML when -HTMLBody is used.

.PARAMETER From
Sender address. Must be one of the validated sender accounts defined in the ValidateSet.

.PARAMETER To
One or more recipient email addresses.

.PARAMETER Cc
Optional carbon copy recipients.

.PARAMETER Bcc
Optional blind carbon copy recipients.

.PARAMETER Importance
Switch parameter. Marks the email as high importance when present.

.PARAMETER HTMLBody
Switch parameter. Sends the email body as HTML when present.

.PARAMETER Attachment
Provide filepaths to the attachments you would like to include in the email

.EXAMPLE
Send-MsGraphMail -Token $token -From "corpo-automation@lb4s.onmicrosoft.com" `
-To "user@domain.com" -Subject "Test Email" -Body "This is a test."

.EXAMPLE
$params = @{
    Token = $token
    Subject = "Test Email"
    From = "corpo-automation@lb4s.onmicrosoft.com"
    To = "userA@lb4s.onmicrosoft.com"
    Cc = "userB@lb4s.onmicrosoft.com", "userC@lb4s.onmicrosoft.com"
    Body = "<b>HELLO BOLD WORDS. THIS IS HTML!</b>"
    HTMLBody = $true
    Importance = $true
    ErrorAction = 'Stop'
    Attachment = "C:\test\TestFile1.txt", "C:\test\TestFile2.txt"
}
try {
    Send-MsGraphMail @params 
}
catch{
    write-Error $_
    throw
}

.OUTPUTS
[pscustomobject]
Contains Fields: Status, Subject, From, To, Cc, Bcc, Importance, FilesAttached, EmailBodyType

.NOTES
Author: Adeel Anwar  
Version: 1.0  
Requires: Microsoft Graph Mail.Send permission on the service principal.
#>
function Send-MsGraphMail{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)] 
            [pscustomobject]$Token,
        [Parameter(Mandatory = $true)] 
            [string]$Subject,
        [Parameter(Mandatory = $true)] 
            [string]$Body,
        [Parameter(Mandatory = $true)] 
            [ValidateSet('corpo-automation@lb4s.onmicrosoft.com', 'corpo-helpdesk@lb4s.onmicrosoft.com')]
            [string]$From,
        [Parameter(Mandatory = $true)] 
            [ValidatePattern('(?i)^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$')] # Regex to see if it looks like an email
            [string[]]$To,
        [Parameter(Mandatory = $false)] 
            [ValidatePattern('(?i)^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$')] # Regex to see if it looks like an email
            [string[]]$Cc,
        [Parameter(Mandatory = $false)] 
            [ValidatePattern('(?i)^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$')] # Regex to see if it looks like an email
            [string[]]$Bcc,
        [Parameter(Mandatory = $false)] 
            [switch] $Importance,
        [Parameter(Mandatory = $false)] 
            [switch]$HTMLBody,
        [Parameter(Mandatory = $false)]
            [string[]]$Attachment
    )
    $headers = @{
        Authorization = "Bearer $($token.access_token)"
        "Content-Type" = "application/json"
    }
    $mailBody = @{ # Build MailBody
        message         = @{
            subject      = $Subject
            body         = @{
                contentType = $HTMLBody ? "HTML" : "Text" # set to HTML if true otherwise text
                content     = $Body
            }
        }
        saveToSentItems = $true
    }
    $mailBody.message.toRecipients = @( # Add To Data to EmailBody
        $To | ForEach-Object {
            @{ emailAddress = @{ address = $_ } }
        }
    )
    if($Cc){
        $mailBody.message.ccRecipients = @( # Optional, Add Cc Data to EmailBody 
            $Cc | ForEach-Object {
                @{ emailAddress = @{ address = $_ } }
            }
        )
    }
    if($Bcc){
        $mailBody.message.bccRecipients = @( # Optional, Add Bcc Data to EmailBody
            $Bcc | ForEach-Object {
                @{ emailAddress = @{ address = $_ } }
            }
        )
    }
    if($Importance){ # set email as high Importance
        $mailBody.message.Importance = "high"
    }
    if ($Attachment){  # add attachments
        $mailBody.message.attachments = @(
            foreach($file in $attachment){
                if (Test-Path $file) {
                    @{
                        "@odata.type" = "#microsoft.graph.fileAttachment"  
                        name          = [IO.Path]::GetFileName($file)   
                        contentType   = "application/octet-stream"       
                        contentBytes  = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file))  
                    }
                }
            }
        )
    }

    $mailBody = $mailBody | convertTo-Json -Depth 10 # Convert the Body to JSON for API
    $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"

    $returnObject = [pscustomobject]@{ # data to return
        Status      = "" # Fills within try/catch
        Subject     = $Subject
        ContentType = $HTMLBody ? "HTML" : "Text"
        From        = $From
        To          = $To
        Cc          = $Cc
        Bcc         = $Bcc
        Importance  = [bool]$Importance
    }
    try {
        Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $mailBody -ErrorAction Stop | Out-Null
        $returnObject.Status = "SUCCESS: Sent Email"
        return $returnObject
    }
    catch {
        $returnObject.Status = "Failed: Unable to Send Email: $(($_ | convertFrom-json).error.message)"
        Write-Error "$($returnObject | ConvertTo-Json)"
        return 
    }
}






# Make HTML Table + Body
Connect-AzAccount
$resources = Get-AzResource | select-object name,ResourceGroupName,Location
$HTMLTable = New-SimpleHtmlTable -object $resources
$HTMLBody = "<b> Below are the list of resources: </b><br><br> $HTMLTable"

# Get Token
$kvName         = ""
$tenantID       = Get-AzKeyVaultSecret -VaultName $kvName -Name "tenantID" -AsPlainText 
$clientName     = ""
$clientID       = ""
$clientSecret   = Get-AzKeyVaultSecret -VaultName $kvName -Name $clientName -AsPlainText
$token = New-MsGraphOauthToken -ClientId $clientID -ClientSecret $clientSecret -TenantID $tenantID 

# Send Email
$params = @{
    Token       = $token
    From        = "<email>"
    To          = "<email>"
    Cc          = "<email>", "<email>"
    Subject     = "<subjectName>"
    Body        = $HTMLBody
    HTMLBody    = $true
}
try {
    Send-MsGraphMail @params -ErrorAction Stop
}
catch {
    Write-Error "$_"
    throw
}
