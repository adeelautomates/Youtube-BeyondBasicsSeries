<#
.SYNOPSIS
Sends an email using Microsoft Graph API with optional CC, BCC, Attachments and HTML Body.

.DESCRIPTION
The Send-CorpoGraphMail function sends mail through Microsoft Graph using an authenticated token.
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
Send-CorpoGraphMail -Token $token -From "corpo-automation@lb4s.onmicrosoft.com" `
-To "user@domain.com" -Subject "Test Email" -Body "This is a test."

.EXAMPLE
$params = @{
    Token = $token
    Subject = "Test Email"
    From = "<emailAddress>"
    To = "<emailAddress>"
    Cc = "<emailAddress>", "<emailAddress>"
    Body = "<b>HELLO BOLD WORDS. THIS IS HTML!</b>"
    HTMLBody = $true
    Importance = $true
    ErrorAction = 'Stop'
    Attachment = "C:\test\TestFile1.txt", "C:\test\TestFile2.txt"
}
try {
    Send-CorpoGraphMail @params 
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
function Send-CorpoGraphMail{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)] 
            [pscustomobject]$Token,
        [Parameter(Mandatory = $true)] 
            [string]$Subject,
        [Parameter(Mandatory = $true)] 
            [string]$Body,
        [Parameter(Mandatory = $true)] 
            [ValidateSet('<emailAddress>', '<emailAddress>')]
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