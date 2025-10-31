# Script exactly likw how it was at the end.
# We will improve it in a future episode for V2

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
    throw "Failed to authenticate: $($_.Exception.Message)"
}

#----------------------
# Function
#----------------------
function Generate_htmlTable {
    param (
        $object
    )

    $columns = $object[0].PSObject.Properties.Name
    $htmlTHData = ""
    $htmlTHData = ($columns | ForEach-Object {"        <th>$_</th>"}) -join "`n"

$htmlTableColumns = @"
    <tr align="left" style="background-color: #e7e7e7; color: #000000; padding: 10px; font-size: 16px">
$htmlTHData
    </tr>
"@

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


$htmlTable = @"
<table cellpadding="10" cellspacing="0" border="0" style="border-collapse: collapse; width:auto; display:inline-table; font-size: 14px; font-family: Calibri, Arial, sans-serif;">
$htmlTableColumns
$htmlTableRows
</table>
"@

    return $htmlTable

}

#--------------------------------------
# Collect data about secrets expiration
#--------------------------------------
# AuthN
$kvName = "<kvName>"
$tenantID = Get-AzKeyVaultSecret -VaultName $kvName -Name "tenantID" -AsPlainText
$clientName = "<appRegistrationName>"
$clientID = "<appRegistrationClientID>"
$clientSecret = Get-AzKeyVaultSecret -VaultName $kvName -Name $clientName -AsPlainText
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$uri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
$connection = Invoke-RestMethod -Uri $uri -Method POST -Body $body 
$token = $connection.access_token | ConvertTo-SecureString -AsPlainText -Force 
Connect-MgGraph -AccessToken $token | Out-Null

# Collect App Data

$apps = get-mgapplication -all
$appDataCollected = New-Object System.Collections.Generic.List[pscustomobject]
foreach ($app in $apps){
    if ([array]$ownersUPN = (Get-MgApplicationOwner -ApplicationId $app.Id).AdditionalProperties.userPrincipalName){}
    $secretExpiry = @{}
    foreach ($secret in $App.PasswordCredentials){
        if ($secret) { $secretExpiry[$secret.DisplayName ?? 'none'] = $secret.EndDateTime } 
    }
    $appDataCollected.Add([pscustomobject]@{
        ApplicationName = $app.DisplayName
        ApplicationID = $app.AppID
        Owners = $ownersUPN
        Secrets = $secretExpiry  
    })
}


#--------------------
# Generate HTML table
#--------------------
$Today = Get-date
$expiryCheck = 3
$allExpiryTableData = New-Object System.Collections.Generic.List[pscustomobject]

foreach($data in $appDataCollected){
    if( !($data.secrets.values | Where-Object { ($_ -as [datetime]) -lt $today.AddMonths($expiryCheck) })){continue} # if secrets dont match condition, skip.
    [array]$secretsOutput = @() # store secrets for outputs
    foreach($secret in $data.Secrets.GetEnumerator()){
        if ($secret.value | where-object { ($_ -as [datetime]) -lt $Today }){
            $secretsOutput += "<b><span style='color: #960000;'>EXPIRED:</span></b> [$($secret.Key)] <b>$($secret.Value.ToString('yyyy/MM/dd'))</b>"
        }
        else{
            $secretsOutput += "<b><span style='color: #ac7e00;'>EXPIRING:</span></b> [$($secret.Key)] <b>$($secret.Value.ToString('yyyy/MM/dd'))</b>"
        }
    }
    $allExpiryTableData.Add([pscustomobject]@{
        Name        = "<a href='https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Credentials/appId/$($data.ApplicationID)/isMSAApp~/false'>$($data.ApplicationName)</a>"
        Owners      = $data.owners -join '<br>'
        Type        = "Secret"
        Expiration  = $secretsOutput -join '<br>'
    })
}

$htmlTable = Generate_htmlTable -object $allExpiryTableData
# $htmlTable | out-file "C:\vsc\Local\HTMLTableSecrets.html"


# Send mail as "sp-it-graph-mail-send" using Graph API
$tenantId = Get-AzKeyVaultSecret -VaultName $kvName -Name "tenantID" -AsPlainText
$clientName = "<appRegistrationName>"
$clientID = "<appRegistrationClientID>"
$ClientSecret = Get-AzKeyVaultSecret -VaultName "corpo-mgmt-kv-001" -Name $clientName -AsPlainText
$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $ClientId
    Client_Secret = $ClientSecret
}
$connection = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $body


$openHTML = @"
Hello IT Team,<br><br>

These are the secrets detected in our tenant that are about to expire:<br>
</p>

<ul>
  <li><b><span style='color: #960000;'>EXPIRED</span></b> means they have expired on the date indicated.</li>
  <li><b><span style='color: #ac7e00;'>EXPIRING</span></b> means they are about to expire on the date indicated.</li>
</ul>

<p>
The owners will be informed of their secrets expiring. However if owners are missing... <b>please take action</b>.
</p>
"@

$closeHTML = @"
<p>
<i>This email was automatically generated by: <b>report_entraID_secretExpiry.ps1</b><i>
</p>
"@

$emailHTML = $openHTML + $htmlTable + $closeHTML

$headers = @{
    Authorization = "Bearer $($connection.access_token)"
    "Content-Type" = "application/json"
}
$mailBody = @{
    message         = @{
        subject      = "Test Sending Email as App Registration"
        body         = @{
            contentType = "HTML"
            content     = $emailHTML
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = "<emailOfRecipient>"
                }
            }
        )
    }
    saveToSentItems = $true
} | convertTo-Json -Depth 10

$userID = "<emailOfSender>"
$uri = "https://graph.microsoft.com/v1.0/users/$userId/sendMail"

Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $mailBody



$uniqueOwners = $allExpiryTableData.Owners -split '<br>' | where-object { $_.Trim() } | Select-Object -Unique
foreach($owner in $uniqueOwners){
    $user = get-mguser -userId $owner -property DisplayName, userprincipalName
    $userExpiryTableData = $allExpiryTableData | where-object {$_.owners -like "*$owner*"}
    $htmlTable = Generate_htmlTable -object $userExpiryTableData

    $openHTML = @"
<p style="color: #d70000;"><b><i>THIS EMAIL IS AUTOGENERATED, DO NOT REPLY.</i></b></p>

Hello $($user.DisplayName),<br><br>

Your secrets for your App Registrations have been detected as expired or about to be expired. Please take action<br>
</p>
<ul>
  <li><b><span style='color: #960000;'>EXPIRED</span></b> means they have expired on the date indicated.</li>
  <li><b><span style='color: #ac7e00;'>EXPIRING</span></b> means they are about to expire on the date indicated.</li>
</ul>
"@

    $closeHTML = @"
<p>
<i>This email was automatically generated by: <b>report_entraID_secretExpiry.ps1</b><i><br>
If you have any questions. Please reach out to the IT Department.
</p>
"@

    $emailHTML = $openHTML + $htmlTable + $closeHTML

    $headers = @{
        Authorization = "Bearer $($connection.access_token)"
        "Content-Type" = "application/json"
    }
    $mailBody = @{
        message         = @{
            subject      = "Test Sending Email as App Registration"
            body         = @{
                contentType = "HTML"
                content     = $emailHTML
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $owner
                    }
                }
            )
        }
        saveToSentItems = $true
    } | convertTo-Json -Depth 10

    $userID = "<emailOfSender>"
    $uri = "https://graph.microsoft.com/v1.0/users/$userId/sendMail"
    Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $mailBody
}