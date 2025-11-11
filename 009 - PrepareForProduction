<#
    NAME: report_entraID_secretExpiry
    CREATOR: Adeel A.
    DESCRIPTION: 
        - Collects Information on App Registrations Secrets Expiring
        - Sends Alerts to Owners of the apps + 
#>

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
    write-error "Failed to authenticate: $($_.Exception.Message)"
    throw
}

#----------
# Variables
#----------
$kvName = "" # key vault holding secrets
$tenantID = Get-AzKeyVaultSecret -VaultName $kvName -Name "tenantID" -AsPlainText
$expiredColor = "#960000" # color for expired secrets in emails
$expiringColor  = "#ac7e00" # color for expiring secrets in emails
$senderEmail = "<email>" # used to send emails out
$itStaffEmail = "<email>" # email to send full reports to
$dateNow = Get-date # get today's date and store it
$expiryCheck = 3 # how far in the future to report secrets that are about to expire (months)

#----------------------
# Function
#----------------------
# Function To Get OAuth Token
function New-GraphOAuthToken {
    param (
        $clientId,
        $clientSecret,
        $tenantID
    )
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
    $uri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
    $token = Invoke-RestMethod -Uri $uri -Method POST -Body $body
    return $token
}

# Function to Build HTML Table
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

$clientName = "sp-it-graph-identity-rw-all"
$clientID = "8699fb98-1d1d-4641-a3e1-d39969786b17"
$clientSecret = Get-AzKeyVaultSecret -VaultName $kvName -Name $clientName -AsPlainText
$connection = New-GraphOAuthToken -clientId $clientID -clientSecret $clientSecret -tenantID $tenantID
$token = $connection.access_token | ConvertTo-SecureString -AsPlainText -Force 
try {
    Connect-MgGraph -AccessToken $token -ErrorAction Stop | Out-Null
    Write-Output "Connected to Graph: $((get-mgcontext).AppName)"
}
catch {
    write-error "Failed to authenticate to Graph as clientID[$clientID]: $($_.Exception.Message)"
    throw
}

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
write-output "Collected App Registration data: ID, Name, Secrets & Owners for all apps."

$allExpiryTableData = New-Object System.Collections.Generic.List[pscustomobject]
foreach($data in $appDataCollected){
    if( !($data.secrets.values | Where-Object { ($_ -as [datetime]) -lt $dateNow.AddMonths($expiryCheck) })){continue} # if secrets dont match condition, skip.
    [array]$secretsOutput = @() # store secrets for outputs
    foreach($secret in $data.Secrets.GetEnumerator()){
        if ($secret.value | where-object { ($_ -as [datetime]) -lt $dateNow }){
            $secretsOutput += "<b><span style='color: $expiredColor;'>EXPIRED:</span></b> [$($secret.Key)] <b>$($secret.Value.ToString('yyyy/MM/dd'))</b>"
        }
        else{
            $secretsOutput += "<b><span style='color: $expiringColor;'>EXPIRING:</span></b> [$($secret.Key)] <b>$($secret.Value.ToString('yyyy/MM/dd'))</b>"
        }
    }
    $allExpiryTableData.Add([pscustomobject]@{
        Name        = "<a href='https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Credentials/appId/$($data.ApplicationID)/isMSAApp~/false'>$($data.ApplicationName)</a>"
        Owners      = $data.owners -join '<br>'
        Type        = "Secret"
        Expiration  = $secretsOutput -join '<br>'
    })
}
if ($allExpiryTableData.count -eq 0){
    return "No Secrets expiring within $expiryCheck month(s). Ending Script"
}
write-output "Collected secrets that are expiring(or expired) that are within $expiryCheck month(s). Proceeding."


#----------------------------------
# Generate HTML table & Send Emails
#----------------------------------
# Generate table for IT Staff
$htmlTable = Generate_htmlTable -object $allExpiryTableData
$emailHTML = @"
Hello IT Team,<br><br>

These are the secrets detected in our tenant that are about to expire:<br>
</p>

<ul>
  <li><b><span style='color: $expiredColor;'>EXPIRED</span></b> means they have expired on the date indicated.</li>
  <li><b><span style='color: $expiringColor;'>EXPIRING</span></b> means they are about to expire on the date indicated.</li>
</ul>

<p>
The owners will be informed of their secrets expiring. However if owners are missing... <b>please take action</b>.
</p>

$htmlTable

<p>
<i>This email was automatically generated by: <b>report_entraID_secretExpiry.ps1</b><i>
</p>
"@

# Send mail as "sp-it-graph-mail-send" using Graph API
write-output "Generating email to send to IT Staff"
$clientId = ""
$clientName = "sp-it-graph-mail-send"
$clientSecret = Get-AzKeyVaultSecret -VaultName $kvName -Name $clientName -AsPlainText
$connection = New-GraphOAuthToken -clientId $clientID -clientSecret $clientSecret -tenantID $tenantID
$headers = @{
    Authorization = "Bearer $($connection.access_token)"
    "Content-Type" = "application/json"
}
$mailBody = @{
    message         = @{
        subject      = "Report - EntraID App Secret Expiration [$($DateNow.ToString('yyyy/MM/dd'))]"
        body         = @{
            contentType = "HTML"
            content     = $emailHTML
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = "<email>"
                }
            }
        )
    }
    saveToSentItems = $true
} | convertTo-Json -Depth 10
$uri = "https://graph.microsoft.com/v1.0/users/$senderEmail/sendMail"

try {
    Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $mailBody -errorAction Stop | Out-Null
    write-output "Full Email Report Sent: $itStaffEmail"
}
catch {
    Write-Error "Failed to send email to $($itStaffEmail): $(($_ | convertFrom-json).error.message)"
}


write-output "Preparing to send emails to all owners..."

$uniqueOwners = $allExpiryTableData.Owners -split '<br>' | where-object { $_.Trim() } | Select-Object -Unique
foreach($owner in $uniqueOwners){
    $user = get-mguser -userId $owner -property DisplayName
    $userExpiryTableData = $allExpiryTableData | where-object {$_.owners -like "*$owner*"}
    $htmlTable = Generate_htmlTable -object $userExpiryTableData

    $emailHTML = @"
<p style="color: #d70000;"><b><i>THIS EMAIL IS AUTOGENERATED, DO NOT REPLY.</i></b></p>

Hello $($user.DisplayName),<br><br>

Your secrets for your App Registrations have been detected as expired or about to be expired. Please take action<br>
</p>
<ul>
  <li><b><span style='color: $expiredColor;'>EXPIRED</span></b> means they have expired on the date indicated.</li>
  <li><b><span style='color: $expiringColor;'>EXPIRING</span></b> means they are about to expire on the date indicated.</li>
</ul>

$htmlTable

<p>
<i>This email was automatically generated by: <b>report_entraID_secretExpiry.ps1</b><i><br>
If you have any questions. Please reach out to the IT Department.
</p>
"@
    $headers = @{
        Authorization = "Bearer $($connection.access_token)"
        "Content-Type" = "application/json"
    }
    $mailBody = @{
        message         = @{
            subject      = "ALERT - Your Secrets in EntraID are Expiring [$($DateNow.ToString('yyyy/MM/dd'))]"
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

    $userID = "corpo-automation@lb4s.onmicrosoft.com"
    $uri = "https://graph.microsoft.com/v1.0/users/$senderEmail/sendMail"
    try{
        Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $mailBody -ErrorAction Stop | Out-Null
        write-output "Sent Email to $($user.DisplayName)"
    }
    catch {
        Write-Error "Failed to send email to $($owner): $(($_ | convertFrom-json).error.message)"
    }
}