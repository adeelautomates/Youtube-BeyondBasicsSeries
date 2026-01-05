# here are the cmdlets I ran while making this video

# connect-azaccount # connect as myself

# authN to Graph with site.readwrite.all
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
$token = (Invoke-RestMethod -Method POST -Uri $uri -Body $body).access_token

$headers = @{
  Authorization = "Bearer $token"
}

# search all sites (including onedrive)
$uri = "https://graph.microsoft.com/v1.0/sites/getAllSites"
$response = Invoke-RestMethod -Method GET -uri $uri -Headers $headers
$response.value | select-object name, webUrl, ID, isPersonalSite | fl

# search all sharepoint sites
$uri = "https://graph.microsoft.com/v1.0/sites?search=*"
$response = Invoke-RestMethod -Method GET -uri $uri -Headers $headers
$response.value

# search by URL by putting it in the URI # "https://tenant.sharepoint.com/sites/MyTeamSite"
$uri = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/MyTeamSite"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response

# search by display name
$displayName = [uri]::EscapeDataString("My Team Site")
$uri = "https://graph.microsoft.com/v1.0/sites?search=$displayName"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value


#-------------------------
# get site ID and make prefix
$headers = @{
  Authorization = "Bearer $Token"
}
$uri = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/MyTeamSite"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$siteID = $response.id
$sitePrefix = "https://graph.microsoft.com/v1.0/sites/$siteID"

# get Shared Document/Drive ID
$uri = "$sitePrefix/drives"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$driveID = $response.value.id

# get root folders
$uri = "$sitePrefix/drives/$($driveID)/root/children" + "?`$select=id,name,webUrl,folder"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value | where-object folder -ne $null

# get specific folder called general
$folderID = ($response.value | where-object name -eq "General").id

# get general folders items
$uri = "$sitePrefix/drives/$($driveID)/items/$($folderID)/children" + "?`$select=id,name,webUrl,folder"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value

# get MyFiles items
$folderID = ($response.value | Where-Object { $_.name -eq "MyFiles" }).id
$uri = "$sitePrefix/drives/$($driveID)/items/$($folderID)/children" + "?`$select=id,name,webUrl,folder"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value


# Navigating using path
$folderPath = "root:/General/MyFiles"
$uri = "$sitePrefix/drives/$($driveID)/$($folderPath):/children"  + "?`$select=id,name,webUrl,folder"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value


# Create folder
$headers = @{
 	 Authorization = "Bearer $Token"
	"Content-Type" = "application/json"
}
$folderPath = "root:/General/MyFiles"
$folderName = "MyCreatedFolder"
$uri = "$sitePrefix/drives/$($driveID)/$($folderPath):/children"
$body = @{
    name = $folderName
    folder = @{}
    '@microsoft.graph.conflictBehavior' = "rename"
} | ConvertTo-Json -Depth 5
$response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body
$response

# Grant Permissions to folder
$groupID = "00000-000-000-00000"
$folderID = $response.id
$uri = "$sitePrefix/drives/$driveID/items/$folderID/invite"
$body = @{
    recipients = @(@{ objectId = $groupID })
    roles          = @("write")       # edit. can make owners and readers
    requireSignIn  = $true 			# Only authenticated users can access it
    sendInvitation = $false           # no email, no “invite”
} | ConvertTo-Json -Depth 10
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


# Get Info on items of a folder
$headers = @{
 	 Authorization = "Bearer $Token"
}
$folderPath = "General/MyFiles"   # inside Shared Documents
$uri = "$sitePrefix/drives/$driveID/root:/$($folderPath):/children" + "?`$select=id,name,webUrl"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value | format-list

# read text file
$fileID = ($response.value | Where-Object { $_.name -eq "HELLO WORLD.txt" }).id
$uri = "$sitePrefix/drives/$($driveID)/items/$($fileID)/content"
Invoke-RestMethod -Method GET -Uri $uri -Headers $headers

# read csv file
$fileID = ($response.value | Where-Object { $_.name -eq "HELLO WORLD CSV.csv" }).id
$uri = "$sitePrefix/drives/$($driveID)/items/$($fileID)/content"
Invoke-RestMethod -Method GET -Uri $uri -Headers $headers | convertFrom-Csv

# download a file
$folderPath = "General/MyFiles"   # inside Shared Documents
$uri = "$sitePrefix/drives/$driveID/root:/$($folderPath):/children"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value | format-list
$file = $response.value | Where-Object { $_.name -eq "MyDocFile.docx" }
invoke-restmethod -Method Get -uri $($file.'@microsoft.graph.downloadUrl') -OutFile "c:\test\$($file.name)"

# upload a file
$fileName = "MyUploadFile.csv"
get-service | select-object -first 5 | export-csv "c:\test\$($fileName)" -NoTypeInformation

$folderPath = "General"   # inside Shared Documents
$uri = "$sitePrefix/drives/$driveID/root:/$($folderPath):/children"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$folderID = ($response.value | where-object name -eq "MyFiles").id

$uri = "$sitePrefix/drives/$driveID/items/$($folderID):/$($fileName):/content"
$response = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -InFile "c:\test\$fileName"


# upload a file without infile
$fileName = "MyUploadFile2.csv"
get-service | select-object -first 5 | Export-Csv "C:\Test\$($fileName)" -NoTypeInformation
$bytes = [System.IO.File]::ReadAllBytes("C:\Test\$($fileName)")
$uploadHeaders = @{
    Authorization = "Bearer $Token"
    "Content-Type" = "application/octet-stream"
}
$uri = "$sitePrefix/drives/$driveID/items/$($folderID):/$($fileName):/content"
$response = Invoke-RestMethod -Method PUT -Uri $uri -Headers $uploadHeaders -body $bytes


#---------------------
# SharePoint Lists
#---------------------
# get lists
$uri = "$sitePrefix/lists"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value | select-object name, webUrl, id

# create lists
$headers = @{
 	 Authorization = "Bearer $Token"
	"Content-Type" = "application/json"
}
$listName = "CharacterSheet"
$uri = "$sitePrefix/lists"
$body = @{
    displayName = $listName
    columns = @(
        @{
            name = "RealName"
            text = @{}
        }
        @{
            name = "SuperPowers"
            text = @{ allowMultipleLines = $true }
        }
        @{
            name = "Alignment"
            choice = @{
                choices         = @("Hero", "Villian", "Antihero", "Neutral")
                displayAs       = "dropDownMenu"
                allowTextEntry  = $false
            }
        }
        @{
            name = "PowerLevel"
            number = @{ minimum = 1; maximum = 10; decimalPlaces = "none"}
        }
        @{
            name = "IsFavourite"
            boolean = @{}
        }
        @{
            name = "DateOfBirth"
            dateTime = @{
                format = "dateOnly"
                displayAs = "default"
            }
        }
    )
    list = @{
        template = "genericList"
    }
} | convertTo-Json -Depth 10
$response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body
$listID = $response.id

# Create Entries
$uri = "$sitePrefix/lists/$listID/items"
$body = @{
  fields = @{
    Title       = "Superman"
    RealName    = "Clark Kent"
    SuperPowers = "Super strength; flight; heat vision; x-ray vision; super speed; invulnerability; super hearing"
    Alignment   = "Hero"
    PowerLevel  = 10
    IsFavourite = $false
    DateOfBirth = "1938-04-18"
  }
} | ConvertTo-Json -Depth 10
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


$uri = "$sitePrefix/lists/$listID/items"
$body = @{
  fields = @{
    Title       = "Batman"
    RealName    = "Bruce Wayne"
    SuperPowers = "Peak human conditioning; martial arts mastery; detective skills; gadgets; tactical"
    Alignment   = "Hero"
    PowerLevel  = 7
    IsFavourite = $true
    DateOfBirth = "1939-05-01"
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body

$uri = "$sitePrefix/lists/$listID/items"
$body = @{
  fields = @{
    Title       = "Joker"
    RealName    = "Unknown"
    SuperPowers = "Master strategist; psychological manipulation; chemistry/toxin expertise; unpredictable tactics"
    Alignment   = "Villain"
    PowerLevel  = 7
    IsFavourite = $false
    DateOfBirth = "1970-01-01"
  }
} | ConvertTo-Json -Depth 10
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


$uri = "$sitePrefix/lists/$listID/items"
$body = @{
  fields = @{
    Title       = "Deadpool"
    RealName    = "Wade Wilson"
    SuperPowers = "Regeneration; pain tolerance; expert marksman; swordsmenship; weird banter"
    Alignment   = "Antihero"
    PowerLevel  = 8
    IsFavourite = $true
    DateOfBirth = "1971-09-22"
  }
} | ConvertTo-Json -Depth 10
Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


$uri = "$sitePrefix/lists/$listID/items"
$body = @{
  fields = @{
    Title       = "Wonderbread Woman" # wrong
    RealName    = "Princess Diana" # wrong
    SuperPowers = "Summons a Loaf!" # wrong
    Alignment   = "Hero"
    PowerLevel  = 1
    IsFavourite = $true
    DateOfBirth = "1984-10-10"
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body


# getting our list's fields
$uri = "$sitePrefix/lists/$listID/items?`$expand=fields(`$select=id,Title,RealName,SuperPowers,Alignment,PowerLevel,DateOfBirth)"
$response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
$response.value.fields

# modifying existing fields
$itemID = ($response.value.fields | where-object Title -eq "Wonderbread Woman").id
$uri = "$sitePrefix/lists/$listID/items/$itemID/fields"
$body = @{
    Title       = "Wonder Woman"
    RealName    = "Diana Prince"
    SuperPowers = "Super strength; combat skill; lasso of truth; enhanced speed and durability"
    PowerLevel  = 9
} | ConvertTo-Json -Depth 10
Invoke-RestMethod -Method PATCH -Uri $uri -Headers $headers -Body $body

$uri = "$sitePrefix/lists/$listID/items/$itemID"
Invoke-RestMethod -Method DELETE -Uri $uri -Headers $headers


# test site.selected
$clientID = "00000-000-000-00000"
$clientSecret = Get-AzKeyVaultSecret -VaultName "example-kv-001" -Name "example-graph-sharepoint-rw-myteamsite" -AsPlainText 
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

$siteID = 'tenant.sharepoint.com,00000-000-000-00000,00000-000-000-00000'
$headers = @{
  Authorization = "Bearer $Token"
}
$uri = "https://graph.microsoft.com/v1.0/sites/$siteID"
invoke-RestMethod -Method GET -Uri $uri -Headers $headers

$siteID = 'tenant.sharepoint.com,00000-000-000-00000,00000-000-000-00000'
$headers = @{
  Authorization = "Bearer $Token"
}
$uri = "https://graph.microsoft.com/v1.0/sites/$siteID"
Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
